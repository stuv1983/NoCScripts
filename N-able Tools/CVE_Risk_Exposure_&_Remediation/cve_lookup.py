"""
cve_lookup.py — fetch fixed version data from public CVE APIs and
                automatically populate config.json fixed_version_rules.

Sources (tried in order, first success wins):
    1. CVE.org JSON API   — cveawg.mitre.org/api/cve/{id}
                            Uses CVE JSON 5.0 schema — most authoritative source
    2. NVD API 2.0        — services.nvd.nist.gov/rest/json/cves/2.0
                            Returns CPE match criteria with version ranges
    3. OSV.dev API        — api.osv.dev/v1/vulns/{id}
                            Best for open-source packages (Chrome, Firefox etc.)

Run manually:
    python cve_lookup.py CVE-2026-5288 CVE-2026-5289
    python cve_lookup.py --auto           # fetch all CVEs in config.json rules
    python cve_lookup.py --dry-run CVE-2026-5288
    python cve_lookup.py --test-nvd-api

NVD API key support:
    Preferred: set environment variable NVD_API_KEY
    Optional:  config.json -> { "api": { "nvd_api_key": "..." } }

Or from code:
    from cve_lookup import lookup_fixed_version, enrich_config

What it does
------------
For each CVE it finds the minimum fixed version (the first version that
contains the fix) and writes it into config.json fixed_version_rules under
the matching canonical product key.

The product key is determined by matching the CVE's vendor/product string
against config.json product_map — the same matching logic used by the pipeline.

Example config.json update:
    "chrome": {
        "CVE-2026-5288": "136.0.7103.116",   ← added automatically
        "_baseline": "148.0.7778.97"
    }
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
import urllib.request
import urllib.error
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_UA      = 'Mozilla/5.0 N-able-CVE-Dashboard/1.0 (automated; contact your-email@example.com)'
_TIMEOUT = 12
_RETRY   = 2
_DELAY   = 0.5   # seconds between API calls — be a polite client


def _mask_key(value: str) -> str:
    """Return a safe display form for secrets, e.g. ****ABCD."""
    value = str(value or '').strip()
    if not value:
        return ''
    return '****' + value[-4:] if len(value) > 4 else '****'


def _load_config(config_path: Optional[str] = None) -> dict:
    """Best-effort config.json loader used only for optional API settings."""
    try:
        if config_path is None:
            config_path = Path(__file__).parent / 'config.json'
        config_path = Path(config_path)
        if not config_path.exists():
            return {}
        with open(config_path, 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception as exc:
        log.debug("cve_lookup: could not read config for API settings: %s", exc)
        return {}


def _get_nvd_api_key(config_path: Optional[str] = None) -> tuple[str, str]:
    """
    Return (api_key, source).

    Priority:
      1. NVD_API_KEY environment variable
      2. config.json -> api.nvd_api_key

    The key is optional. Blank means unauthenticated NVD requests.
    """
    env_key = os.getenv('NVD_API_KEY', '').strip()
    if env_key:
        return env_key, 'environment variable NVD_API_KEY'

    cfg_key = str(_load_config(config_path).get('api', {}).get('nvd_api_key', '')).strip()
    if cfg_key:
        return cfg_key, 'config.json api.nvd_api_key'

    return '', 'none'


def _nvd_headers(config_path: Optional[str] = None, log_status: bool = False) -> dict[str, str]:
    """Return the extra headers for NVD requests, including apiKey when configured."""
    key, source = _get_nvd_api_key(config_path)
    if key:
        if log_status:
            log.info("cve_lookup: NVD API key detected from %s (%s)", source, _mask_key(key))
        return {'apiKey': key}

    if log_status:
        log.warning(
            "cve_lookup: no NVD API key configured; using unauthenticated NVD requests "
            "which may hit HTTP 403/429 rate limits"
        )
    return {}


def test_nvd_api_access(config_path: Optional[str] = None, cve_id: str = 'CVE-2024-21413') -> bool:
    """
    Perform a direct NVD API test and log clear pass/fail feedback.

    This is intentionally independent of CVE.org fallback logic so it confirms
    NVD access specifically. It returns True only when NVD returns at least one
    vulnerability record.
    """
    cve_id = cve_id.strip().upper()
    headers = _nvd_headers(config_path, log_status=True)
    url = f'https://services.nvd.nist.gov/rest/json/cves/2.0?cveId={cve_id}'
    log.info("cve_lookup: testing NVD API access with %s", cve_id)
    data = _get(url, extra_headers=headers)
    count = len(data.get('vulnerabilities', [])) if isinstance(data, dict) else 0
    if count:
        log.info("cve_lookup: NVD API access OK — received %d vulnerability record(s) for %s", count, cve_id)
        return True
    log.error("cve_lookup: NVD API access test failed — no vulnerability records returned for %s", cve_id)
    return False


# ==============================================================================
# API FETCHERS
# ==============================================================================

def _get(url: str, extra_headers: Optional[dict[str, str]] = None) -> Optional[dict]:
    """HTTP GET with retries. Returns parsed JSON or None on failure."""
    headers = {'User-Agent': _UA, 'Accept': 'application/json'}
    if extra_headers:
        headers.update(extra_headers)

    for attempt in range(_RETRY):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=_TIMEOUT) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            if e.code == 404:
                log.debug("HTTP 404 from %s — CVE not found", url[:80])
                return None          # CVE not found — not a retry candidate
            if e.code in (403, 429):
                log.warning("Rate limited or blocked by %s — HTTP %d; waiting 5s", url[:80], e.code)
                time.sleep(5)
            else:
                log.debug("HTTP %d from %s (attempt %d)", e.code, url[:80], attempt + 1)
        except Exception as e:
            log.debug("Fetch failed %s: %s (attempt %d)", url[:80], e, attempt + 1)
        time.sleep(_DELAY)
    return None


def _parse_cve_org(data: dict) -> list[dict]:
    """
    Extract affected version ranges from CVE.org JSON 5.0 schema.

    Returns list of:
        { vendor, product, fixed_version }
    """
    results = []
    try:
        cna = data.get('containers', {}).get('cna', {})
        for affected in cna.get('affected', []):
            vendor  = str(affected.get('vendor', '')).strip()
            product = str(affected.get('product', '')).strip()
            for v in affected.get('versions', []):
                if v.get('status') != 'affected':
                    continue
                # lessThan = first version that is NOT affected (the fix)
                fixed = v.get('lessThan') or v.get('lessThanOrEqual')
                if fixed:
                    results.append({
                        'vendor':        vendor,
                        'product':       product,
                        'fixed_version': str(fixed).strip(),
                    })
    except Exception as e:
        log.debug("CVE.org parse error: %s", e)
    return results


def _parse_nvd(data: dict) -> list[dict]:
    """
    Extract version ranges from NVD API 2.0 CPE match criteria.

    Returns list of:
        { vendor, product, fixed_version }
    """
    results = []
    try:
        vulns = data.get('vulnerabilities', [])
        if not vulns:
            return results
        cve_data = vulns[0].get('cve', {})
        for config in cve_data.get('configurations', []):
            for node in config.get('nodes', []):
                for cpe in node.get('cpeMatch', []):
                    if not cpe.get('vulnerable'):
                        continue
                    fixed = (cpe.get('versionEndExcluding')
                             or cpe.get('versionEndIncluding'))
                    if not fixed:
                        continue
                    # CPE URI: cpe:2.3:a:vendor:product:...
                    cpe_uri = cpe.get('criteria', '')
                    parts   = cpe_uri.split(':')
                    vendor  = parts[3] if len(parts) > 3 else ''
                    product = parts[4] if len(parts) > 4 else ''
                    results.append({
                        'vendor':        vendor,
                        'product':       product,
                        'fixed_version': str(fixed).strip(),
                    })
    except Exception as e:
        log.debug("NVD parse error: %s", e)
    return results


def _parse_osv(data: dict) -> list[dict]:
    """
    Extract fixed versions from OSV.dev API response.
    OSV uses explicit 'fixed' events in version ranges.
    """
    results = []
    try:
        for affected in data.get('affected', []):
            pkg     = affected.get('package', {})
            product = pkg.get('name', '')
            vendor  = pkg.get('ecosystem', '')
            for rng in affected.get('ranges', []):
                for event in rng.get('events', []):
                    fixed = event.get('fixed')
                    if fixed:
                        results.append({
                            'vendor':        vendor,
                            'product':       product,
                            'fixed_version': str(fixed).strip(),
                        })
    except Exception as e:
        log.debug("OSV parse error: %s", e)
    return results


# ==============================================================================
# PRODUCT MATCHING
# ==============================================================================

_VERSION_ONLY = re.compile(r'^\d+(?:\.\d+){1,5}$')

def _match_product(vendor: str, product: str,
                   product_map: list[tuple[str, str]]) -> Optional[str]:
    """
    Try to match a vendor+product string against config.json product_map.
    Returns the canonical product key or None.
    """
    combined = f'{vendor} {product}'.lower().strip()
    for key, canonical in product_map:
        if key in combined:
            return canonical
    return None


def _is_valid_version(v: str) -> bool:
    """Basic sanity check — reject garbage strings."""
    return bool(v) and bool(_VERSION_ONLY.match(v.strip()))


# ==============================================================================
# EDGE CHROMIUM CVE RESOLVER
# ==============================================================================
#
# Problem: Chromium CVEs (e.g. CVE-2026-5288/5289/5290) are filed against
# Google Chrome by the CNA.  CVE.org and NVD return a Chrome fixed version
# only.  Microsoft Edge is a Chromium derivative on a different version train,
# so the Chrome fixed version is wrong for Edge.
#
# Solution: when a CVE maps to 'chrome', also attempt to resolve an Edge-
# specific fixed version from Microsoft's own Edge update API.  The two
# results are stored separately — chrome and edge are independent canonical
# keys and never share a fixed version.
#
# Source: edgeupdates.microsoft.com/api/products
# This is the same source used by version_sync.py for baseline tracking.
# It exposes the per-release security CVE list via the "CVEs" field on each
# release, making it the authoritative Microsoft source for Edge/CVE mapping.
#
# Fallback: if the API doesn't list the CVE (Microsoft sometimes omits CVE
# details for Chromium-inherited fixes), we use the Chromium milestone to
# find the corresponding Edge release.  Edge Stable incorporates a Chromium
# milestone on a ~1-week delay; the milestone can be extracted from the
# Chrome fixed version's major number.

_EDGE_PRODUCTS_URL = 'https://edgeupdates.microsoft.com/api/products'
_EDGE_CHROMIUM_MILESTONE_URL = (
    'https://edgeupdates.microsoft.com/api/products?view=enterprise'
)
_EDGE_CACHE: Optional[list] = None   # module-level cache, one fetch per session


def _get_edge_releases() -> list:
    """Fetch and cache the Edge products/releases list."""
    global _EDGE_CACHE
    if _EDGE_CACHE is not None:
        return _EDGE_CACHE
    try:
        data = _get(_EDGE_PRODUCTS_URL)
        if isinstance(data, list):
            _EDGE_CACHE = data
            return _EDGE_CACHE
    except Exception as e:
        log.debug("cve_lookup: Edge releases fetch failed — %s", e)
    _EDGE_CACHE = []
    return _EDGE_CACHE


def _edge_fixed_version_for_cve(cve_id: str) -> Optional[str]:
    """
    Return the earliest Edge Stable version that lists cve_id in its
    security release notes, or None if not found.

    Microsoft's Edge release API includes a 'CVEs' list per release for
    CVEs that Microsoft explicitly credits in their release notes.  This
    covers most Chromium-inherited CVEs but not all — Microsoft sometimes
    omits CVE IDs when they describe a release only as "incorporates
    Chromium security updates".
    """
    releases = _get_edge_releases()
    if not releases:
        return None

    cve_upper = cve_id.strip().upper()

    # Find the Stable channel product
    stable = next((p for p in releases if str(p.get('Product', '')).lower() == 'stable'), None)
    if not stable:
        return None

    # Walk releases sorted oldest-first; return first one that mentions the CVE
    sorted_releases = sorted(
        stable.get('Releases', []),
        key=lambda r: [int(x) for x in str(r.get('ProductVersion', '0')).split('.')
                       if x.isdigit()],
    )

    for rel in sorted_releases:
        rel_cves = [str(c).strip().upper() for c in rel.get('CVEs', [])]
        if cve_upper in rel_cves:
            ver = str(rel.get('ProductVersion', '')).strip()
            if _is_valid_version(ver):
                log.debug("cve_lookup: Edge release notes: %s fixed in Edge %s", cve_id, ver)
                return ver

    return None


def _edge_fixed_version_for_chromium_milestone(chrome_fixed: str) -> Optional[str]:
    """
    Given a Chrome fixed version string (e.g. '146.0.7680.178'), extract the
    Chromium major milestone (146) and return the earliest Edge Stable release
    that is built on that milestone or later.

    Used as a fallback when the Edge release notes don't explicitly list the
    CVE — Edge incorporates each Chromium security update within ~1 week.
    """
    releases = _get_edge_releases()
    if not releases or not chrome_fixed:
        return None

    try:
        chrome_major = int(chrome_fixed.split('.')[0])
    except (ValueError, IndexError):
        return None

    stable = next((p for p in releases if str(p.get('Product', '')).lower() == 'stable'), None)
    if not stable:
        return None

    candidates = []
    for rel in stable.get('Releases', []):
        ver = str(rel.get('ProductVersion', '')).strip()
        if not _is_valid_version(ver):
            continue
        try:
            edge_major = int(ver.split('.')[0])
        except (ValueError, IndexError):
            continue
        if edge_major >= chrome_major:
            candidates.append(ver)

    if not candidates:
        return None

    # Return the lowest qualifying version (earliest safe Edge release)
    candidates.sort(key=lambda v: [int(x) for x in v.split('.') if x.isdigit()])
    result = candidates[0]
    log.debug(
        "cve_lookup: Edge milestone fallback: Chrome major %d → Edge %s",
        chrome_major, result,
    )
    return result


def _resolve_edge_version(cve_id: str, chrome_fixed: Optional[str]) -> Optional[str]:
    """
    Resolve the Edge-specific fixed version for a Chromium CVE.

    Strategy (in order):
      1. Direct CVE lookup in Edge release notes (authoritative)
      2. Chromium milestone mapping from chrome_fixed version (fallback)

    Returns None if no Edge version can be determined.
    """
    # Step 1: direct CVE lookup
    ver = _edge_fixed_version_for_cve(cve_id)
    if ver:
        log.info("cve_lookup: %s → edge=%s (from Edge release notes)", cve_id, ver)
        return ver

    # Step 2: milestone mapping
    if chrome_fixed:
        ver = _edge_fixed_version_for_chromium_milestone(chrome_fixed)
        if ver:
            log.info(
                "cve_lookup: %s → edge=%s (milestone fallback from chrome=%s)",
                cve_id, ver, chrome_fixed,
            )
            return ver

    log.debug("cve_lookup: %s — could not resolve Edge version", cve_id)
    return None


# Canonical keys that represent Chromium-based products where a separate
# Edge resolution should be attempted when Chrome is matched.
_CHROMIUM_PRODUCTS: frozenset[str] = frozenset({'chrome'})



# Noise words to strip when building a canonical slug from a MITRE product name.
# Keeps the slug short and stable across minor wording variations.
_NOISE_WORDS = re.compile(
    r'\b(version|v|update|security|runtime|framework|sdk|client|server|'
    r'x64|x86|32.bit|64.bit|for|and|the|rtm|lts|esr|release|channel)\b',
    re.IGNORECASE,
)
_VER_STRIP  = re.compile(r'\b\d+(?:\.\d+)+\b')           # strip dotted versions e.g. 6.0.25
_YEAR_STRIP = re.compile(r'\b(?:19|20)\d{2}\b')           # strip 4-digit years e.g. 2022
_ARCH_STRIP = re.compile(r'\(x64\)|\(x86\)|\(64.bit\)|\(32.bit\)', re.IGNORECASE)
_SLUG_CHARS = re.compile(r'[^a-z0-9]+')            # keep only alphanum for slug

# Products whose MITRE name differs significantly from how N-able names them.
# Checked before the generic derivation logic — add more as you find them.
_KNOWN_CANONICAL: list[tuple[str, str]] = [
    # (lowercase substring in "vendor product", canonical_key)
    ('microsoft .net',      'dotnet'),
    (r'\.net framework',    'dotnet'),
    (r'\.net core',         'dotnet'),
    (r'\.net runtime',      'dotnet'),
    ('haxx curl',           'curl'),
    (' curl',               'curl'),
    ('madler zlib',         'zlib'),
    (' zlib',               'zlib'),
    ('microsoft visual studio', 'visualstudio'),
    ('winzip',              'winzip'),
    ('7-zip',               '7zip'),
    ('notepad++',           'notepadpp'),
    ('adobe acrobat',       'acrobat'),
    ('adobe reader',        'acrobat'),
]


def _derive_canonical(vendor: str, product: str) -> str:
    """
    Derive a stable canonical key for a MITRE vendor/product pair.

    Priority:
      1. _KNOWN_CANONICAL lookup — handles products whose MITRE naming
         diverges from N-able's naming (e.g. "haxx curl" → "curl").
      2. Generic slug derived from the product name (vendor-prefixed only
         when needed to avoid collisions, e.g. "microsoft_edge" vs "edge").

    Returns a lowercase alphanum+underscore string ≤ 32 chars.
    """
    combined = f'{vendor} {product}'.lower().strip()
    for fragment, canonical in _KNOWN_CANONICAL:
        if re.search(fragment, combined):
            return canonical

    # Generic path: clean the product name into a stable slug
    p = product.lower()
    p = _VER_STRIP.sub('', p)         # drop version numbers
    p = _NOISE_WORDS.sub(' ', p)      # drop noise words
    p = _SLUG_CHARS.sub('_', p)       # slug
    p = p.strip('_')
    # If slug is very short or empty, prefix vendor
    if len(p) < 3:
        v = _SLUG_CHARS.sub('_', vendor.lower()).strip('_')
        p = f'{v}_{p}'.strip('_')
    return p[:32] or 'unknown'


def _derive_search_key(vendor: str, product: str) -> str:
    """
    Derive the product_map search key — a lowercase substring that will be
    found inside both:
      - the CVE export's "Affected products" column  (e.g. "Microsoft .NET Runtime 6.0 (x64)")
      - the patch report's "Patch" column            (e.g. "2026-04 .NET 8.0.26 Security Update…")

    Strip version numbers, architecture tags, and noise words so the key stays
    stable across releases and is broad enough to match N-able's various naming
    conventions.  Prefer shorter keys that match more product name variants.
    """
    p = _ARCH_STRIP.sub(' ', product)     # remove (x64) / (x86) etc.
    p = _VER_STRIP.sub('', p)             # remove 6.0, 8.0.26, …
    p = _YEAR_STRIP.sub('', p)            # remove 2022, 2019, …
    p = _NOISE_WORDS.sub(' ', p)          # remove runtime, framework, version, …
    p = re.sub(r'\s+', ' ', p).strip()

    combined = f'{vendor} {p}'.lower().strip()
    combined = re.sub(r'\s+', ' ', combined)

    # If the vendor name is embedded in the product string already, use
    # product-only key to avoid ugly "microsoft microsoft edge" patterns
    v_low = vendor.lower().strip()
    p_low = p.lower().strip()
    if v_low and v_low in p_low:
        return p_low

    return combined


def _auto_update_product_map(
    vendor: str,
    product: str,
    fixed_version: str,
    cve_id: str,
    cfg: dict,
) -> Optional[str]:
    """
    Derive a (search_key, canonical) pair from MITRE data and, if the search_key
    is not already in product_map, insert it and create a fixed_version_rules
    entry.  Updates cfg in-place.  Returns the canonical key on success, None
    if nothing was added (key already present or derivation failed).

    Entries are inserted BEFORE the first generic "windows" entry so more-
    specific strings are matched first, as required by the top-to-bottom
    product_map matching logic in data_pipeline.py.
    """
    canonical  = _derive_canonical(vendor, product)
    search_key = _derive_search_key(vendor, product)

    if not canonical or not search_key:
        return None

    pm = cfg.get('product_map', [])

    # Already present — nothing to add
    existing_keys = {str(k).lower() for k, _ in pm}
    if search_key in existing_keys:
        return None

    # Find insertion point: before the first "windows" or generic catch-all entry
    insert_at = len(pm)
    for i, (k, _) in enumerate(pm):
        if str(k).lower() in ('windows', 'windows 10', 'windows 11'):
            insert_at = i
            break

    pm.insert(insert_at, [search_key, canonical])
    cfg['product_map'] = pm

    # Ensure fixed_version_rules has an entry for this canonical
    fvr = cfg.setdefault('fixed_version_rules', {})
    if canonical not in fvr:
        fvr[canonical] = {}

    log.info(
        "cve_lookup: auto-added product_map [%r → %r] from %s (%s %s → fixed %s)",
        search_key, canonical, cve_id, vendor, product, fixed_version,
    )
    return canonical


# ==============================================================================
# MAIN LOOKUP
# ==============================================================================

def lookup_fixed_version(
    cve_id: str,
    product_map: list[tuple[str, str]],
    cfg: Optional[dict] = None,
    auto_add_products: bool = False,
    config_path: Optional[str] = None,
) -> dict[str, str]:
    """
    Fetch the fixed version(s) for a CVE from public APIs.

    Returns dict of { canonical_product_key: fixed_version_string }
    e.g. { 'chrome': '136.0.7103.116', 'edge': '136.0.3240.64' }

    Returns empty dict if CVE not found or no version data available.

    auto_add_products
        When True and cfg is provided, any vendor/product in the CVE data
        that has no product_map match is automatically added to cfg's
        product_map and fixed_version_rules (cfg is mutated in place).
        The caller is responsible for persisting cfg to disk.
        Entries are derived via _derive_canonical / _derive_search_key and
        inserted before the generic "windows" catch-all so specificity order
        is preserved.
    """
    cve_id = cve_id.strip().upper()
    log.info("cve_lookup: fetching %s", cve_id)

    results: list[dict] = []

    # Source 1: CVE.org
    data = _get(f'https://cveawg.mitre.org/api/cve/{cve_id}')
    if data:
        results = _parse_cve_org(data)
        if results:
            log.debug("cve_lookup: %s → CVE.org returned %d affected entries", cve_id, len(results))

    # Source 2: NVD (fallback)
    if not results:
        time.sleep(_DELAY)
        data = _get(
            f'https://services.nvd.nist.gov/rest/json/cves/2.0?cveId={cve_id}',
            extra_headers=_nvd_headers(config_path),
        )
        if data:
            results = _parse_nvd(data)
            if results:
                log.info("cve_lookup: %s → NVD returned %d CPE entries", cve_id, len(results))
            else:
                vulns = len(data.get('vulnerabilities', [])) if isinstance(data, dict) else 0
                log.info("cve_lookup: NVD request succeeded for %s but no fixed-version CPE ranges were found (%d record(s))", cve_id, vulns)

    # Source 3: OSV.dev (fallback)
    if not results:
        time.sleep(_DELAY)
        data = _get(f'https://api.osv.dev/v1/vulns/{cve_id}')
        if data:
            results = _parse_osv(data)
            if results:
                log.debug("cve_lookup: %s → OSV returned %d affected entries", cve_id, len(results))

    if not results:
        log.warning("cve_lookup: no version data found for %s", cve_id)
        return {}

    # Match each result against product_map — collect unmatched for auto-add
    matched:   dict[str, str]  = {}
    unmatched: list[dict]      = []

    for r in results:
        fv = r.get('fixed_version', '').strip()
        if not _is_valid_version(fv):
            continue
        pk = _match_product(r.get('vendor', ''), r.get('product', ''), product_map)
        if not pk:
            unmatched.append(r)
            log.debug("cve_lookup: no product_map match for %s %s",
                      r.get('vendor'), r.get('product'))
            continue
        # Keep the lowest fixed version per product (most conservative)
        if pk not in matched or _version_lt(fv, matched[pk]):
            matched[pk] = fv

    # Auto-add: derive and register new product_map entries from unmatched results
    if auto_add_products and cfg is not None and unmatched:
        live_pm = [(str(k).lower(), str(v).lower())
                   for k, v in cfg.get('product_map', [])]
        for r in unmatched:
            vendor  = r.get('vendor', '')
            product = r.get('product', '')
            fv      = r.get('fixed_version', '')

            # Skip if another result already matched this vendor/product
            if _match_product(vendor, product, live_pm):
                continue

            canonical = _auto_update_product_map(vendor, product, fv, cve_id, cfg)
            if canonical:
                # Rebuild live_pm from the now-updated cfg so subsequent
                # results in this loop benefit from the new entry
                live_pm = [(str(k).lower(), str(v).lower())
                           for k, v in cfg.get('product_map', [])]
                # Add this result to matched using the new canonical
                if not _is_valid_version(fv):
                    continue
                if canonical not in matched or _version_lt(fv, matched[canonical]):
                    matched[canonical] = fv

    if matched:
        log.info("cve_lookup: %s → %s",
                 cve_id, ', '.join(f'{k}={v}' for k, v in matched.items()))
    elif not (auto_add_products and unmatched):
        # Only warn if we didn't handle the unmatched entries ourselves
        log.warning("cve_lookup: %s — data found but no product_map matches", cve_id)

    # ── Edge Chromium supplement ───────────────────────────────────────────────
    # When a CVE matched 'chrome' (upstream CNA records only list Google Chrome),
    # also attempt to resolve the Microsoft Edge fixed version separately.
    # Chrome and Edge are on different version trains — the Chrome version must
    # never be stored under the 'edge' canonical key.
    # Only run if 'edge' is in the configured product_map (checked via product_map
    # having an entry that maps to 'edge') and 'edge' not already matched.
    if _CHROMIUM_PRODUCTS & set(matched.keys()):
        edge_in_pm = any(v == 'edge' for _, v in product_map)
        if edge_in_pm and 'edge' not in matched:
            chrome_fixed = matched.get('chrome')
            edge_ver = _resolve_edge_version(cve_id, chrome_fixed)
            if edge_ver:
                matched['edge'] = edge_ver
                log.info("cve_lookup: %s → supplemented edge=%s", cve_id, edge_ver)

    return matched


def _version_lt(a: str, b: str) -> bool:
    """Return True if version a < version b."""
    def _t(v):
        try:
            return tuple(int(x) for x in v.split('.'))
        except ValueError:
            return (0,)
    return _t(a) < _t(b)


# ==============================================================================
# CONFIG ENRICHMENT
# ==============================================================================

def enrich_config(cve_ids: Optional[list[str]] = None,
                  config_path: Optional[str] = None,
                  dry_run: bool = False,
                  overwrite: bool = False,
                  auto_add_products: bool = True) -> dict[str, dict[str, str]]:
    """
    Look up fixed versions for a list of CVEs and write them into config.json.

    cve_ids:           list of CVE IDs to look up. If None, reads all CVE IDs already
                       present in config.json fixed_version_rules (auto mode).
    config_path:       path to config.json. Defaults to same directory as this file.
    dry_run:           if True, print what would be written without modifying config.
    overwrite:         if True, replace existing entries. Default: skip if already set.
    auto_add_products: if True (default), automatically add product_map entries and
                       fixed_version_rules keys for products found in CVE data that
                       have no existing product_map match.  Each new entry is derived
                       from the MITRE vendor/product string and logged at INFO level
                       so you can review and rename the canonical key if needed.

    Returns dict of { cve_id: { product: fixed_version } } for all found results.
    """
    if config_path is None:
        config_path = Path(__file__).parent / 'config.json'
    config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"config.json not found at {config_path}")

    with open(config_path, 'r', encoding='utf-8') as fh:
        cfg = json.load(fh)

    product_map  = [(str(k).lower(), str(v).lower()) for k, v in cfg.get('product_map', [])]
    fvr: dict    = cfg.get('fixed_version_rules', {})

    # Auto mode: collect all CVE IDs already in the rules (these need version data)
    if cve_ids is None:
        cve_ids = []
        for pk_rules in fvr.values():
            if isinstance(pk_rules, dict):
                for k in pk_rules:
                    if k.upper().startswith('CVE-') and k not in cve_ids:
                        cve_ids.append(k.upper())

    results: dict[str, dict[str, str]] = {}
    config_dirty = False   # track whether product_map was modified

    for cve_id in cve_ids:
        cve_id = cve_id.strip().upper()

        # Pass the live cfg so lookup_fixed_version can mutate product_map
        # in place when auto_add_products is True.  After a successful auto-add
        # the updated product_map is in cfg['product_map']; rebuild product_map
        # tuple list so subsequent CVEs in this loop benefit from new entries.
        matched = lookup_fixed_version(
            cve_id,
            product_map,
            cfg=cfg if auto_add_products else None,
            auto_add_products=auto_add_products,
            config_path=str(config_path),
        )

        # Rebuild product_map from cfg in case auto-add mutated it
        new_pm = [(str(k).lower(), str(v).lower()) for k, v in cfg.get('product_map', [])]
        if new_pm != product_map:
            product_map  = new_pm
            config_dirty = True

        if not matched:
            continue

        results[cve_id] = matched

        # Sync fvr from cfg (auto-add may have added new canonical keys)
        fvr = cfg.setdefault('fixed_version_rules', {})

        for pk, fv in matched.items():
            if pk not in fvr:
                fvr[pk] = {}
            if not isinstance(fvr[pk], dict):
                fvr[pk] = {}

            if cve_id in fvr[pk] and not overwrite:
                log.info("cve_lookup: %s/%s already has value %s — skipping (use --overwrite to replace)",
                         pk, cve_id, fvr[pk][cve_id])
                continue

            old = fvr[pk].get(cve_id, '(not set)')
            fvr[pk][cve_id] = fv
            cfg['fixed_version_rules'] = fvr
            config_dirty = True
            log.info("cve_lookup: config.json %s/%s: %s → %s", pk, cve_id, old, fv)

        time.sleep(_DELAY)   # be polite between CVE lookups

    if config_dirty and not dry_run:
        with open(config_path, 'w', encoding='utf-8') as fh:
            json.dump(cfg, fh, indent=2)
        n_results = len(results)
        n_pm      = len(cfg.get('product_map', [])) - len(
            [(k, v) for k, v in cfg.get('product_map', [])])
        log.info("cve_lookup: config.json updated — %d CVE(s) enriched", n_results)
    elif dry_run:
        log.info("cve_lookup: dry run — no changes written")

    return results


# ==============================================================================
# INTEGRATION WITH ORCHESTRATOR
# ==============================================================================

def enrich_from_detections(cve_df: 'pd.DataFrame',
                            config_path: Optional[str] = None,
                            auto_add_products: bool = True) -> int:
    """
    Called by the orchestrator with the loaded CVE DataFrame.
    Finds all CVE IDs in the data that don't have a fixed_version_rules entry
    and looks them up automatically.

    config.json is the persistent cache — each CVE is only looked up once per
    product. Once written, subsequent runs read from config.json and skip the API.

    auto_add_products
        When True (default), products found in CVE data that have no product_map
        match are automatically added to config.json so future runs can match
        and version-check them.  Set to False to preserve the old behaviour of
        only enriching products already in product_map.

    Returns the number of CVEs enriched.
    """
    try:
        from data_pipeline import extract_cve_id

        if config_path is None:
            config_path = Path(__file__).parent / 'config.json'

        with open(config_path, 'r', encoding='utf-8') as fh:
            cfg = json.load(fh)
        fvr          = cfg.get('fixed_version_rules', {})
        product_map  = [(str(k).lower(), str(v).lower())
                        for k, v in cfg.get('product_map', [])]

        # Build set of CVE IDs already fully covered in config.json
        known: set[str] = set()
        for pk_rules in fvr.values():
            if isinstance(pk_rules, dict):
                known.update(k.upper() for k in pk_rules
                             if k.upper().startswith('CVE-'))

        all_cves = cve_df['Vulnerability Name'].apply(extract_cve_id).unique()
        missing  = [c for c in all_cves
                    if c.upper() not in known and c.upper().startswith('CVE-')]

        if not missing:
            log.debug("cve_lookup: all %d CVEs already have version data", len(all_cves))
            return 0

        log.info("cve_lookup: %d of %d CVE(s) without version data — looking up",
                 len(missing), len(all_cves))
        results = enrich_config(
            cve_ids           = missing,
            config_path       = config_path,
            auto_add_products = auto_add_products,
        )
        return len(results)

    except Exception as e:
        log.warning("cve_lookup: auto-enrich failed: %s", e)
        return 0


# ==============================================================================
# CLI
# ==============================================================================

if __name__ == '__main__':
    import sys
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s  %(levelname)-8s  %(message)s',
        datefmt='%H:%M:%S',
    )

    p = argparse.ArgumentParser(
        description='Fetch CVE fixed versions from CVE.org/NVD and update config.json.'
    )
    p.add_argument('cves', nargs='*', metavar='CVE-ID',
                   help='CVE IDs to look up (e.g. CVE-2026-5288 CVE-2026-5289)')
    p.add_argument('--auto',      action='store_true',
                   help='Auto mode: look up all CVEs already in config.json rules')
    p.add_argument('--dry-run',   action='store_true',
                   help='Show what would be written without modifying config.json')
    p.add_argument('--overwrite', action='store_true',
                   help='Replace existing fixed_version_rules entries')
    p.add_argument('--no-auto-add', action='store_true',
                   help='Do not auto-add products to product_map when no match is found '
                        '(default: auto-add is enabled)')
    p.add_argument('--config',    default=None, metavar='PATH',
                   help='Path to config.json')
    p.add_argument('--test-nvd-api', action='store_true',
                   help='Test direct NVD API access and print clear pass/fail feedback')
    p.add_argument('--test-cve', default='CVE-2024-21413', metavar='CVE-ID',
                   help='CVE ID to use with --test-nvd-api (default: CVE-2024-21413)')
    args = p.parse_args()

    if args.test_nvd_api:
        ok = test_nvd_api_access(config_path=args.config, cve_id=args.test_cve)
        sys.exit(0 if ok else 2)

    cve_ids = None if args.auto else (args.cves or None)

    if cve_ids is None and not args.auto:
        p.print_help()
        sys.exit(1)

    results = enrich_config(
        cve_ids           = cve_ids,
        config_path       = args.config,
        dry_run           = args.dry_run,
        overwrite         = args.overwrite,
        auto_add_products = not args.no_auto_add,
    )

    if results:
        print(f"\nEnriched {len(results)} CVE(s):")
        for cve_id, matched in results.items():
            for pk, fv in matched.items():
                print(f"  {cve_id}  →  {pk}: {fv}")
    else:
        print("No new version data found.")
