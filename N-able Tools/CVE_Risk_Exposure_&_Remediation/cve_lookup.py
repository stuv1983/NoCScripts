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
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_UA      = 'Mozilla/5.0 N-able-CVE-Dashboard/1.0 (automated; contact your-email@example.com)'
_TIMEOUT    = 8    # connect + read timeout per request attempt
_RETRY      = 2    # reduced: 2 retries × 3 sources × 98 CVEs = still slow if all fail
_DELAY   = 0.5   # seconds between API calls — be a polite client


def _make_session() -> requests.Session:
    """Build a requests.Session with automatic retry and exponential backoff.
    Retries up to 3 times on 429/500/502/503/504 with backoff 1s/2s/4s.
    Connection pooling is reused across all calls within a lookup run.
    """
    session = requests.Session()
    retry = Retry(
        total            = _RETRY,
        backoff_factor   = 1,
        status_forcelist = {429, 500, 502, 503, 504},
        allowed_methods  = {'GET'},
        raise_on_status  = False,
    )
    session.mount('https://', HTTPAdapter(max_retries=retry))
    session.headers.update({'User-Agent': _UA, 'Accept': 'application/json'})
    return session


import threading as _threading
_thread_local = _threading.local()


def _get_session() -> requests.Session:
    """
    Return a requests.Session scoped to the current thread.

    Using threading.local() ensures each worker thread in the ThreadPoolExecutor
    owns its own session and connection pool, eliminating the risk of pool
    exhaustion or state leakage when multiple threads make concurrent requests.
    The module-level _SESSION global is no longer used.
    """
    session = getattr(_thread_local, 'session', None)
    if session is None:
        session = _make_session()
        _thread_local.session = session
    return session


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
    """
    HTTP GET with automatic retry and exponential backoff via requests.Session.

    The session handles 429/500/502/503/504 retries transparently with
    exponential backoff (1s, 2s, 4s).  Returns parsed JSON on success,
    None on 404/403 or after all retries are exhausted.
    """
    try:
        session = _get_session()
        resp = session.get(
            url,
            headers=extra_headers or {},
            timeout=_TIMEOUT,
        )
        if resp.status_code == 404:
            log.debug("HTTP 404 from %s — resource not found", url[:80])
            return None
        if resp.status_code == 403:
            log.warning("HTTP 403 from %s — check API key or access permissions", url[:80])
            return None
        resp.raise_for_status()
        time.sleep(_DELAY)
        return resp.json()
    except requests.exceptions.RetryError as e:
        log.warning("Fetch failed after retries: %s — %s", url[:80], e)
    except requests.exceptions.Timeout:
        log.debug("Timeout fetching %s (limit %ds)", url[:80], _TIMEOUT)
    except requests.exceptions.RequestException as e:
        log.debug("Fetch error %s: %s", url[:80], e)
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


def _cve_local_path(cve_id: str, repo_path) -> 'Optional[Path]':
    """
    Return the expected path to a CVE JSON file inside a local cvelistV5 clone.
    Layout:  cves/{year}/{prefix}xxx/CVE-{year}-{num}.json
    e.g.     cves/2024/1xxx/CVE-2024-1234.json
             cves/2024/12xxx/CVE-2024-12345.json
    """
    import re as _re
    m = _re.match(r'CVE-(\d{4})-(\d+)', str(cve_id).strip().upper())
    if not m:
        return None
    year, num = m.group(1), m.group(2)
    prefix = (num[:-3] if len(num) > 3 else '0') + 'xxx'
    return Path(repo_path) / 'cves' / year / prefix / f'CVE-{year}-{num}.json'


def _load_cve_org_local(cve_id: str, repo_path) -> 'Optional[dict]':
    """
    Read a CVE JSON 5.0 file from a local cvelistV5 clone.
    Returns the parsed dict (identical structure to CVE.org API) or None.
    Falls back silently — callers must handle None and try the API.
    """
    if not repo_path:
        return None
    p = _cve_local_path(cve_id, repo_path)
    if p is None:
        return None
    try:
        if not p.exists():
            return None
        with open(p, 'r', encoding='utf-8') as _fh:
            return json.load(_fh)
    except Exception as _e:
        log.debug('cve_lookup: local file read failed for %s: %s', cve_id, _e)
        return None


def _extract_cvss_score(cve_org_data: 'Optional[dict]',
                         nvd_data:     'Optional[dict]') -> 'Optional[float]':
    """
    Extract the best available CVSS base score from CVE.org or NVD response.
    Priority: CVSSv4.0 > CVSSv3.1 > CVSSv3.0 > CVSSv2.0.
    Checks CVE.org CNA + ADP containers first, then NVD metrics.
    Returns None if no score found in either source.
    """
    if isinstance(cve_org_data, dict):
        _containers = cve_org_data.get('containers', {})
        for _src in ([_containers.get('cna', {})] + (_containers.get('adp') or [])):
            for _m in _src.get('metrics', []):
                for _key in ('cvssV4_0', 'cvssV3_1', 'cvssV3_0', 'cvssV2_0'):
                    try:
                        s = float(_m.get(_key, {}).get('baseScore', ''))
                        return s
                    except (TypeError, ValueError):
                        pass
    if isinstance(nvd_data, dict):
        try:
            _vulns = nvd_data.get('vulnerabilities', [])
            if _vulns:
                _mets = _vulns[0].get('cve', {}).get('metrics', {})
                for _key in ('cvssMetricV40', 'cvssMetricV31', 'cvssMetricV30', 'cvssMetricV2'):
                    for _ent in _mets.get(_key, []):
                        try:
                            s = float(_ent.get('cvssData', {}).get('baseScore', ''))
                            return s
                        except (TypeError, ValueError):
                            pass
        except Exception:
            pass
    return None


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


def _edge_release_date(rel: dict) -> Optional['datetime']:
    """Parse a release date from an Edge releases API entry."""
    import datetime as _dt
    for key in ('PublishedTime', 'ReleaseTime', 'ReleaseDate'):
        val = rel.get(key)
        if val:
            try:
                # Handle ISO 8601 with optional Z/offset
                s = str(val).replace('Z', '+00:00')
                return _dt.datetime.fromisoformat(s).replace(tzinfo=None)
            except (ValueError, TypeError):
                pass
    return None


def _edge_fixed_version_for_chromium_milestone(
    chrome_fixed: str,
    cve_published: Optional[str] = None,
) -> Optional[str]:
    """
    Given a Chrome fixed version string (e.g. '146.0.7680.178'), return the
    earliest Edge Stable release that:
      (a) has the same Chromium major or later, AND
      (b) has a release date on or after the CVE publish date (when known)

    Guard (b) prevents choosing an early Edge 146 build that existed before
    the Chromium security fix actually landed in that milestone.  An Edge
    release on the right major but before the CVE was published cannot contain
    the fix.

    If cve_published is not available, (b) is skipped and only (a) applies —
    this is less precise but still better than no version at all.

    Returns None if no qualifying Edge release is found, rather than guessing.
    The Fixed Version Source will be marked accordingly so the analyst can see
    why no Edge version was stored.
    """
    import datetime as _dt

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

    # Parse CVE publish date for guard (b)
    pub_dt: Optional[_dt.datetime] = None
    if cve_published:
        try:
            s = str(cve_published).replace('Z', '+00:00')
            pub_dt = _dt.datetime.fromisoformat(s).replace(tzinfo=None)
        except (ValueError, TypeError):
            pass

    candidates = []
    for rel in stable.get('Releases', []):
        ver = str(rel.get('ProductVersion', '')).strip()
        if not _is_valid_version(ver):
            continue
        try:
            edge_major = int(ver.split('.')[0])
        except (ValueError, IndexError):
            continue

        # Guard (a): same major or later
        if edge_major < chrome_major:
            continue

        # Guard (b): release date must be on/after CVE publish date
        if pub_dt is not None:
            rel_dt = _edge_release_date(rel)
            if rel_dt is not None and rel_dt < pub_dt:
                continue  # this release predates the CVE — cannot contain the fix

        candidates.append(ver)

    if not candidates:
        log.debug(
            "cve_lookup: Edge milestone fallback: no qualifying release for "
            "Chrome major %d (cve_published=%s) — not guessing",
            chrome_major, cve_published,
        )
        return None

    # Return the lowest qualifying version (earliest safe Edge release)
    candidates.sort(key=lambda v: [int(x) for x in v.split('.') if x.isdigit()])
    result = candidates[0]
    log.debug(
        "cve_lookup: Edge milestone fallback: Chrome major %d → Edge %s "
        "(release date guard applied: cve_published=%s)",
        chrome_major, result, cve_published,
    )
    return result


def _resolve_edge_version(
    cve_id: str,
    chrome_fixed: Optional[str],
    cve_published: Optional[str] = None,
) -> tuple[Optional[str], str]:
    """
    Resolve the Edge-specific fixed version for a Chromium CVE.

    Returns (version_string_or_None, source_description).

    Strategy (in order):
      1. Edge release notes explicitly list the CVE — authoritative.
      2. Edge release on same Chromium major with release date >= CVE publish date.
      3. None — do not guess silently.

    source_description is one of:
      "Microsoft Edge release API - CVE listed"
      "Microsoft Edge release API - milestone fallback"
      "No Edge baseline defined"
    """
    # Step 1: direct CVE lookup in release notes
    ver = _edge_fixed_version_for_cve(cve_id)
    if ver:
        log.info("cve_lookup: %s → edge=%s (Microsoft Edge release API - CVE listed)", cve_id, ver)
        return ver, 'Microsoft Edge release API - CVE listed'

    # Step 2: milestone + date guard
    if chrome_fixed:
        ver = _edge_fixed_version_for_chromium_milestone(chrome_fixed, cve_published)
        if ver:
            log.info(
                "cve_lookup: %s → edge=%s (Microsoft Edge release API - milestone fallback "
                "from chrome=%s, cve_published=%s)",
                cve_id, ver, chrome_fixed, cve_published,
            )
            return ver, 'Microsoft Edge release API - milestone fallback'

    log.info(
        "cve_lookup: %s — could not resolve Edge version with date guard "
        "(chrome_fixed=%s, cve_published=%s) — stored as No Edge baseline defined",
        cve_id, chrome_fixed, cve_published,
    )
    return None, 'No Edge baseline defined'


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
    cve_repo_path=None,
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

    # Source 0: local cvelistV5 repo (fastest — file read, no network round trip)
    cve_org_data = _load_cve_org_local(cve_id, cve_repo_path) if cve_repo_path else None
    if cve_org_data:
        results = _parse_cve_org(cve_org_data)
        if results:
            log.debug("cve_lookup: %s → local repo (%d entries)", cve_id, len(results))

    # Source 1: CVE.org API (fallback when local repo unavailable or has no entry)
    # Capture separately so the Edge supplement can read cveMetadata.datePublished
    # even when NVD or OSV is used as the version-data fallback.
    if not cve_org_data:
        cve_org_data = _get(f'https://cveawg.mitre.org/api/cve/{cve_id}')
    data = cve_org_data   # alias kept so existing code below is unchanged
    if cve_org_data and not results:
        results = _parse_cve_org(cve_org_data)
        if results:
            log.debug("cve_lookup: %s → CVE.org returned %d affected entries", cve_id, len(results))

    # Source 2: NVD (fallback)
    nvd_data = None   # tracked separately for CVSS extraction
    if not results:
        time.sleep(_DELAY)
        nvd_data = _get(
            f'https://services.nvd.nist.gov/rest/json/cves/2.0?cveId={cve_id}',
            extra_headers=_nvd_headers(config_path),
        )
        data = nvd_data
        if nvd_data:
            results = _parse_nvd(nvd_data)
            if results:
                log.info("cve_lookup: %s → NVD returned %d CPE entries", cve_id, len(results))
            else:
                vulns = len(nvd_data.get('vulnerabilities', [])) if isinstance(nvd_data, dict) else 0
                log.info("cve_lookup: NVD request succeeded for %s but no fixed-version CPE ranges were found (%d record(s))", cve_id, vulns)

    # Source 3: OSV.dev (fallback)
    if not results:
        time.sleep(_DELAY)
        data = _get(f'https://api.osv.dev/v1/vulns/{cve_id}')
        if data:
            results = _parse_osv(data)
            if results:
                log.debug("cve_lookup: %s → OSV returned %d affected entries", cve_id, len(results))

    # Extract CVSS score from whichever API source had data — available even
    # when no version ranges were found (e.g. disputed/rejected CVEs).
    _cvss = _extract_cvss_score(cve_org_data, nvd_data)

    if not results:
        if _cvss is not None:
            log.debug("cve_lookup: %s — no version data but CVSS %.1f extracted",
                      cve_id, _cvss)
            return {'_cvss_score': str(_cvss)}
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
            chrome_fixed  = matched.get('chrome')
            # Extract CVE publish date from CVE.org data for the fallback date guard.
            # Use cve_org_data explicitly — 'data' may have been overwritten by
            # the NVD or OSV fallback, neither of which contains cveMetadata.
            _cve_published: Optional[str] = None
            if cve_org_data and isinstance(cve_org_data, dict):
                try:
                    _meta = cve_org_data.get('cveMetadata', {})
                    _cve_published = (
                        _meta.get('datePublished')
                        or _meta.get('dateUpdated')
                    )
                except Exception:
                    pass
            edge_ver, edge_src = _resolve_edge_version(cve_id, chrome_fixed, _cve_published)
            if edge_ver:
                matched['edge'] = edge_ver
                log.info("cve_lookup: %s → supplemented edge=%s (%s)", cve_id, edge_ver, edge_src)
            else:
                log.info(
                    "cve_lookup: %s — Edge version not stored (%s)",
                    cve_id, edge_src,
                )

    if _cvss is not None:
        matched['_cvss_score'] = str(_cvss)
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
                  auto_add_products: bool = True,
                  cve_repo_path=None) -> dict[str, dict[str, str]]:
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
    before_product_map_count = len(cfg.get('product_map', []))

    # ── No-data cache ─────────────────────────────────────────────────────────
    # CVEs that returned no version data from any source are cached in
    # config.json so subsequent runs skip the API entirely.
    # Cache is invalidated after 30 days to pick up delayed NVD/CVE.org entries.
    import datetime as _dt
    _no_data_cache: dict = cfg.setdefault('cve_no_data_cache', {})
    _cache_ttl_days = 30
    _now_str = _dt.datetime.utcnow().strftime('%Y-%m-%d')

    def _cache_is_fresh(cve: str) -> bool:
        entry = _no_data_cache.get(cve)
        if not entry:
            return False
        try:
            age = (_dt.datetime.utcnow() -
                   _dt.datetime.strptime(entry, '%Y-%m-%d')).days
            return age < _cache_ttl_days
        except (ValueError, TypeError):
            return False

    skipped_cached = [c for c in cve_ids if _cache_is_fresh(c.strip().upper())]
    cve_ids = [c for c in cve_ids if not _cache_is_fresh(c.strip().upper())]
    if skipped_cached:
        log.info("cve_lookup: skipping %d CVE(s) with no-data cache hit "
                 "(cached within %d days): %s%s",
                 len(skipped_cached), _cache_ttl_days,
                 ', '.join(skipped_cached[:5]),
                 ' ...' if len(skipped_cached) > 5 else '')

    # ── Concurrent lookup ─────────────────────────────────────────────────────
    # Sequential lookups at _DELAY apart are slow when many CVEs have no data
    # (each burns full timeout × retries across 3 sources).  Run lookups
    # concurrently with a bounded thread pool and a semaphore to stay polite.
    #
    # Thread safety notes:
    #   - cfg / product_map mutations (auto_add_products) are serialised via lock
    #   - results dict is written once per CVE after the lookup completes
    #   - _SESSION (requests) is thread-safe (connection pool is thread-safe)
    import threading
    from concurrent.futures import ThreadPoolExecutor, as_completed

    _MAX_WORKERS  = 6   # concurrent CVE lookups
    _API_SEMAPHORE = threading.Semaphore(_MAX_WORKERS)
    _cfg_lock      = threading.Lock()

    def _lookup_one(cve: str) -> tuple[str, dict]:
        """Lookup one CVE and return (cve_id, matched_dict)."""
        with _API_SEMAPHORE:
            time.sleep(_DELAY)   # polite delay inside semaphore
            return cve, lookup_fixed_version(
                cve,
                product_map,     # read-only snapshot per call
                cfg=None,        # auto_add done serially after (thread safety)
                auto_add_products=False,
                config_path=str(config_path),
                cve_repo_path=cve_repo_path,
            )

    for cve_id in cve_ids:
        cve_ids_upper = [c.strip().upper() for c in cve_ids]

    with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as pool:
        futures = {pool.submit(_lookup_one, c.strip().upper()): c
                   for c in cve_ids}
        for future in as_completed(futures):
            try:
                cve_id, matched = future.result()
            except Exception as exc:
                log.warning("cve_lookup: unexpected error for %s: %s", futures[future], exc)
                continue

            if not matched:
                # Cache the no-data result so next run skips this CVE
                with _cfg_lock:
                    _no_data_cache[cve_id] = _now_str
                    cfg['cve_no_data_cache'] = _no_data_cache
                    config_dirty = True
                continue

            results[cve_id] = matched

            # Cache CVSS score (present even when no version match was found)
            _cvss_str = matched.get('_cvss_score', '')
            if _cvss_str:
                try:
                    with _cfg_lock:
                        _csc = cfg.setdefault('cvss_score_cache', {})
                        _csc[cve_id] = float(_cvss_str)
                        config_dirty = True
                except (ValueError, TypeError):
                    pass

            with _cfg_lock:
                # Auto-add products serially (mutates product_map / cfg)
                if auto_add_products:
                    for r_vendor, r_product, r_fv in matched.get('_raw_entries', []):
                        canonical = _auto_update_product_map(
                            r_vendor, r_product, r_fv, cve_id, cfg)
                        if canonical:
                            product_map = [(str(k).lower(), str(v).lower())
                                           for k, v in cfg.get('product_map', [])]
                            config_dirty = True

                # Sync fvr from cfg
                fvr = cfg.setdefault('fixed_version_rules', {})

                for pk, fv in matched.items():
                    if pk.startswith('_'):
                        continue
                    if pk not in fvr:
                        fvr[pk] = {}
                    if not isinstance(fvr[pk], dict):
                        fvr[pk] = {}
                    if cve_id in fvr[pk] and not overwrite:
                        log.info("cve_lookup: %s/%s already has value %s — skipping",
                                 pk, cve_id, fvr[pk][cve_id])
                        continue
                    old = fvr[pk].get(cve_id, '(not set)')
                    fvr[pk][cve_id] = fv
                    cfg['fixed_version_rules'] = fvr
                    config_dirty = True
                    log.info("cve_lookup: config.json %s/%s: %s → %s", pk, cve_id, old, fv)

                # Rebuild product_map if auto-add mutated it
                new_pm = [(str(k).lower(), str(v).lower())
                          for k, v in cfg.get('product_map', [])]
                if new_pm != product_map:
                    product_map  = new_pm
                    config_dirty = True

    if config_dirty and not dry_run:
        with open(config_path, 'w', encoding='utf-8') as fh:
            json.dump(cfg, fh, indent=2)
        n_results      = len(results)
        added_products = len(cfg.get('product_map', [])) - before_product_map_count
        log.info(
            "cve_lookup: config.json updated — %d CVE(s) enriched, %d product mapping(s) added",
            n_results, added_products,
        )
    elif dry_run:
        log.info("cve_lookup: dry run — no changes written")

    return results


# ==============================================================================
# INTEGRATION WITH ORCHESTRATOR
# ==============================================================================

def enrich_from_detections(cve_df: 'pd.DataFrame',
                            config_path: Optional[str] = None,
                            auto_add_products: bool = True,
                            cve_repo_path=None) -> int:
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

        # Also skip CVEs in the no-data cache (returned nothing from all sources
        # within the last 30 days).  This prevents re-hitting APIs for CVEs
        # with no public version data yet — the main cause of slow runs.
        import datetime as _dt
        _no_data = cfg.get('cve_no_data_cache', {})
        _ttl = 30
        def _fresh(cve: str) -> bool:
            d = _no_data.get(cve)
            if not d:
                return False
            try:
                return (_dt.datetime.utcnow() -
                        _dt.datetime.strptime(d, '%Y-%m-%d')).days < _ttl
            except (ValueError, TypeError):
                return False

        all_cves = cve_df['Vulnerability Name'].apply(extract_cve_id).unique()
        missing  = [c for c in all_cves
                    if (c.upper() not in known and c.upper().startswith('CVE-')
                        and not _fresh(c.upper()))]

        cached_skip = sum(1 for c in all_cves
                          if c.upper().startswith('CVE-') and _fresh(c.upper()))
        if cached_skip:
            log.info("cve_lookup: skipping %d CVE(s) in no-data cache", cached_skip)

        if not missing:
            log.debug("cve_lookup: all %d CVEs already have version data", len(all_cves))
            return 0

        log.info("cve_lookup: %d of %d CVE(s) without version data — looking up",
                 len(missing), len(all_cves))
        results = enrich_config(
            cve_ids           = missing,
            config_path       = config_path,
            auto_add_products = auto_add_products,
            cve_repo_path     = cve_repo_path,
        )

        # Fast CVSS pre-pass: read local JSON for all CVEs that have version
        # data cached but no CVSS score yet.  File reads are ~1ms each — far
        # cheaper than API calls, so we do this for the full dataset.
        if cve_repo_path:
            try:
                with open(config_path, 'r', encoding='utf-8') as _fh:
                    _cfg2 = json.load(_fh)
                _csc  = _cfg2.setdefault('cvss_score_cache', {})
                _need = [c for c in all_cves
                         if c.upper().startswith('CVE-') and c.upper() not in _csc]
                _added = 0
                for _cve in _need:
                    _data = _load_cve_org_local(_cve.upper(), cve_repo_path)
                    if _data:
                        _sc = _extract_cvss_score(_data, None)
                        if _sc is not None:
                            _csc[_cve.upper()] = _sc
                            _added += 1
                if _added:
                    _cfg2['cvss_score_cache'] = _csc
                    with open(config_path, 'w', encoding='utf-8') as _fh:
                        json.dump(_cfg2, _fh, indent=2)
                    log.info('cve_lookup: CVSS pre-pass populated %d score(s) from local repo',
                             _added)
            except Exception as _pe:
                log.debug('cve_lookup: CVSS pre-pass failed: %s', _pe)

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