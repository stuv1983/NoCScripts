"""
data_pipeline.py — all pandas data loading, merging, patch matching, and trend comparison.
No GUI imports. No xlsxwriter. Pure data in, data out.
"""

import logging
import os
import re
from datetime import datetime
from typing import Optional, Set, Tuple

import pandas as pd

from config import (
    CVE_PATTERN, PRODUCT_MAP, FIXED_VERSION_RULES,
    STATUS_RANK, STATUS_LABEL, INSTALLED_STATUSES,
    _CONFIG,
)

log = logging.getLogger(__name__)

# ==============================================================================
# PRE-COMPILED REGEX  (compile once at import, reuse for every row)
# ==============================================================================

_KB_RE       = re.compile(r'KB\d+',                    re.IGNORECASE)
_CVE_RE      = re.compile(r'CVE-\d{4}-\d{4,7}',       re.IGNORECASE)
_VERSION_RE  = re.compile(r'\b\d+(?:\.\d+){1,4}\b')
_DIGITS_RE   = re.compile(r'\d+')
_NORM_CHARS  = re.compile(r'[^a-z0-9]+')

# get_base_product patterns
_ARCH_X64    = re.compile(r'\bx64\b',    re.IGNORECASE)
_ARCH_X86    = re.compile(r'\bx86\b',    re.IGNORECASE)
_ARCH_32     = re.compile(r'\b32-bit\b', re.IGNORECASE)
_ARCH_64     = re.compile(r'\b64-bit\b', re.IGNORECASE)
_EMPTY_PAREN = re.compile(r'\s*\(\s*\)')
_TRAILING_VER= re.compile(r'\s+v?\d[\d.+]*\s*$')
_SHEET_CHARS = re.compile(r'[\[\]\:\*\?\/\\\'\000]')

# RMM inventory column config (updateable via config.json)
_RMM_CFG = _CONFIG.get('rmm_inventory_columns', {})
_RMM_POSITIONAL = _RMM_CFG.get('positional_headers',
    ['Type','Client','Site','Device','Description','OS','Username','Last Response','Last Boot'])
_RMM_DEVICE_COL = _RMM_CFG.get('device_col', 'Device')
_RMM_RESP_COL   = _RMM_CFG.get('last_response_col', 'Last Response')
_RMM_OS_COL     = _RMM_CFG.get('os_col', 'OS')


# ==============================================================================
# PATCH GAP CLASSIFICATION
# ==============================================================================

# Maps Patch Match Result strings → explicit gap category (no-match cases).
_GAP_NO_MATCH: dict[str, str] = {
    'Not found in patch report':                  'coverage_gap',
    'Device in patch report - product not found': 'unmanaged_app',
}

# Matched-but-unresolved = patch tool says installed, N-able still detects CVE.
# These are the _STATUS_LABEL values for installed/reboot-required states.
_MATCHED_INSTALLED = {'Matched - installed', 'Matched - reboot required'}


def classify_patch_gap(patch_match_result: str,
                       resolved: Optional[str] = None) -> Optional[str]:
    """
    Return the explicit gap category for a patch row, or None if no gap.

    Categories:
        coverage_gap         device not in the patch report at all
        unmanaged_app        device present, product not tracked by patch tool
        detection_mismatch   patch tool says installed but N-able still detects
                             the CVE — scanner vs patch tool disagreement, OR
                             the patch was applied before the CVE was first
                             detected (pre-existing install, not a real fix)
    """
    pmr = str(patch_match_result).strip()

    # No-match cases
    if pmr in _GAP_NO_MATCH:
        return _GAP_NO_MATCH[pmr]

    # Matched + still unresolved → detection mismatch
    if pmr in _MATCHED_INSTALLED and str(resolved).strip() == 'Unresolved':
        return 'detection_mismatch'

    return None

def load_data(file_path: str) -> pd.DataFrame:
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    return pd.read_csv(file_path)

def normalize_device_name(name: str) -> str:
    name = str(name).strip().upper()
    if '\\' in name: name = name.split('\\')[-1]
    if '.'  in name: name = name.split('.')[0]
    return name

def get_base_product(prod_name: str) -> str:
    p = str(prod_name).strip()
    p = _ARCH_X64.sub('', p)
    p = _ARCH_X86.sub('', p)
    p = _ARCH_32.sub('', p)
    p = _ARCH_64.sub('', p)
    p = _EMPTY_PAREN.sub('', p)
    p = _TRAILING_VER.sub('', p)
    return p.strip()

def clean_sheet_name(name: str, used_names: Set[str]) -> str:
    if pd.isna(name) or str(name).strip() == '': name = 'Unknown Product'
    clean = _SHEET_CHARS.sub('', str(name)).strip()[:31].strip()
    if not clean: clean = 'Unknown Product'
    final, counter = clean, 1
    while final.lower() in {n.lower() for n in used_names}:
        suffix = f'_{counter}'
        final = clean[:31 - len(suffix)] + suffix
        counter += 1
    used_names.add(final)
    return final

def extract_cve_id(val: str) -> str:
    """Pull a bare CVE-YYYY-NNNNN from either a raw string or a HYPERLINK formula."""
    m = CVE_PATTERN.search(str(val))
    return m.group(1).upper() if m else str(val).strip().upper()

def determine_device_type(os_string: str) -> str:
    val = str(os_string).lower()
    if val in ('nan', 'unknown'): return 'Unknown'
    if 'server' in val: return 'Server'
    if 'windows 10' in val or 'windows 11' in val: return 'Workstation'
    return 'Workstation'

def parse_last_response(val):
    """Parse Last Response values into sortable timestamps."""
    val = str(val).strip()
    epoch = pd.to_datetime('1900-01-01')
    if val in ['Not Found in RMM', 'N/A', '']: return epoch
    try: return pd.to_datetime(val)
    except: pass
    if val.startswith('overdue_'):
        try: return pd.to_datetime(val.replace('overdue_', '').split(' -')[0])
        except: pass
    if 'days' in val or 'hrs' in val:
        try:
            m = _DIGITS_RE.search(val)
            days = int(m.group(0)) if m else 0
            return pd.Timestamp.now() - pd.Timedelta(days=days)
        except: pass
    return epoch

def get_col_letter(col_idx):
    letter = ''
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def _drop_internal(df):
    """Drop pipeline-only columns before writing to Excel."""
    return df.drop(columns=[c for c in ('Name_Join', 'Device_Join', 'Base Product',
                                         '_Sort_Time', '_Name_Key', '_CVE_Key',
                                         '_Checkbox_Resolved')
                             if c in df.columns], errors='ignore').copy()


# ==============================================================================
# PATCH MATCH HELPER FUNCTIONS
# ==============================================================================

def _norm_compact(v): return _NORM_CHARS.sub('', str(v).lower()).strip()
def _norm_text(v):    return _NORM_CHARS.sub(' ', str(v).lower()).strip()

STATUS_RANK = {'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
                'Pending': 3, 'Missing': 2, 'Failed': 1}
STATUS_LABEL = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
INSTALLED_STATUSES = {'Installed', 'Reboot Required'}

def _detect_product(text):
    t = _norm_text(str(text))
    for key, product in PRODUCT_MAP:
        if _norm_text(key) in t: return product
    return ''

def _extract_kbs(text) -> list:
    return sorted({kb.upper() for kb in _KB_RE.findall(str(text))})

def _extract_cves(text) -> list:
    return sorted({c.upper() for c in _CVE_RE.findall(str(text))})

def _extract_best_version(text) -> str:
    versions = _VERSION_RE.findall(str(text))
    if not versions: return ''
    return sorted(versions, key=lambda v: (len(v.split('.')), [int(x) for x in v.split('.')]))[-1]

def _parse_version(value) -> Optional[tuple]:
    parts = _DIGITS_RE.findall(str(value).strip())
    return tuple(int(p) for p in parts) if parts else None

def _version_gte(left, right):
    l, r = _parse_version(left), _parse_version(right)
    if l is None or r is None: return None
    n = max(len(l), len(r))
    return (l + (0,) * (n - len(l))) >= (r + (0,) * (n - len(r)))

def _make_excel_safe(df):
    out = df.copy()
    for col in out.columns:
        if isinstance(out[col].dtype, pd.DatetimeTZDtype):
            out[col] = out[col].dt.tz_localize(None)
    return out

def _resolve_fixed_version(row):
    """
    Look up the minimum fixed version for a CVE+product combination.

    Priority order:
      1. 'Fixed Version' column in the CVE workbook (explicit override)
      2. CVE-specific rule in config.json fixed_version_rules
      3. '_baseline' rolling minimum in config.json (updated monthly)
    """
    if 'Fixed Version' in row.index:
        v = str(row.get('Fixed Version', '')).strip()
        if v: return v, 'CVE workbook column'
    product = row.get('_pk', '')
    if not product: return '', ''
    rules = FIXED_VERSION_RULES.get(product, {})
    # CVE-specific rule takes priority
    for cve in _extract_cves(str(row.get('Vulnerability Name', ''))):
        if cve in rules: return rules[cve], f'config rule ({cve})'
    # Rolling baseline fallback
    baseline = rules.get('_baseline', '').strip()
    if baseline: return baseline, 'rolling baseline'
    return '', ''

def _classify_version_check(row):
    status = str(row.get('Status', '')).strip()
    pv     = str(row.get('Matched Patch Version', '')).strip()
    fv     = str(row.get('Fixed Version Used', '')).strip()
    if status not in INSTALLED_STATUSES:
        return 'Patch not yet installed' if status else 'No patch evidence'
    if not fv:
        return 'Installed version found - no fixed baseline' if pv else 'Installed - version unknown'
    if not pv: return 'Fixed baseline known - installed version not found'
    cmp = _version_gte(pv, fv)
    if cmp is True:  return 'Version compliant'
    if cmp is False: return 'Below fixed version'
    return 'Version comparison failed'

def _classify_resolution(row):
    """
    A patch is 'Resolved' only when:
      1. Status is Installed or Reboot Required, AND
      2. Patch Install Date is AFTER both Date Published and First detected
         (confirms the patch was applied after the vulnerability was known/active).
    If no CVE date is available for comparison, status alone is sufficient.
    """
    if str(row.get('Status', '')).strip() not in INSTALLED_STATUSES:
        return 'Unresolved'

    install_date = row.get('Patch Install Date')
    if pd.isna(install_date) if hasattr(install_date, '__class__') else not install_date:
        return 'Unresolved'  # Installed status but no recorded install date

    try:
        install_dt = pd.to_datetime(install_date, errors='coerce')
        if pd.isna(install_dt):
            return 'Unresolved'

        # Collect CVE reference dates — patch must post-date ALL of these
        cve_dates = []
        for col in ('Date Published', 'First detected'):
            v = row.get(col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                dt = pd.to_datetime(v, errors='coerce')
                if not pd.isna(dt):
                    cve_dates.append(dt)

        if not cve_dates:
            return 'Resolved'   # No reference dates — trust the installed status

        # Patch must be installed on or after the latest CVE reference date
        if install_dt >= max(cve_dates):
            return 'Resolved'
        return 'Unresolved'     # Patch predates CVE detection — not confirmed resolved
    except Exception:
        return 'Resolved'       # Fallback: date parse failed, trust installed status


# ==============================================================================
# DATA PIPELINE: CVE + RMM
# ==============================================================================

def load_vulnerability_data(file_path: str) -> pd.DataFrame:
    """
    Load vulnerability/CVE data from a CSV or XLSX file.

    If the file is a previously generated dashboard workbook (identified by
    having a 'Raw Data' sheet), that sheet is used directly rather than the
    default first sheet, which would be 'Trend Summary' and contain no CVE rows.
    """
    if str(file_path).lower().endswith(('.xlsx', '.xls')):
        xl = pd.ExcelFile(file_path)
        if 'Raw Data' in xl.sheet_names:
            log.info("Detected dashboard workbook — reading 'Raw Data' sheet")
            df = xl.parse('Raw Data')
        else:
            df = xl.parse(xl.sheet_names[0])
    else:
        df = pd.read_csv(file_path)

    rename = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if   c in ('asset name', 'device name', 'endpoint'):
            rename[col] = 'Name'
        elif c in ('vulnerability id', 'cve id', 'cve'):
            rename[col] = 'Vulnerability Name'
        elif c in ('cvss score', 'cvss v3.1 base score', 'cvss v3 base score',
                   'base score', 'score'):
            rename[col] = 'Vulnerability Score'
        elif c in ('affected products', 'product'):
            rename[col] = 'Affected Products'
        elif c in ('severity', 'risk'):
            rename[col] = 'Vulnerability Severity'
        elif c in ('threat status',):
            rename[col] = 'Threat Status'
        # N-able exports use 'Customer Name' / 'Site Name' — normalise for patch matching
        elif c in ('customer name', 'client name', 'account name', 'client'):
            rename[col] = 'Customer'
        elif c in ('site name', 'location name'):
            rename[col] = 'Site'
        # Exploit/KEV fields vary across N-able export versions
        # Only rename if the canonical target column doesn't already exist
        elif c in ('has exploit', 'exploit') and 'Has Known Exploit' not in df.columns:
            rename[col] = 'Has Known Exploit'
        elif c == 'cisa kev' and col != 'CISA KEV' and 'CISA KEV' not in df.columns:
            rename[col] = 'CISA KEV'
        # Date field capitalisation varies (e.g. 'First Detected' vs 'First detected')
        elif c == 'first detected' and col != 'First detected':
            rename[col] = 'First detected'
        elif c == 'last updated' and col != 'Last updated':
            rename[col] = 'Last updated'
        elif c in ('updates available', 'update available'):
            rename[col] = 'Update Available'
        elif c in ('operating system role', 'os role'):
            rename[col] = 'Operating System Role'
    df.rename(columns=rename, inplace=True)

    if 'Threat Status' in df.columns:
        df = df[df['Threat Status'].astype(str).str.strip().str.upper() != 'RESOLVED']

    defaults = {
        'Name': 'Unknown Device',          'Vulnerability Name': 'Unknown CVE',
        'Affected Products': 'Unknown Product', 'Vulnerability Score': 0.0,
        'Vulnerability Severity': 'Unknown',    'Has Known Exploit': 'No',
        'CISA KEV': 'No',                       'Risk Severity Index': 'Unknown',
    }
    for col, default in defaults.items():
        if col not in df.columns: df[col] = default

    df['Vulnerability Name'] = df['Vulnerability Name'].fillna('Unknown CVE')
    df['Name_Join']          = df['Name'].apply(normalize_device_name)
    df['Affected Products']  = df['Affected Products'].fillna('Unknown Product')
    df['Base Product']       = df['Affected Products'].apply(get_base_product)
    return df

def load_rmm_data(file_path):
    df        = load_data(file_path)
    col_lower = {c.lower(): c for c in df.columns}
    dev_col = resp_col = os_col = device_type_col = None

    for key in ('device name', 'device', 'name', 'asset name', 'hostname'):
        if key in col_lower: dev_col = col_lower[key]; break
    for key in ('last response (local time)', 'last response (utc)', 'last response', 'last check-in'):
        if key in col_lower: resp_col = col_lower[key]; break
    for key in ('os version', 'os'):
        if key in col_lower: os_col = col_lower[key]; break
    # N-able Device Inventory exports include a direct 'Device type' column
    # (values: LAPTOP, SERVER, WORKSTATION, DESKTOP, etc.) — use it when present
    for key in ('device type',):
        if key in col_lower: device_type_col = col_lower[key]; break

    if not dev_col or not resp_col:
        # ── Positional fallback for N-able Device Inventory header-less exports ──
        # Column names are configurable in config.json "rmm_inventory_columns"
        # so export format changes can be fixed without touching Python source.
        cols_are_positional = all(
            isinstance(c, int) or str(c).startswith('Unnamed')
            for c in df.columns
        )
        if len(df.columns) == len(_RMM_POSITIONAL) and cols_are_positional:
            df.columns = _RMM_POSITIONAL
            dev_col, resp_col, os_col = _RMM_DEVICE_COL, _RMM_RESP_COL, _RMM_OS_COL
        else:
            raise ValueError(
                "Could not identify required columns in RMM/Device Inventory file.\n\n"
                f"Looking for:  '{_RMM_DEVICE_COL}' and '{_RMM_RESP_COL}'.\n"
                f"Found columns: {', '.join(str(c) for c in df.columns[:12])}"
                + (' ...' if len(df.columns) > 12 else '') + "\n\n"
                "To fix without code changes, update 'rmm_inventory_columns' in config.json."
            )

    df.rename(columns={dev_col: 'Device', resp_col: 'Last Response'}, inplace=True)
    df['Device_Join'] = df['Device'].apply(normalize_device_name)

    if device_type_col:
        # Map N-able device type strings → Server / Workstation / Unknown
        def _map_device_type(val):
            v = str(val).strip().upper()
            if v == 'SERVER':  return 'Server'
            if v in ('', 'NAN', 'UNKNOWN'): return 'Unknown'
            return 'Workstation'
        df['Device Type'] = df[device_type_col].apply(_map_device_type)
    elif os_col:
        df['Device Type'] = df[os_col].apply(determine_device_type)
    else:
        df['Device Type'] = 'Unknown'

    return df.drop_duplicates(subset=['Device_Join'], keep='first')

def merge_data(df_vuln, df_rmm, skip_rmm):
    # Some N-able CVE exports already include Last Response and Device Type inline.
    # Detect this upfront so we don't create _x/_y collision columns during the merge.
    vuln_has_lr = 'Last Response' in df_vuln.columns
    vuln_has_dt = 'Device Type'   in df_vuln.columns

    if not skip_rmm and df_rmm is not None:
        # Only pull the columns from RMM that the vuln data is missing
        rmm_pull = ['Device_Join']
        if not vuln_has_lr: rmm_pull.append('Last Response')
        if not vuln_has_dt: rmm_pull.append('Device Type')

        if len(rmm_pull) > 1:
            before = len(df_vuln)
            # INNER join: devices absent from the inventory are decommissioned.
            # Their CVEs are excluded entirely — no "Not Found in RMM" tracking.
            merged = pd.merge(df_vuln, df_rmm[rmm_pull],
                              left_on='Name_Join', right_on='Device_Join', how='inner')
            dropped = before - len(merged)
            if dropped:
                decom_names = (
                    set(df_vuln['Name_Join'].unique())
                    - set(df_rmm['Device_Join'].unique())
                )
                log.info(
                    "Excluded %d CVE rows for %d decommissioned device(s) "
                    "(not in Device Inventory): %s%s",
                    dropped, len(decom_names),
                    ', '.join(sorted(decom_names)[:5]),
                    ' ...' if len(decom_names) > 5 else '',
                )
            if not vuln_has_dt:
                merged['Device Type'] = merged['Device Type'].fillna('Unknown')
        else:
            merged = df_vuln.copy()
    else:
        merged = df_vuln.copy()
        if not vuln_has_lr:
            merged['Last Response'] = 'N/A'
        if not vuln_has_dt:
            if 'Operating System Role' in merged.columns:
                merged['Device Type'] = merged['Operating System Role'].str.title()
            else:
                merged['Device Type'] = 'Unknown'

    # Guarantee these columns always exist downstream
    if 'Last Response' not in merged.columns: merged['Last Response'] = 'N/A'
    if 'Device Type'   not in merged.columns: merged['Device Type']   = 'Unknown'

    # ── Device Type fallback layer 1: Operating System Role ──────────────────
    # N-able CVE exports carry an 'Operating System Role' column (WORKSTATION /
    # SERVER / UNKNOWN) that is more reliable than the free-text OS string.
    if 'Operating System Role' in merged.columns:
        _OS_ROLE_MAP = {'WORKSTATION': 'Workstation', 'SERVER': 'Server'}
        mask_unk = merged['Device Type'].astype(str).str.strip().str.lower() == 'unknown'
        merged.loc[mask_unk, 'Device Type'] = (
            merged.loc[mask_unk, 'Operating System Role']
            .astype(str).str.strip().str.upper()
            .map(_OS_ROLE_MAP)
            .fillna('Unknown')
        )

    # ── Device Type fallback layer 2: OS string heuristic ────────────────────
    # For any devices still Unknown, inspect the OS string.
    # "Windows 10 / 11" → Workstation; anything with "server" → Server.
    if 'OS' in merged.columns:
        mask_still_unk = merged['Device Type'].astype(str).str.strip().str.lower() == 'unknown'
        if mask_still_unk.any():
            def _infer_from_os(val):
                v = str(val).lower()
                if 'server' in v:     return 'Server'
                if 'windows' in v:    return 'Workstation'
                return 'Unknown'
            merged.loc[mask_still_unk, 'Device Type'] = (
                merged.loc[mask_still_unk, 'OS'].apply(_infer_from_os)
            )

    merged['Vulnerability Score'] = pd.to_numeric(merged['Vulnerability Score'], errors='coerce')

    # _Sort_Time drives the date filter.  When RMM is skipped, Last Response = 'N/A'
    # which parse_last_response maps to epoch (1900-01-01), causing the date filter
    # to exclude every row.  Fall back to CVE detection dates in that case.
    merged['_Sort_Time'] = merged['Last Response'].apply(parse_last_response)

    _epoch = pd.to_datetime('1900-01-01')
    _stale_mask = merged['_Sort_Time'] <= _epoch

    if _stale_mask.any():
        # Try CVE date columns in preference order
        for _date_col in ('Last updated', 'First detected', 'Date Published'):
            if _date_col in merged.columns:
                _parsed = pd.to_datetime(
                    merged[_date_col].astype(str).str.replace(' UTC', '', regex=False),
                    errors='coerce',
                    utc=True,
                ).dt.tz_localize(None)
                merged.loc[_stale_mask & _parsed.notna(), '_Sort_Time'] = (
                    _parsed[_stale_mask & _parsed.notna()]
                )
                # Update mask — only rows still at epoch need the next fallback
                _stale_mask = merged['_Sort_Time'] <= _epoch
                if not _stale_mask.any():
                    break
    return merged


# ==============================================================================
# DATA PIPELINE: PATCH MATCH
# ==============================================================================

def _apply_cascade_resolution(df: pd.DataFrame) -> pd.DataFrame:
    """
    Version-based cascade resolution.

    After per-row `_classify_resolution`, find any rows still Unresolved where
    we already know from another row that the device has a product version that
    satisfies the fixed-version threshold for that CVE.

    Logic:
      1. Collect best known (device, product_key, installed_version, install_date)
         from any rows where Matched Patch Version is populated.
      2. For each CVE in FIXED_VERSION_RULES, check whether any device's known
         version for that product_key >= fixed version AND install_date >= first_detected.
      3. If so, mark the Unresolved row as 'Resolved (cascade)'.

    This specifically handles the case where CVE-5288, 5289, 5290 all share the
    same Edge fixed version. Once we confirm Edge is at 147.x (which resolves 5290),
    5288 and 5289 are also satisfied by the same installed version.

    Does NOT cross product boundaries — an Edge install never resolves a Chrome CVE.
    """
    if df.empty or 'Matched Patch Version' not in df.columns:
        return df

    df = df.copy()

    # Step 1: build best installed version per (device, product_key)
    df['_cascade_pk'] = df.get('_pk', df['Affected Products'].apply(
        lambda v: _detect_product(str(v).lower())))
    has_ver = df[df['Matched Patch Version'].astype(str).str.strip().str.len() > 2].copy()
    has_ver['_vt'] = has_ver['Matched Patch Version'].apply(_parse_version)
    has_ver = has_ver[has_ver['_vt'].notna()]
    if has_ver.empty:
        return df

    best_ver: dict[tuple, dict] = {}
    for _, row in has_ver.iterrows():
        key = (str(row['Name']), str(row['_cascade_pk']))
        vt  = row['_vt']
        if key not in best_ver or vt > best_ver[key]['_vt']:
            best_ver[key] = {
                '_vt':          vt,
                'version_str':  str(row['Matched Patch Version']),
                'install_date': row.get('Patch Install Date', pd.NaT),
            }

    # Step 2: build set of (device, product_key, cve_id) that cascade-satisfy
    # Priority: explicit CVE rule > _baseline for that product.
    # For CVEs with no explicit rule, the product _baseline acts as the threshold —
    # "if the device is above the minimum safe version, all CVEs fixed by that release
    # are considered resolved."
    cascade_resolve: set[tuple[str, str, str]] = set()

    # Build per-product baseline lookup for CVEs without explicit rules
    product_baselines: dict[str, tuple] = {}
    for pk, rules in FIXED_VERSION_RULES.items():
        if isinstance(rules, dict) and '_baseline' in rules:
            bt = _parse_version(rules['_baseline'])
            if bt:
                product_baselines[pk] = bt

    for (device, pk), ver_info in best_ver.items():
        rules = FIXED_VERSION_RULES.get(pk, {})
        if not isinstance(rules, dict):
            continue

        # Explicit per-CVE rules first
        for cve_id, fixed_str in rules.items():
            if cve_id.startswith('_'):
                continue
            fixed_t = _parse_version(fixed_str)
            if fixed_t and ver_info['_vt'] >= fixed_t:
                cascade_resolve.add((device, pk, cve_id.upper()))

        # Baseline fallback: covers all CVEs on this product that have no explicit rule.
        # We don't know the per-CVE threshold, but if the device is above the
        # minimum safe release we treat all unresolved CVEs on that product as covered.
        if pk in product_baselines and ver_info['_vt'] >= product_baselines[pk]:
            # Mark with a sentinel so we can apply to any CVE on this product
            cascade_resolve.add((device, pk, '_BASELINE_'))

    if not cascade_resolve:
        return df

    # Step 3: apply to Unresolved rows that match (device, product_key, cve_id)
    cascade_applied = 0
    for idx, row in df.iterrows():
        if str(row.get('Resolved (from Patch Report)', '')).strip() != 'Unresolved':
            continue
        device  = str(row['Name'])
        pk      = str(row.get('_cascade_pk', ''))
        cve_ids = [c.upper() for c in _extract_cves(str(row.get('Vulnerability Name', '')))]

        # Date check: known install date must post-date first_detected for this row
        key = (device, pk)
        if key not in best_ver:
            continue
        install_dt = pd.to_datetime(best_ver[key]['install_date'], errors='coerce')
        first_dt   = pd.to_datetime(row.get('First detected', pd.NaT), errors='coerce')
        if pd.isna(install_dt) or pd.isna(first_dt) or install_dt < first_dt:
            continue

        for cve_id in cve_ids:
            # Match explicit CVE rule OR baseline sentinel
            if (device, pk, cve_id) in cascade_resolve or \
               (device, pk, '_BASELINE_') in cascade_resolve:
                df.at[idx, 'Resolved (from Patch Report)'] = 'Resolved'
                cascade_applied += 1
                break

    if cascade_applied:
        log.info("Cascade resolution: %d additional rows resolved via version compliance",
                 cascade_applied)

    return df


def process_patch_match(patch_path, cve_df, min_score=9.0):
    """
    Match a Patch Report against an already-loaded CVE DataFrame.
    Returns (overview_df, full_df, patch_df, total_rows, filtered_rows).
    """
    patch  = load_data(patch_path)
    miss_p = {'Client', 'Site', 'Device', 'Status', 'Patch',
              'Discovered / Install Date'} - set(patch.columns)
    if miss_p:
        raise ValueError(f'Patch report missing required columns: {", ".join(sorted(miss_p))}')

    cve    = _drop_internal(cve_df)
    miss_c = {'Vulnerability Name', 'Name', 'Affected Products'} - set(cve.columns)
    if miss_c:
        raise ValueError(f'CVE data missing columns for patch matching: {", ".join(sorted(miss_c))}')

    if 'Customer' not in cve.columns: cve['Customer'] = ''
    if 'Site'     not in cve.columns: cve['Site']     = ''

    total_rows = len(cve)
    if 'Vulnerability Score' in cve.columns:
        cve = cve[pd.to_numeric(cve['Vulnerability Score'], errors='coerce').fillna(0) >= min_score]
    filtered_rows = len(cve)

    patch = patch.copy()
    patch['_ck']  = patch['Client'].map(_norm_compact)
    patch['_sk']  = patch['Site'].map(_norm_compact)
    patch['_dk']  = patch['Device'].map(_norm_compact)
    patch['_pk']  = patch['Patch'].map(_detect_product)
    patch['_pd']  = pd.to_datetime(patch['Discovered / Install Date'], errors='coerce')
    patch['_sr']  = patch['Status'].map(STATUS_RANK).fillna(0)
    patch['_kbs'] = patch['Patch'].apply(_extract_kbs)
    patch['_pv']  = patch['Patch'].apply(_extract_best_version)

    patch_devices = set(zip(patch['_ck'], patch['_sk'], patch['_dk']))

    cve['_ck']   = cve['Customer'].map(_norm_compact)
    cve['_sk']   = cve['Site'].map(_norm_compact)
    cve['_dk']   = cve['Name'].map(_norm_compact)
    cve['_pk']   = cve['Affected Products'].map(_detect_product)
    cve['_cves'] = cve['Vulnerability Name'].apply(_extract_cves)

    for dc in ('Date Published', 'First detected', 'Last updated'):
        if dc in cve.columns:
            cve[dc] = pd.to_datetime(
                cve[dc].astype(str).str.replace(' UTC', '', regex=False),
                errors='coerce', utc=True).dt.tz_localize(None)

    merged = cve.merge(
        patch[['_ck', '_sk', '_dk', '_pk', 'Status', 'Patch', '_pd', '_sr', '_kbs', '_pv']]
              .rename(columns={'_ck': '_mck'}),
        left_on=['_ck', '_sk', '_dk', '_pk'],
        right_on=['_mck', '_sk', '_dk', '_pk'],
        how='left', suffixes=('', '_p'),
    )
    merged = merged.sort_values(['_sr', '_pd'], ascending=[False, False], na_position='last')
    gcols  = [c for c in cve.columns if not c.startswith('_')]
    best   = merged.groupby(gcols, dropna=False, as_index=False).first()

    def _classify_match(row):
        if not pd.isna(row.get('Patch')):
            return STATUS_LABEL.get(str(row.get('Status', '')).strip(),
                                     f"Matched - {str(row.get('Status', '')).lower()}")
        if (row['_ck'], row['_sk'], row['_dk']) in patch_devices:
            return 'Device in patch report - product not found'
        return 'Not found in patch report'

    best['Patch Match Result'] = best.apply(_classify_match, axis=1)

    fv = best.apply(_resolve_fixed_version, axis=1, result_type='expand')
    fv.columns = ['Fixed Version Used', 'Fixed Version Source']
    best = pd.concat([best, fv], axis=1)

    best['Matched Patch Version']        = best['_pv'].fillna('')
    best['Matched KBs']                  = best['_kbs'].apply(
        lambda v: ', '.join(v) if isinstance(v, list) else '')
    best['Version Check Result']         = best.apply(_classify_version_check, axis=1)

    # Rename _pd → Patch Install Date BEFORE calling _classify_resolution so the
    # date-comparison logic can find the column by its final name.
    best = best.rename(columns={'Patch': 'Matched Patch', '_pd': 'Patch Install Date'})
    best['Resolved (from Patch Report)'] = best.apply(_classify_resolution, axis=1)

    # ── Cascade resolution: version-based bulk resolve ────────────────────────
    # If a device has product P at installed version V, and V >= fixed_version(CVE-X, P)
    # for the same product, mark CVE-X as Resolved regardless of whether that specific
    # CVE row had a matched patch entry.
    #
    # Example: Edge 147 resolves CVE-5290 → Edge 147 also satisfies the fixed version
    # for CVE-5288 and CVE-5289 (same Edge threshold) → cascade to Resolved.
    #
    # Constraint: install date must still post-date the CVE's First detected.
    # This does NOT cross products (Edge resolving doesn't fix a Chrome CVE).
    best = _apply_cascade_resolution(best)

    best = best.drop(columns=[c for c in best.columns if c.startswith('_')], errors='ignore')

    ov_cols = ['Name', 'Device Type', 'Threat Status', 'Vulnerability Score',
               'Affected Products', 'Date Published', 'First detected', 'Last updated',
               'Last Response', 'Matched Patch', 'Patch Install Date',
               'Patch Match Result', 'Resolved (from Patch Report)']
    overview = _make_excel_safe(best[[c for c in ov_cols if c in best.columns]])
    return overview, _make_excel_safe(best), _make_excel_safe(patch), total_rows, filtered_rows


# ==============================================================================
# DATA PIPELINE: TREND / MONTH-OVER-MONTH COMPARISON
# ==============================================================================

def load_previous_report(file_path):
    """
    Load CVE data from a previously generated dashboard workbook.
    Prefers the 'Raw Data' sheet (clean strings), falls back to 'All Detections'
    (may contain HYPERLINK formulas — extract_cve_id handles those too).
    """
    try:
        xl = pd.ExcelFile(file_path)
    except PermissionError:
        fname = os.path.basename(file_path)
        raise ValueError(
            f"'{fname}' is currently open in Excel.\n\n"
            f"Please close the file in Excel and try again."
        )
    except FileNotFoundError:
        fname = os.path.basename(file_path)
        raise ValueError(
            f"'{fname}' could not be found.\n\n"
            f"Please check the file path and try again."
        )
    except Exception as e:
        fname = os.path.basename(file_path)
        raise ValueError(
            f"Could not open '{fname}'.\n\n"
            f"Details: {e}"
        )

    target = next((s for s in ('Raw Data', 'All Detections') if s in xl.sheet_names),
                  xl.sheet_names[0])
    df = xl.parse(target)

    missing = {'Name', 'Vulnerability Name'} - set(df.columns)
    if missing:
        raise ValueError(
            f"Previous report sheet '{target}' is missing columns: {', '.join(sorted(missing))}.\n"
            "Please load a dashboard generated by this tool."
        )

    # Normalise column names that vary between N-able export versions
    prev_rename = {}
    for col in df.columns:
        c = str(col).strip().lower()
        if c in ('customer name', 'client name', 'client') and 'Customer' not in df.columns:
            prev_rename[col] = 'Customer'
        elif c in ('site name', 'location name') and 'Site' not in df.columns:
            prev_rename[col] = 'Site'
        elif c == 'first detected' and col != 'First detected':
            prev_rename[col] = 'First detected'
        elif c == 'last updated' and col != 'Last updated':
            prev_rename[col] = 'Last updated'
    if prev_rename:
        df.rename(columns=prev_rename, inplace=True)

    # Build normalised join keys; extract_cve_id handles raw IDs and HYPERLINK formulas
    df['_Name_Key'] = df['Name'].apply(normalize_device_name)
    df['_CVE_Key']  = df['Vulnerability Name'].apply(extract_cve_id)
    df['Vulnerability Score'] = pd.to_numeric(
        df.get('Vulnerability Score', 0), errors='coerce').fillna(0)

    # ── Collect manually-resolved pairs (☑) from product sheets ─────────────
    # Product sheets have a 'Resolved' column with ☑/☐ checkboxes.
    # Any pair marked ☑ in the previous report should be counted as resolved
    # in the trend analysis even if N-able still detects the CVE.
    _RESERVED = {
        'trend summary', 'overview', 'all detections', 'raw data',
        'stale excluded devices', 'new this month', 'resolved', 'persisting cves',
        'patch match overview', 'patch match full data', 'patch report (full)',
    }
    resolved_pairs = set()
    for sheet in xl.sheet_names:
        if sheet.lower() in _RESERVED:
            continue
        try:
            sdf = xl.parse(sheet)
            if not {'Resolved', 'Name', 'Vulnerability Name'}.issubset(sdf.columns):
                continue
            checked = sdf[sdf['Resolved'].astype(str).str.strip() == '☑']
            for _, row in checked.iterrows():
                resolved_pairs.add((
                    normalize_device_name(row['Name']),
                    extract_cve_id(row['Vulnerability Name']),
                ))
        except Exception:
            continue

    # Tag checkbox-resolved rows so compute_trends can honour them
    if resolved_pairs:
        df['_Checkbox_Resolved'] = df.apply(
            lambda r: (r['_Name_Key'], r['_CVE_Key']) in resolved_pairs, axis=1
        )
    else:
        df['_Checkbox_Resolved'] = False

    return df


def compute_trends(current_df, previous_df, threshold):
    """
    Compare current and previous reports at or above the score threshold.

    Snapshot metrics come from the FULL filtered datasets (no product-scope
    restriction) so "Previous Report" always reflects the actual previous run.

    CVE Movement and Device Movement are computed on the COMMON-PRODUCT scope
    so comparisons are apples-to-apples across months.

    Processing order (important — each step feeds correctly into the next):
      1. Score threshold + deduplication
      2. Snapshot metrics captured from full data
      3. Common-product scope applied to both datasets
      4. Checkbox-resolved pairs removed from cur only
      5. Set arithmetic for CVE/device movement

    Returns
    -------
    dict with keys:
        metrics       – headline numbers dict
        new_df        – CVEs in current but NOT in previous  (deduplicated)
        resolved_df   – CVEs in previous but NOT in current  (deduplicated)
        persisting_df – CVEs in BOTH reports (deduplicated)
    """
    cur  = current_df.copy()
    cur['_Name_Key'] = cur['Name'].apply(normalize_device_name)
    cur['_CVE_Key']  = cur['Vulnerability Name'].apply(extract_cve_id)

    prev = previous_df.copy()  # already has _Name_Key, _CVE_Key from load_previous_report

    # ── Step 1: Score threshold + deduplication ───────────────────────────────
    cur_t  = cur[cur['Vulnerability Score']  >= threshold].copy()
    prev_t = prev[prev['Vulnerability Score'] >= threshold].copy()

    cur_t  = cur_t.sort_values('Vulnerability Score', ascending=False)\
                  .drop_duplicates(subset=['_Name_Key', '_CVE_Key'], keep='first')
    prev_t = prev_t.sort_values('Vulnerability Score', ascending=False)\
                   .drop_duplicates(subset=['_Name_Key', '_CVE_Key'], keep='first')

    # ── Step 2: Snapshot metrics — captured from FULL data before any scope filter
    # "Previous Report had X CVEs" means ALL CVEs in that report, not just the
    # subset that overlaps with this month's product coverage.
    def _kev_count(df):
        if 'CISA KEV' not in df.columns: return 0
        return int(df[df['CISA KEV'].astype(str).str.strip().str.lower()
                       .isin(['yes', 'true', '1', 'y'])]['Vulnerability Name'].nunique())

    def _exploit_count(df):
        if 'Has Known Exploit' not in df.columns: return 0
        return int(df[df['Has Known Exploit'].astype(str).str.strip().str.lower()
                       .isin(['yes', 'true', '1', 'y'])]['Vulnerability Name'].nunique())

    def _srv_count(df):
        if 'Device Type' not in df.columns: return 0
        return int(df[df['Device Type'] == 'Server']['Name'].nunique())

    snap_prev_cves    = prev_t['_CVE_Key'].nunique()
    snap_cur_cves     = cur_t['_CVE_Key'].nunique()
    snap_prev_devices = int(prev_t['Name'].nunique())
    snap_cur_devices  = int(cur_t['Name'].nunique())
    snap_prev_kev     = _kev_count(prev_t)
    snap_cur_kev      = _kev_count(cur_t)
    snap_prev_exploit = _exploit_count(prev_t)
    snap_cur_exploit  = _exploit_count(cur_t)
    snap_prev_servers = _srv_count(prev_t)
    snap_cur_servers  = _srv_count(cur_t)

    # ── Step 3: Common-product scope (must happen BEFORE checkbox removal) ────
    # Restrict movement math to products present in BOTH reports so that
    # newly onboarded products / changed export formats don't inflate counts.
    # This is computed on the raw cur_t so checkbox removal can't affect which
    # products are considered "common".
    if 'Base Product' not in cur_t.columns:
        cur_t['Base Product']  = cur_t['Affected Products'].apply(get_base_product)
    if 'Base Product' not in prev_t.columns:
        prev_t['Base Product'] = prev_t['Affected Products'].apply(get_base_product)

    common_products = (set(cur_t['Base Product'].unique())
                       & set(prev_t['Base Product'].unique()))
    cur_scoped  = cur_t[cur_t['Base Product'].isin(common_products)].copy()
    prev_scoped = prev_t[prev_t['Base Product'].isin(common_products)].copy()

    # ── Step 4: Checkbox-resolved pairs removed from cur_scoped only ─────────
    # Pairs manually marked ☑ in previous product sheets are treated as resolved
    # even if N-able still detects them in the current scan.
    checkbox_resolved = set()
    if '_Checkbox_Resolved' in prev_t.columns:
        checkbox_resolved = set(
            zip(prev_t.loc[prev_t['_Checkbox_Resolved'], '_Name_Key'],
                prev_t.loc[prev_t['_Checkbox_Resolved'], '_CVE_Key'])
        )
    if checkbox_resolved:
        cur_scoped = cur_scoped[~cur_scoped.apply(
            lambda r: (r['_Name_Key'], r['_CVE_Key']) in checkbox_resolved, axis=1
        )]

    # ── Step 5: Set arithmetic on scoped data ─────────────────────────────────
    cur_keys  = set(zip(cur_scoped['_Name_Key'],  cur_scoped['_CVE_Key']))
    prev_keys = set(zip(prev_scoped['_Name_Key'], prev_scoped['_CVE_Key']))

    new_keys        = cur_keys  - prev_keys
    resolved_keys   = prev_keys - cur_keys
    persisting_keys = cur_keys  & prev_keys

    def _filter(df, keys):
        mask = [k in keys for k in zip(df['_Name_Key'], df['_CVE_Key'])]
        return _drop_internal(df[mask].copy())

    new_df        = _filter(cur_scoped,  new_keys).sort_values('Vulnerability Score', ascending=False)
    resolved_df   = _filter(prev_scoped, resolved_keys).sort_values('Vulnerability Score', ascending=False)
    persisting_df = _filter(cur_scoped,  persisting_keys).sort_values('Vulnerability Score', ascending=False)

    cur_cve_ids  = set(cur_scoped['_CVE_Key'].unique())
    prev_cve_ids = set(prev_scoped['_CVE_Key'].unique())

    new_cve_ids        = cur_cve_ids  - prev_cve_ids
    resolved_cve_ids   = prev_cve_ids - cur_cve_ids
    persisting_cve_ids = cur_cve_ids  & prev_cve_ids

    cur_dev_set  = set(cur_scoped['_Name_Key'].unique())
    prev_dev_set = set(prev_scoped['_Name_Key'].unique())

    metrics = {
        # ── Snapshot — full data, no product-scope restriction ────────────────
        'cur_cves':            snap_cur_cves,
        'prev_cves':           snap_prev_cves,
        'cur_devices':         snap_cur_devices,
        'prev_devices':        snap_prev_devices,
        'cur_kev':             snap_cur_kev,
        'prev_kev':            snap_prev_kev,
        'cur_exploit':         snap_cur_exploit,
        'prev_exploit':        snap_prev_exploit,
        'cur_servers':         snap_cur_servers,
        'prev_servers':        snap_prev_servers,
        # ── CVE-type movement — common-product scope + checkbox ───────────────
        # new_cve + persisting_cve == scoped cur CVEs  ✓
        # resolved_cve + persisting_cve == scoped prev CVEs  ✓
        'new_cve_count':        len(new_cve_ids),
        'resolved_cve_count':   len(resolved_cve_ids),
        'persisting_cve_count': len(persisting_cve_ids),
        # ── Device movement — common-product scope ────────────────────────────
        'new_devices':         len(cur_dev_set  - prev_dev_set),
        'remediated_devices':  len(prev_dev_set - cur_dev_set),
    }

    # ── Product-level trend (Top 10, common-product scope) ───────────────────
    cur_prod  = cur_scoped.groupby('Base Product')['_Name_Key'].nunique()
    prev_prod = prev_scoped.groupby('Base Product')['_Name_Key'].nunique()
    cur_cve_prod  = cur_scoped.groupby('Base Product')['_CVE_Key'].nunique()
    prev_cve_prod = prev_scoped.groupby('Base Product')['_CVE_Key'].nunique()
    product_trend = (
        pd.DataFrame({
            'Current':     cur_prod,
            'Previous':    prev_prod,
            'CVE_Current': cur_cve_prod,
            'CVE_Previous':prev_cve_prod,
        })
        .fillna(0).astype(int)
    )
    product_trend['Change']     = product_trend['Current']     - product_trend['Previous']
    product_trend['CVE_Change'] = product_trend['CVE_Current'] - product_trend['CVE_Previous']
    product_trend = product_trend.sort_values('Current', ascending=False).head(10)

    return {
        'metrics':                 metrics,
        'new_df':                  new_df,
        'resolved_df':             resolved_df,
        'persisting_df':           persisting_df,
        'product_trend':           product_trend,
        'checkbox_resolved_count': len(checkbox_resolved),
    }


# ==============================================================================
# PATCH DIAGNOSTICS  (lag · version drift · mismatch)
# ==============================================================================

def compute_patch_diagnostics(patch_full_df: pd.DataFrame) -> dict:
    """
    Answer "why is patching failing on these devices?" by computing three signals
    that go beyond the binary resolved/unresolved classification.

    patch_full_df
        The 'Patch Match Full Data' DataFrame produced by process_patch_match.

    Returns a dict with three keys:

    patch_lag_df
        Per device-CVE pair: how many days between CVE first detection and
        the patch being applied.  Negative lag = patch predated detection
        (likely why it shows as Unresolved despite being installed).
        Sorted by lag descending (longest-outstanding first).

    version_drift_df
        Per product: the min, max, and spread of installed patch versions
        across the device fleet.  Large spread = inconsistent update cadence.
        Sorted by version_spread descending.

    mismatch_summary_df
        Device-CVE pairs classified as detection_mismatch.
        Includes the installed version, the fixed-version baseline (if known),
        and the lag — so the analyst can see whether the issue is a stale
        scanner signature or a genuinely ineffective patch.
    """
    df = patch_full_df.copy()
    required = {'Name', 'Vulnerability Name', 'Patch Match Result',
                'Resolved (from Patch Report)'}
    if not required.issubset(df.columns):
        log.warning("compute_patch_diagnostics: missing columns %s — skipping",
                    required - set(df.columns))
        return {'patch_lag_df': pd.DataFrame(),
                'version_drift_df': pd.DataFrame(),
                'mismatch_summary_df': pd.DataFrame()}

    # ── Patch lag ─────────────────────────────────────────────────────────────
    lag_rows = []
    if 'Patch Install Date' in df.columns and 'First detected' in df.columns:
        for _, row in df.iterrows():
            install_dt = pd.to_datetime(row.get('Patch Install Date'), errors='coerce')
            first_dt   = pd.to_datetime(row.get('First detected'),    errors='coerce')
            if pd.isna(install_dt) or pd.isna(first_dt):
                continue
            lag_days = (install_dt - first_dt).days
            lag_rows.append({
                'Device':             row.get('Name', ''),
                'CVE':                extract_cve_id(str(row.get('Vulnerability Name', ''))),
                'Product':            row.get('Affected Products', ''),
                'First Detected':     first_dt.date(),
                'Patch Install Date': install_dt.date(),
                'Lag (days)':         lag_days,
                'Status':             row.get('Resolved (from Patch Report)', ''),
            })
    patch_lag_df = (pd.DataFrame(lag_rows)
                    .sort_values('Lag (days)', ascending=False)
                    .reset_index(drop=True)
                    if lag_rows else pd.DataFrame())
    if not patch_lag_df.empty:
        avg = patch_lag_df['Lag (days)'].mean()
        neg = (patch_lag_df['Lag (days)'] < 0).sum()
        log.info("Patch lag: avg=%.0f days, %d pairs with negative lag (install predates detection)",
                 avg, neg)

    # ── Version drift ─────────────────────────────────────────────────────────
    drift_rows = []
    if 'Matched Patch Version' in df.columns:
        df['_bp'] = df['Affected Products'].apply(get_base_product)
        for product, grp in df.groupby('_bp'):
            versions = (grp['Matched Patch Version']
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .loc[lambda s: s.str.len() > 0]
                        .unique()
                        .tolist())
            if len(versions) < 2:
                continue
            parsed = [v for v in (_parse_version(v) for v in versions) if v]
            if len(parsed) < 2:
                continue
            spread = len(set(versions))
            drift_rows.append({
                'Product':          product,
                'Distinct Versions': spread,
                'Min Version':      min(versions, key=lambda v: _parse_version(v) or (0,)),
                'Max Version':      max(versions, key=lambda v: _parse_version(v) or (0,)),
                'Versions Seen':    ', '.join(sorted(set(versions))),
                'Device Count':     grp['Name'].nunique(),
            })
    version_drift_df = (pd.DataFrame(drift_rows)
                        .sort_values('Distinct Versions', ascending=False)
                        .reset_index(drop=True)
                        if drift_rows else pd.DataFrame())
    if not version_drift_df.empty:
        log.info("Version drift: %d products with inconsistent installed versions",
                 len(version_drift_df))

    # ── Detection mismatch ────────────────────────────────────────────────────
    mismatch_rows = []
    for _, row in df.iterrows():
        gap = classify_patch_gap(
            row.get('Patch Match Result', ''),
            row.get('Resolved (from Patch Report)', ''),
        )
        if gap != 'detection_mismatch':
            continue
        install_dt = pd.to_datetime(row.get('Patch Install Date'), errors='coerce')
        first_dt   = pd.to_datetime(row.get('First detected'),    errors='coerce')
        lag = (install_dt - first_dt).days if not (pd.isna(install_dt) or pd.isna(first_dt)) else None
        mismatch_rows.append({
            'Device':               row.get('Name', ''),
            'CVE':                  extract_cve_id(str(row.get('Vulnerability Name', ''))),
            'Product':              row.get('Affected Products', ''),
            'Patch Match Result':   row.get('Patch Match Result', ''),
            'Installed Version':    row.get('Matched Patch Version', ''),
            'Fixed Version Needed': row.get('Fixed Version Used', ''),
            'Patch Install Date':   row.get('Patch Install Date', ''),
            'First Detected':       row.get('First detected', ''),
            'Lag (days)':           lag,
            'Likely Cause':         (
                'Install predates CVE detection — patch may not address this CVE'
                if lag is not None and lag < 0
                else 'Patch installed but CVE still detected — scanner/patch tool disagreement'
            ),
        })
    mismatch_summary_df = (pd.DataFrame(mismatch_rows)
                           .reset_index(drop=True)
                           if mismatch_rows else pd.DataFrame())
    if not mismatch_summary_df.empty:
        log.warning("Detection mismatches: %d device-CVE pairs where patch "
                    "is installed but CVE still detected", len(mismatch_summary_df))

    return {
        'patch_lag_df':       patch_lag_df,
        'version_drift_df':   version_drift_df,
        'mismatch_summary_df': mismatch_summary_df,
    }


# ==============================================================================
# PATCH FAILURE REPORT
# ==============================================================================

_FAIL_CATEGORY_MAP = {
    'reboot_pending':          'Reboot required before patch can install',
    'catalog_miss':            'Patch not found in WUA catalog — may be superseded',
    'network_timeout':         'Network timeout during patch download',
    'cert_failure':            'Certificate verification failed — PME cache may need clearing',
    'checksum_failure':        'Patch file checksum error — corrupted download',
    'feature_update_conflict': 'Feature Update in progress — retry after update completes',
    'third_party_unknown':     'Third-party patch application not recognised by RMM',
    'agent_timeout':           'RMM agent timed out during install',
    'install_error':           'Installer returned error code',
    'unknown':                 'Unknown failure',
}

def _classify_failure_reason(reason: str) -> str:
    r = str(reason).lower()
    if 'reboot is required'         in r: return 'reboot_pending'
    if 'not found in wua catalog'   in r: return 'catalog_miss'
    if 'certificate verification'   in r: return 'cert_failure'
    if 'checksum'                   in r: return 'checksum_failure'
    if 'feature update'             in r: return 'feature_update_conflict'
    if 'unknown application'        in r: return 'third_party_unknown'
    if 'timeout'                    in r: return 'network_timeout'
    if "couldn't download"          in r: return 'network_timeout'
    if 'incomplete download'        in r: return 'network_timeout'
    if 'timed out'                  in r: return 'agent_timeout'
    if 'operation was canceled'     in r: return 'agent_timeout'
    if 'value does not fall'        in r: return 'install_error'
    if 'installation error'         in r: return 'install_error'
    if 'fatal error'                in r: return 'install_error'
    if 'process exited'             in r: return 'install_error'
    return 'unknown'


def load_patch_failure_report(file_path: str) -> pd.DataFrame:
    """
    Load and classify a patch failure report CSV.

    Returns a DataFrame with columns:
        Device, Site, Patch, Failure Status, Failure Reason,
        _device_norm, _failure_cat, _failure_desc, _kbs
    """
    df = load_data(file_path)

    rename = {}
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'device':         rename[col] = 'Device'
        elif cl == 'site':         rename[col] = 'Site'
        elif cl == 'client':       rename[col] = 'Client'
        elif cl == 'patch':        rename[col] = 'Patch'
        elif 'failure status' in cl: rename[col] = 'Failure Status'
        elif 'failure reason' in cl: rename[col] = 'Failure Reason'
        elif 'time' in cl:         rename[col] = 'Time'
    df = df.rename(columns=rename)

    df['_device_norm']   = df['Device'].apply(normalize_device_name)
    df['_failure_cat']   = df['Failure Reason'].apply(_classify_failure_reason)
    df['_failure_desc']  = df['_failure_cat'].map(_FAIL_CATEGORY_MAP)
    df['_kbs']           = df['Patch'].astype(str).apply(_extract_kbs)

    log.info("Patch failure report: %d rows, %d devices, %d distinct KBs",
             len(df), df['_device_norm'].nunique(),
             len({kb for kbs in df['_kbs'] for kb in kbs}))
    return df


def build_patch_failure_lookup(failure_df: pd.DataFrame) -> dict:
    """
    Build a dict keyed by normalised device name → failure summary dict.

    {
      'DEVICE-01': {
          'failure_count': 40,
          'unique_kbs': 6,
          'top_category': 'catalog_miss',
          'top_description': 'Patch not found in WUA catalog ...',
          'categories': {'catalog_miss': 30, 'network_timeout': 10},
      },
      ...
    }
    """
    result = {}
    for device, grp in failure_df.groupby('_device_norm'):
        cats    = grp['_failure_cat'].value_counts().to_dict()
        top_cat = grp['_failure_cat'].value_counts().index[0]
        result[device] = {
            'failure_count':    len(grp),
            'unique_kbs':       len({kb for kbs in grp['_kbs'] for kb in kbs}),
            'top_category':     top_cat,
            'top_description':  _FAIL_CATEGORY_MAP.get(top_cat, top_cat),
            'categories':       cats,
        }
    return result
