# ==============================================================================
# N-ABLE CVE DASHBOARD & PATCH MATCH UTILITY
# Description: Merges N-able Vulnerability and RMM Device reports to create
#              an actionable, Excel-based Executive Risk Dashboard for MSPs.
#              Optionally matches a Patch Report against detected CVEs.
#              Optionally compares against a previous report to track trends.
# Features:    Executive risk metrics, stale device tracking, dynamic date
#              filtering, automated triage sheets, optional patch match /
#              version compliance analysis, and month-over-month trend analysis.
# ==============================================================================

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import re
import threading
from datetime import datetime

# Pre-compile regex for performance
CVE_PATTERN = re.compile(r'(CVE-\d{4}-\d{4,7})', re.IGNORECASE)


# ==============================================================================
# SHARED HELPER FUNCTIONS
# ==============================================================================

def select_file(label_var, filetypes=None):
    if filetypes is None:
        filetypes = [("Data Files", "*.csv *.xlsx *.xls"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        label_var.set(path)

def load_data(file_path):
    if file_path.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    return pd.read_csv(file_path)

def normalize_device_name(name):
    name = str(name).strip().upper()
    if '\\' in name: name = name.split('\\')[-1]
    if '.'  in name: name = name.split('.')[0]
    return name

def get_base_product(prod_name):
    p = str(prod_name).strip()
    p = re.sub(r'\bx64\b',    '', p, flags=re.IGNORECASE)
    p = re.sub(r'\bx86\b',    '', p, flags=re.IGNORECASE)
    p = re.sub(r'\b32-bit\b', '', p, flags=re.IGNORECASE)
    p = re.sub(r'\b64-bit\b', '', p, flags=re.IGNORECASE)
    # Remove empty parentheses left behind after arch strip e.g. "(x86)" → "()" → ""
    p = re.sub(r'\s*\(\s*\)', '', p)
    # Strip trailing version numbers including "80+" style (N-able names like "Edge 80+")
    p = re.sub(r'\s+v?\d[\d.+]*\s*$', '', p)
    return p.strip()

def clean_sheet_name(name, used_names):
    if pd.isna(name) or str(name).strip() == '': name = 'Unknown Product'
    clean = re.sub(r'[\[\]\:\*\?\/\\\'\000]', '', str(name)).strip()[:31].strip()
    if not clean: clean = 'Unknown Product'
    final, counter = clean, 1
    while final.lower() in {n.lower() for n in used_names}:
        suffix = f'_{counter}'
        final = clean[:31 - len(suffix)] + suffix
        counter += 1
    used_names.add(final)
    return final

def make_cve_org_link(val):
    val_str = str(val) if not isinstance(val, str) else val
    if pd.isna(val) or val_str.strip() == '' or val_str.lower() == 'nan': return val
    m = CVE_PATTERN.search(val_str)
    if m:
        cve_id = m.group(1).upper()
        display = val_str.replace('"', '""')
        if len(display) > 250: display = display[:247] + '...'
        return f'=HYPERLINK("https://www.cve.org/CVERecord?id={cve_id}", "{display}")'
    return val

def extract_cve_id(val):
    """Pull a bare CVE-YYYY-NNNNN from either a raw string or a HYPERLINK formula."""
    m = CVE_PATTERN.search(str(val))
    return m.group(1).upper() if m else str(val).strip().upper()

def determine_device_type(os_string):
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
            m = re.search(r'(\d+)\s*days', val)
            days = int(m.group(1)) if m else 0
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

def _norm_compact(v): return re.sub(r'[^a-z0-9]+', '', str(v).lower()).strip()
def _norm_text(v):    return re.sub(r'[^a-z0-9]+', ' ', str(v).lower()).strip()

_PRODUCT_MAP = [
    ('mozilla firefox (x64)',                  'firefox'),
    ('mozilla firefox (x86)',                  'firefox'),
    ('firefox (x64)',                          'firefox'),
    ('firefox (x86)',                          'firefox'),
    ('mozilla firefox',                        'firefox'),
    ('firefox',                                'firefox'),
    ('google chrome',                          'chrome'),
    ('chrome (x64)',                           'chrome'),
    ('chrome (x86)',                           'chrome'),
    ('chrome',                                 'chrome'),
    ('microsoft edge 80',                      'edge'),
    ('microsoft edge',                         'edge'),
    ('edge',                                   'edge'),
    ('vlc media player (x64)',                 'vlc'),
    ('vlc media player (x86)',                 'vlc'),
    ('vlc media player',                       'vlc'),
    ('vlc (x64)',                              'vlc'),
    ('vlc (x86)',                              'vlc'),
    ('vlc',                                    'vlc'),
    ('microsoft sql server management studio', 'ssms'),
    ('sql server management studio',           'ssms'),
    ('microsoft office 365',                   'office365'),
    ('office 365',                             'office365'),
    ('microsoft office',                       'office'),
    ('office',                                 'office'),
    ('windows 11',                             'windows'),
    ('windows 10',                             'windows'),
    ('windows',                                'windows'),
]

# Optional: populate fixed-version baselines here.
# Example:  BUILTIN_FIXED_VERSION_RULES = {"chrome": {"CVE-2026-12345": "136.0.7103.114"}}
BUILTIN_FIXED_VERSION_RULES = {}

_STATUS_RANK = {'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
                'Pending': 3, 'Missing': 2, 'Failed': 1}
_STATUS_LABEL = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
_INSTALLED_STATUSES = {'Installed', 'Reboot Required'}

def _detect_product(text):
    t = _norm_text(str(text))
    for key, product in _PRODUCT_MAP:
        if _norm_text(key) in t: return product
    return ''

def _extract_kbs(text):
    return sorted({kb.upper() for kb in re.findall(r'KB\d+', str(text), re.IGNORECASE)})

def _extract_cves(text):
    return sorted({c.upper() for c in re.findall(r'CVE-\d{4}-\d{4,7}', str(text), re.IGNORECASE)})

def _extract_best_version(text):
    versions = re.findall(r'\b\d+(?:\.\d+){1,4}\b', str(text))
    if not versions: return ''
    return sorted(versions, key=lambda v: (len(v.split('.')), [int(x) for x in v.split('.')]))[-1]

def _parse_version(value):
    parts = re.findall(r'\d+', str(value).strip())
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
    if 'Fixed Version' in row.index:
        v = str(row.get('Fixed Version', '')).strip()
        if v: return v, 'CVE workbook column'
    product = row.get('_pk', '')
    if not product: return '', ''
    rules = BUILTIN_FIXED_VERSION_RULES.get(product, {})
    for cve in _extract_cves(str(row.get('Vulnerability Name', ''))):
        if cve in rules: return rules[cve], f'Built-in rule ({cve})'
    return '', ''

def _classify_version_check(row):
    status = str(row.get('Status', '')).strip()
    pv     = str(row.get('Matched Patch Version', '')).strip()
    fv     = str(row.get('Fixed Version Used', '')).strip()
    if status not in _INSTALLED_STATUSES:
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
    if str(row.get('Status', '')).strip() not in _INSTALLED_STATUSES:
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

def load_vulnerability_data(file_path):
    df = load_data(file_path)
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
        if len(df.columns) == 9:
            df.columns = ['Type', 'Client', 'Site', 'Device', 'Description',
                          'OS', 'Username', 'Last Response', 'Last Boot']
            dev_col, resp_col, os_col = 'Device', 'Last Response', 'OS'
        else:
            raise ValueError("Could not identify 'Device name' and 'Last response' columns in RMM data.")

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
            merged = pd.merge(df_vuln, df_rmm[rmm_pull],
                              left_on='Name_Join', right_on='Device_Join', how='left')
            if not vuln_has_lr:
                merged['Last Response'] = merged['Last Response'].fillna('Not Found in RMM')
            if not vuln_has_dt:
                merged['Device Type'] = merged['Device Type'].fillna('Unknown')
        else:
            # CVE export already has both columns — no merge needed
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
    merged['_Sort_Time']          = merged['Last Response'].apply(parse_last_response)
    return merged


# ==============================================================================
# DATA PIPELINE: PATCH MATCH
# ==============================================================================

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
    patch['_sr']  = patch['Status'].map(_STATUS_RANK).fillna(0)
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
            return _STATUS_LABEL.get(str(row.get('Status', '')).strip(),
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
    import os
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
# EXCEL SHEET BUILDERS
# ==============================================================================

# ── Link helpers ──────────────────────────────────────────────────────────────
# IMPORTANT: do NOT use =HYPERLINK(...) formula strings via df.to_excel().
# xlsxwriter writes them as formulas with no cached result → openpyxl/pandas
# reads the cell back as 0.  Use write_url() instead: it stores the display
# text as the cached value so the cell is correctly readable by any reader.

def _write_cve_links(ws, vuln_name_series, col_idx, link_fmt):
    """
    Overwrite the Vulnerability Name column cells with proper CVE.org hyperlinks.
    Call AFTER df.to_excel() — row 0 is the header, data starts at row 1.
    write_url() caches the display string so openpyxl reads 'CVE-YYYY-NNNNN'
    instead of the formula's uncalculated default value of 0.
    """
    for row_i, val in enumerate(vuln_name_series, start=1):
        val_str = str(val)
        m = CVE_PATTERN.search(val_str)
        if m:
            cve_id  = m.group(1).upper()
            display = val_str[:255] if len(val_str) <= 255 else val_str[:252] + '...'
            ws.write_url(row_i, col_idx,
                         f'https://www.cve.org/CVERecord?id={cve_id}',
                         link_fmt, string=display)

def _write_nvd_links(ws, vuln_name_series, col_idx, link_fmt):
    """
    Write NVD links into the NVD column using write_url (not formula strings).
    Call AFTER df.to_excel().
    """
    for row_i, val in enumerate(vuln_name_series, start=1):
        m = CVE_PATTERN.search(str(val))
        if m:
            cve_id = m.group(1).upper()
            ws.write_url(row_i, col_idx,
                         f'https://nvd.nist.gov/vuln/detail/{cve_id}',
                         link_fmt, string='View')

# ── Trend Summary Sheet ───────────────────────────────────────────────────────

def build_trend_summary_sheet(workbook, trend, threshold, prev_report_name, header_fmt,
                               customer_name=''):
    """
    First sheet in the workbook. Color-coded headline M-o-M comparison.
    Green = improving, red = worsening.
    """
    ws = workbook.add_worksheet('Trend Summary')
    m  = trend['metrics']

    title_fmt = workbook.add_format({
        'bold': True, 'font_size': 14,
        'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1,
    })
    sub_fmt   = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
    lbl_fmt   = workbook.add_format({'bold': True})
    up_fmt    = workbook.add_format({'font_color': '#C00000', 'bold': True})  # worse
    down_fmt  = workbook.add_format({'font_color': '#375623', 'bold': True})  # better
    same_fmt  = workbook.add_format({'font_color': '#595959'})
    sect_fmt  = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})

    ws.set_column('A:A', 38)
    ws.set_column('B:D', 16)

    title_text = (f'{customer_name}  —  ' if customer_name else '') + 'Month-over-Month Trend Analysis'
    ws.merge_range('A1:D1', title_text, title_fmt)
    ws.write('A2', f'Compared against:  {prev_report_name}')
    ws.write('A3', f'Score threshold:   {threshold}+')

    row = 4
    for col, hdr in enumerate(['Metric', 'Previous Report', 'This Report', 'Change']):
        ws.write(row, col, hdr, sub_fmt)

    def write_row(r, label, prev, cur, lower_is_better=True):
        diff = cur - prev
        if diff == 0:
            ch_str, ch_fmt = '  —  no change', same_fmt
        elif (diff < 0) == lower_is_better:
            ch_str, ch_fmt = f'  ▼  {abs(diff):,}', down_fmt
        else:
            ch_str, ch_fmt = f'  ▲  {abs(diff):,}', up_fmt
        ws.write(r, 0, label, lbl_fmt)
        ws.write(r, 1, f'{prev:,}')
        ws.write(r, 2, f'{cur:,}')
        ws.write(r, 3, ch_str, ch_fmt)

    row += 1; ws.merge_range(row, 0, row, 3, f'  Snapshot  (score ≥ {threshold})', sect_fmt)
    row += 1; write_row(row, 'Unique CVEs (vulnerability types)', m['prev_cves'],    m['cur_cves'])
    row += 1; write_row(row, 'Unique devices affected',           m['prev_devices'], m['cur_devices'])
    # 'Unique CVE-device pairs' removed — metric causes more confusion than value
    row += 1; write_row(row, 'CVEs with known exploit',           m['prev_exploit'], m['cur_exploit'])
    row += 1; write_row(row, 'CISA KEV CVEs',                     m['prev_kev'],     m['cur_kev'])
    row += 1; write_row(row, 'Servers affected',                  m['prev_servers'], m['cur_servers'])

    row += 2; ws.merge_range(row, 0, row, 3, '  CVE Movement  (unique CVE types)', sect_fmt)
    nc, rc, pc = m['new_cve_count'], m['resolved_cve_count'], m['persisting_cve_count']
    row += 1
    ws.write(row, 0, 'New CVE types introduced', lbl_fmt)
    ws.write(row, 2, f'{nc:,}')
    ws.write(row, 3, f'  ▲  {nc:,}' if nc else '  —  none', up_fmt if nc else same_fmt)
    row += 1
    ws.write(row, 0, 'CVE types resolved / no longer detected', lbl_fmt)
    ws.write(row, 2, f'{rc:,}')
    ws.write(row, 3, f'  ▼  {rc:,}' if rc else '  —  none', down_fmt if rc else same_fmt)
    row += 1
    ws.write(row, 0, 'CVE types persisting from last period', lbl_fmt)
    ws.write(row, 2, f'{pc:,}')
    ws.write(row, 3, '  (see Persisting CVEs sheet)', same_fmt)
    row += 1
    note_fmt = workbook.add_format({'font_color': '#595959', 'italic': True})
    ws.write(row, 0, f'  ✓  {nc} + {pc} = {nc+pc} unique CVEs this report  |  {rc} + {pc} = {rc+pc} unique CVEs previous', note_fmt)

    row += 2; ws.merge_range(row, 0, row, 3, '  Device Movement', sect_fmt)
    row += 1
    ws.write(row, 0, 'New devices appearing with CVEs', lbl_fmt)
    ws.write(row, 2, f"{m['new_devices']:,}")
    ws.write(row, 3,
             f"  ▲  {m['new_devices']:,}" if m['new_devices'] else '  —  none',
             up_fmt if m['new_devices'] else same_fmt)
    row += 1
    ws.write(row, 0, 'Devices fully remediated (no CVEs remaining)', lbl_fmt)
    ws.write(row, 2, f"{m['remediated_devices']:,}")
    ws.write(row, 3,
             f"  ▼  {m['remediated_devices']:,}" if m['remediated_devices'] else '  —  none',
             down_fmt if m['remediated_devices'] else same_fmt)

    # ── Top 10 Products (Trend) ───────────────────────────────────────────────
    product_trend = trend.get('product_trend')
    if product_trend is not None and not product_trend.empty:
        prod_hdr_fmt  = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
        prod_hdr_fmt2 = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1})
        prod_up_fmt   = workbook.add_format({'font_color': '#C00000', 'bold': True})
        prod_dn_fmt   = workbook.add_format({'font_color': '#375623', 'bold': True})
        prod_eq_fmt   = workbook.add_format({'font_color': '#595959'})

        ws.set_column('A:A', 40)
        ws.set_column('B:H', 14)

        row += 2
        ws.merge_range(row, 0, row, 7, '  Top 10 Affected Products (by unique devices)', sect_fmt)
        row += 1
        # Two header groups: Devices (blue) | Unique CVEs (green)
        ws.merge_range(row, 1, row, 3, 'Unique Devices', prod_hdr_fmt)
        ws.merge_range(row, 5, row, 7, 'Unique CVE Types', prod_hdr_fmt2)
        row += 1
        for col_i, hdr in enumerate(['Product', 'Prev', 'This', 'Δ', '', 'Prev', 'This', 'Δ']):
            ws.write(row, col_i, hdr, prod_hdr_fmt if col_i <= 3 else (prod_hdr_fmt2 if col_i >= 5 else None))

        def _ch(diff, up_f, dn_f, eq_f):
            if diff > 0:  return f'▲ {diff:,}',  up_f
            if diff < 0:  return f'▼ {abs(diff):,}', dn_f
            return '—', eq_f

        for prod, prow in product_trend.iterrows():
            row += 1
            ws.write(row, 0, str(prod), lbl_fmt)
            # Devices block
            ws.write(row, 1, int(prow['Previous']))
            ws.write(row, 2, int(prow['Current']))
            dv_str, dv_fmt = _ch(int(prow['Change']), prod_up_fmt, prod_dn_fmt, prod_eq_fmt)
            ws.write(row, 3, dv_str, dv_fmt)
            ws.write(row, 4, '')   # spacer
            # CVE Types block
            ws.write(row, 5, int(prow['CVE_Previous']))
            ws.write(row, 6, int(prow['CVE_Current']))
            cv_str, cv_fmt = _ch(int(prow['CVE_Change']), prod_up_fmt, prod_dn_fmt, prod_eq_fmt)
            ws.write(row, 7, cv_str, cv_fmt)

    # ── Checkbox-resolved note ────────────────────────────────────────────────
    cb_count = trend.get('checkbox_resolved_count', 0)
    if cb_count:
        row += 2
        note_fmt2 = workbook.add_format({'font_color': '#375623', 'italic': True})
        ws.write(row, 0,
                 f'  ☑  {cb_count} device-CVE pair(s) counted as resolved because '
                 f'they were manually marked ☑ in the previous report.',
                 note_fmt2)

    row += 2; ws.merge_range(row, 0, row, 3, '  Detail Sheets in This Workbook', sect_fmt)
    row += 1; ws.write(row, 0, f'  📋  New This Month    →  {m["new_cve_count"]} new CVE types × all affected devices')
    row += 1; ws.write(row, 0, f'  ✅  Resolved          →  {m["resolved_cve_count"]} CVE types no longer detected')
    row += 1; ws.write(row, 0, f'  ⏳  Persisting CVEs   →  {m["persisting_cve_count"]} CVE types carried over from previous report')


# ── Trend Detail Sheets ───────────────────────────────────────────────────────

def build_trend_detail_sheets(writer, workbook, trend, link_fmt, sheets_subset=None):
    """
    Write trend detail sheets: New This Month / Resolved / Persisting CVEs.
    sheets_subset: if given, only write sheets whose names are in this set.
    """
    new_bg  = workbook.add_format({'bg_color': '#FCE4D6'})  # orange – new
    res_bg  = workbook.add_format({'bg_color': '#E2EFDA'})  # green  – resolved
    per_bg  = workbook.add_format({'bg_color': '#FFF2CC'})  # yellow – persisting

    detail_cols = ['Name', 'Device Type', 'Vulnerability Name', 'Vulnerability Score',
                   'Vulnerability Severity', 'Affected Products',
                   'Has Known Exploit', 'CISA KEV', 'Last Response']

    all_sheets = [
        ('New This Month',  trend['new_df'],        new_bg,
         'New CVEs not seen in the previous report — investigate and prioritise.'),
        ('Resolved',        trend['resolved_df'],   res_bg,
         'CVEs present last report that are no longer detected — confirmed remediated.'),
        ('Persisting CVEs', trend['persisting_df'], per_bg,
         'CVEs carried over from the previous report — still unresolved.'),
    ]

    for sheet_name, df, row_fmt, note in all_sheets:
        if sheets_subset and sheet_name not in sheets_subset:
            continue
        if df.empty:
            ws = workbook.add_worksheet(sheet_name)
            ws.write(0, 0, f'No records — {note}')
            continue

        df = df.copy()
        present = [c for c in detail_cols if c in df.columns]
        df = df[present]
        df['NVD'] = ''    # placeholder; filled below with write_url

        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

        cl = df.columns.tolist()
        if 'Name'               in cl: ws.set_column(cl.index('Name'),               cl.index('Name'),               25)
        if 'Device Type'        in cl: ws.set_column(cl.index('Device Type'),        cl.index('Device Type'),        15)
        if 'Affected Products'  in cl: ws.set_column(cl.index('Affected Products'),  cl.index('Affected Products'),  30)
        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25, link_fmt)
            _write_cve_links(ws, df['Vulnerability Name'], vn_idx, link_fmt)
        if 'NVD' in cl:
            nvd_idx = cl.index('NVD')
            ws.set_column(nvd_idx, nvd_idx, 10, link_fmt)
            _write_nvd_links(ws, df['Vulnerability Name'], nvd_idx, link_fmt)

        ws.conditional_format(1, 0, len(df), len(cl) - 1,
                               {'type': 'no_blanks', 'format': row_fmt})
        ws.write(len(df) + 2, 0, f'ℹ  {note}')


# ── CVE Dashboard Sheets ──────────────────────────────────────────────────────

def build_overview_sheet(workbook, merged_df, filtered_df, triage_df, threshold,
                          product_to_sheet, header_fmt, link_fmt, customer_name='',
                          patch_confirmed_count=0, redetected_count=0,
                          sheet_name='Detections', trend_metrics=None):
    ws = workbook.add_worksheet(sheet_name)

    # ── Title row with client name ────────────────────────────────────────────
    if customer_name:
        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 14,
            'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1,
        })
        ws.merge_range('A1:F1',
                       f'{customer_name}  —  CVE Risk Dashboard  '
                       f'(Score ≥ {threshold})  —  '
                       f'{datetime.now().strftime("%B %Y")}',
                       title_fmt)
        row_offset = 2   # data starts at row index 2 (0-based: row 2 = Excel row 3)
    else:
        row_offset = 0

    is_kev     = filtered_df['CISA KEV'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])
    is_exploit = filtered_df['Has Known Exploit'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])

    kev_cves    = filtered_df[is_kev]['Vulnerability Name'].nunique()
    kev_devices = filtered_df[is_kev]['Name'].nunique()
    expl_cves   = filtered_df[is_exploit]['Vulnerability Name'].nunique()
    total_det   = filtered_df['Vulnerability Name'].nunique()   # unique CVE types ≥ threshold
    uniq_dev    = filtered_df['Name'].nunique()
    avg_per_dev = round(total_det / uniq_dev, 1) if uniq_dev > 0 else 0
    total_srv   = merged_df[merged_df['Device Type'] == 'Server']['Name'].nunique()
    srv_aff     = filtered_df[filtered_df['Device Type'] == 'Server']['Name'].nunique()
    srv_pct     = f'{round((srv_aff / total_srv) * 100, 1)}%' if total_srv > 0 else '0%'

    # Devices Not Found in RMM — build full table (device + last response + days)
    missing_df = filtered_df[filtered_df['Last Response'] == 'Not Found in RMM'].copy()
    missing_devices = sorted(missing_df['Name'].unique())

    r0 = row_offset
    ws.write(r0, 0, 'Exploitability Risk', header_fmt)
    ws.write(r0+1, 0, 'KEV CVEs');          ws.write(r0+1, 1, kev_cves)
    ws.write(r0+2, 0, 'Devices w/ KEV');    ws.write(r0+2, 1, kev_devices)
    ws.write(r0+3, 0, 'Known Exploits');    ws.write(r0+3, 1, expl_cves)

    ws.write(r0, 4, f'Exposure Density (Score {threshold}+)', header_fmt)
    ws.write(r0+1, 4, 'Unique CVEs');       ws.write(r0+1, 5, total_det)
    ws.write(r0+2, 4, 'Unique Devices');    ws.write(r0+2, 5, uniq_dev)
    ws.write(r0+3, 4, 'Avg CVEs / Device'); ws.write(r0+3, 5, avg_per_dev)
    ws.write(r0+4, 4, 'Servers Impacted');  ws.write(r0+4, 5, f'{srv_aff} ({srv_pct})')

    # ── CVE Movement Context (only shown when trend data available) ────────────
    # Answers the question "is this CVE count jump real risk or noise?"
    if trend_metrics:
        m = trend_metrics
        ctx_row = r0 + 6
        ctx_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 11,
            'bg_color': '#2E4057', 'font_color': 'white', 'border': 1,
        })
        new_fmt  = workbook.add_format({'bold': True, 'font_color': '#C00000'})  # red – new risk
        res_fmt  = workbook.add_format({'bold': True, 'font_color': '#375623'})  # green – resolved
        per_fmt  = workbook.add_format({'bold': True, 'font_color': '#7F6000'})  # amber – persisting
        note_ctx = workbook.add_format({'font_color': '#595959', 'italic': True, 'font_size': 9})

        nc = m['new_cve_count']
        rc = m['resolved_cve_count']
        pc = m['persisting_cve_count']
        prev_c = m['prev_cves']
        cur_c  = m['cur_cves']

        # Scope-mismatch delta: snapshot counts full scope, movement counts common-product scope
        # e.g. snapshot = 100, movement total = 99 → 1 CVE from new-product scope
        scope_delta_cur  = cur_c  - (nc + pc)    # CVEs in snapshot but not movement (new products)
        scope_delta_prev = prev_c - (rc + pc)    # CVEs in prev snapshot but not movement

        ws.merge_range(ctx_row, 4, ctx_row, 6,
                       f'CVE Change Context  ({prev_c} last period → {cur_c} this period)',
                       ctx_title_fmt)
        ws.write(ctx_row+1, 4, f'▲  {nc} New CVE types',       new_fmt)
        ws.write(ctx_row+1, 5,
                 'Not seen last period — genuinely new risk '
                 '(driven primarily by new vendor disclosures, not expanded scanning)', note_ctx)
        ws.write(ctx_row+2, 4, f'▼  {rc} Resolved CVE types',  res_fmt)
        ws.write(ctx_row+2, 5, 'No longer detected in environment', note_ctx)
        ws.write(ctx_row+3, 4, f'⏳  {pc} Persisting CVE types', per_fmt)
        ws.write(ctx_row+3, 5, 'Carried over — still unresolved', note_ctx)
        ws.write(ctx_row+4, 4,
                 f'✓  {nc} new + {pc} persisting = {nc+pc} in scope this period  |  '
                 f'{rc} resolved + {pc} persisting = {rc+pc} in scope previous period',
                 note_ctx)
        # Scope mismatch note — only shown when snapshot ≠ movement totals
        if scope_delta_cur > 0 or scope_delta_prev > 0:
            parts = []
            if scope_delta_cur  > 0: parts.append(f'{scope_delta_cur} this period')
            if scope_delta_prev > 0: parts.append(f'{scope_delta_prev} previous period')
            ws.write(ctx_row+5, 4,
                     f'ℹ  {" / ".join(parts)} CVE(s) excluded from movement comparison — '
                     f'tied to products not present in both periods (like-for-like scope only)',
                     note_ctx)
        ws.set_column('E:E', 48)
        ws.set_column('F:G', 38)

    # ── Fix 1: Severity counts use filtered_df (score ≥ threshold) only ──────
    row_t = r0 + 7
    ws.write(row_t, 0, f'Unique CVEs by Severity (Score {threshold}+)', header_fmt)
    sev_counts = filtered_df.drop_duplicates(subset=['Vulnerability Name'])['Vulnerability Severity'].value_counts()
    r = row_t + 1
    for sev, cnt in sev_counts.items():
        ws.write(r, 0, str(sev)); ws.write(r, 1, cnt); r += 1

    # ── Fix 2: Top 10 Products — Unique Devices + Unique CVEs columns ─────────
    row_p = max(r + 2, r0 + 14)
    hdr_small = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    ws.write(row_p, 0, f'Top 10 Products (Score {threshold}+)', header_fmt)
    ws.write(row_p, 1, 'Unique Devices', hdr_small)   # devices with ≥1 CVE for this product
    ws.write(row_p, 2, 'Unique CVE Types', hdr_small) # distinct CVE IDs for this product

    prod_devices = triage_df.groupby('Base Product')['Name'].nunique()
    prod_cves    = triage_df.groupby('Base Product')['Vulnerability Name'].nunique()
    prod_summary = pd.DataFrame({'devices': prod_devices, 'cves': prod_cves})\
                     .sort_values('devices', ascending=False).head(10)

    p = row_p + 1
    for prod, prow in prod_summary.iterrows():
        if prod in product_to_sheet:
            ws.write_url(p, 0, f"internal:'{product_to_sheet[prod]}'!A1",
                         string=str(prod), cell_format=link_fmt)
        else:
            ws.write(p, 0, str(prod))
        ws.write(p, 1, int(prow['devices']))
        ws.write(p, 2, int(prow['cves']))
        p += 1

    ws.write(row_t, 4, f'Devices by Type (Score {threshold}+)', header_fmt)
    dt_counts = filtered_df.groupby('Device Type')['Name'].nunique()
    r2 = row_t + 1
    for dt, cnt in dt_counts.items():
        ws.write(r2, 4, str(dt)); ws.write(r2, 5, cnt); r2 += 1

    row_r = max(r2 + 2, r0 + 14)
    ws.write(row_r, 4, f'Resolution Status (Score {threshold}+)', header_fmt)
    sub_grey = workbook.add_format({'font_color': '#595959', 'indent': 1})
    note_fmt_small = workbook.add_format({'font_color': '#595959', 'italic': True, 'font_size': 9})
    if product_to_sheet:
        f_res   = ' + '.join([f"COUNTIF('{s}'!A:A, \"☑\")" for s in product_to_sheet.values()])
        f_unres = ' + '.join([f"COUNTIF('{s}'!A:A, \"☐\")" for s in product_to_sheet.values()])
    else:
        f_res, f_unres = '0', '0'

    ws.write(row_r + 1, 4, 'Resolved (☑)');               ws.write_formula(row_r + 1, 5, f'={f_res}')
    ws.write(row_r + 2, 4, 'Unresolved (☐)');             ws.write_formula(row_r + 2, 5, f'={f_unres}')
    ws.write(row_r + 3, 4, 'Total (device × CVE pairs)'); ws.write_formula(row_r + 3, 5, f'={f_res} + {f_unres}')
    ws.write(row_r + 3, 6, f'— {triage_df["Name"].nunique()} unique devices,  '
                            f'{triage_df["Vulnerability Name"].nunique()} unique CVE types', note_fmt_small)

    extra_rows = 3   # rows used so far (resolved, unresolved, total)
    if patch_confirmed_count > 0:
        ws.write(row_r + 4, 4, '── Resolved breakdown ──', sub_grey)
        ws.write(row_r + 5, 4, '  Patch via RMM', sub_grey)
        ws.write(row_r + 5, 5, patch_confirmed_count)
        ws.write(row_r + 5, 6, 'pre-filled ☑ by patch report', note_fmt_small)
        ws.write(row_r + 6, 4, '  Manually Marked', sub_grey)
        ws.write_formula(row_r + 6, 5, f'={f_res} - {patch_confirmed_count}')
        ws.write(row_r + 6, 6, 'user-checked ☑', note_fmt_small)
        extra_rows = 6

    if redetected_count > 0:
        rr = row_r + extra_rows + 1
        ws.write(rr, 4, '⚠ Re-detected After Patch')
        ws.write(rr, 5, redetected_count)
        ws.write(rr, 6, 'CVEs resolved last report but still present — investigate', note_fmt_small)
        extra_rows += 1

    # Push "Devices Not Found in RMM" section below all resolution rows
    row_m = row_r + extra_rows + 2
    ws.write(row_m, 4, f'Devices Not Found in RMM (Score {threshold}+, All Dates)', header_fmt)
    ws.write(row_m, 5, 'Last Response', hdr_small)
    ws.write(row_m, 6, 'Days Since Last Response', hdr_small)

    now = datetime.now()
    mi = row_m + 1
    if not missing_devices:
        ws.write(mi, 4, 'All devices synced')
    else:
        for dev in missing_devices:
            # Get any last-response value for this device (may be a date or 'Not Found in RMM')
            dev_rows = filtered_df[filtered_df['Name'] == dev]
            lr_vals  = dev_rows['Last Response'].dropna().unique()
            lr_val   = lr_vals[0] if len(lr_vals) else 'Not Found in RMM'

            ws.write(mi, 4, str(dev))
            ws.write(mi, 5, str(lr_val))

            # Calculate days since last response
            if str(lr_val).strip() not in ('Not Found in RMM', 'N/A', ''):
                try:
                    lr_dt  = parse_last_response(lr_val)
                    days   = (now - lr_dt).days
                    ws.write(mi, 6, days if days >= 0 else '—')
                except Exception:
                    ws.write(mi, 6, '—')
            else:
                ws.write(mi, 6, '—')
            mi += 1

    ws.set_column('A:A', 38)
    ws.set_column('B:C', 14)
    ws.set_column('E:E', 48)
    ws.set_column('F:F', 22)
    ws.set_column('G:G', 24)

def build_all_detections_sheet(writer, merged_df, link_fmt, missing_row_fmt):
    df = _drop_internal(merged_df)
    df['NVD'] = ''                       # placeholder; filled below with write_url
    # DO NOT apply make_cve_org_link here — formula strings cache as 0 on disk

    cols = df.columns.tolist()
    if 'Device Type' in cols and 'Name' in cols:
        cols.insert(cols.index('Name') + 1, cols.pop(cols.index('Device Type')))
        df = df[cols]

    df = df.sort_values(by=['Vulnerability Score', 'Name'], ascending=[False, True])
    df.to_excel(writer, sheet_name='All Detections', index=False)

    ws = writer.sheets['All Detections']
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    cl = df.columns.tolist()

    # Write proper hyperlinks AFTER to_excel so display text is cached correctly
    if 'Vulnerability Name' in cl:
        vn_idx = cl.index('Vulnerability Name')
        ws.set_column(vn_idx, vn_idx, 25, link_fmt)
        _write_cve_links(ws, df['Vulnerability Name'], vn_idx, link_fmt)
    if 'NVD' in cl:
        nvd_idx = cl.index('NVD')
        ws.set_column(nvd_idx, nvd_idx, 10, link_fmt)
        _write_nvd_links(ws, df['Vulnerability Name'], nvd_idx, link_fmt)
    if 'Name' in cl:
        ws.set_column(cl.index('Name'), cl.index('Name'), 25)
    if 'Last Response' in cl:
        lr = get_col_letter(cl.index('Last Response'))
        ws.conditional_format(1, 0, len(df), len(cl) - 1, {
            'type': 'formula', 'criteria': f'=${lr}2="Not Found in RMM"',
            'format': missing_row_fmt,
        })

def build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                          patch_resolved_pairs=None, no_patch_pairs=None):
    """
    patch_resolved_pairs: (device, cve) pairs confirmed resolved via patch report → blue row
    no_patch_pairs:       (device, cve) pairs with no patch evidence in RMM → yellow row
    Unresolved rows with a known exploit get orange regardless of patch status.
    Priority: blue > orange > yellow > white
    """
    if patch_resolved_pairs is None:
        patch_resolved_pairs = set()
    if no_patch_pairs is None:
        no_patch_pairs = set()

    cols_order = ['Resolved', 'Vulnerability Name', 'Name', 'Device Type',
                  'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index',
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Affected Products', 'NVD']
    for product, group in triage_df.groupby('Base Product'):
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        group = group.sort_values(
            by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])

        # Pre-fill Resolved column: ☑ for patch-confirmed, ☐ otherwise
        def _resolved_value(row):
            nk = normalize_device_name(row['Name'])
            ck = extract_cve_id(row['Vulnerability Name'])
            return '☑' if (nk, ck) in patch_resolved_pairs else '☐'

        group.insert(0, 'Resolved', group.apply(_resolved_value, axis=1))
        group['NVD'] = ''

        final_cols = [c for c in cols_order if c in group.columns]
        group[final_cols].to_excel(writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]
        ws.autofilter(0, 0, len(group), len(final_cols) - 1)

        # Format definitions
        wb_ = writer.book
        patch_res_fmt = wb_.add_format({'bg_color': '#DEEAF1'})   # light blue   — patch via RMM
        exploit_fmt   = wb_.add_format({'bg_color': '#FFE0CC'})   # light orange — known exploit
        no_patch_fmt  = wb_.add_format({'bg_color': '#FFF2CC'})   # light yellow — no patch evidence
        cl = final_cols

        if 'Resolved' in cl:
            ri = cl.index('Resolved')
            ws.data_validation(1, ri, len(group), ri, {'validate': 'list', 'source': ['☐', '☑']})
            ws.set_column(ri, ri, 10)

        # Row highlights — priority: blue (patched) > orange (exploit) > yellow (no patch) > white
        _TRUE_VALS = {'yes', 'true', '1', 'y'}
        for row_i, (_, row) in enumerate(group[final_cols].iterrows(), start=1):
            nk = normalize_device_name(str(row.get('Name', '')))
            ck = extract_cve_id(str(row.get('Vulnerability Name', '')))
            if (nk, ck) in patch_resolved_pairs:
                ws.set_row(row_i, None, patch_res_fmt)
            elif str(row.get('Has Known Exploit', '')).strip().lower() in _TRUE_VALS:
                ws.set_row(row_i, None, exploit_fmt)
            elif (nk, ck) in no_patch_pairs:
                ws.set_row(row_i, None, no_patch_fmt)

        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25, link_fmt)
            _write_cve_links(ws, group['Vulnerability Name'], vn_idx, link_fmt)
        if 'NVD' in cl:
            nvd_idx = cl.index('NVD')
            ws.set_column(nvd_idx, nvd_idx, 10, link_fmt)
            _write_nvd_links(ws, group['Vulnerability Name'], nvd_idx, link_fmt)
        if 'Name'        in cl: ws.set_column(cl.index('Name'),        cl.index('Name'),        25)
        if 'Device Type' in cl: ws.set_column(cl.index('Device Type'), cl.index('Device Type'), 15)

        # ── Legend ───────────────────────────────────────────────────────────
        legend_row = len(group) + 3
        wb_ = writer.book
        legend_title_fmt = wb_.add_format({'bold': True, 'font_size': 9,
                                           'bg_color': '#F2F2F2', 'border': 1})
        legend_cell_fmt  = wb_.add_format({'font_size': 9, 'border': 1})
        patch_leg_fmt    = wb_.add_format({'bg_color': '#DEEAF1', 'font_size': 9, 'border': 1})
        exploit_leg_fmt  = wb_.add_format({'bg_color': '#FFE0CC', 'font_size': 9, 'border': 1})
        no_patch_leg_fmt = wb_.add_format({'bg_color': '#FFF2CC', 'font_size': 9, 'border': 1})
        normal_leg_fmt   = wb_.add_format({'bg_color': '#FFFFFF',  'font_size': 9, 'border': 1})

        ws.write(legend_row, 0, 'Legend', legend_title_fmt)
        ws.write(legend_row + 1, 0, '  (blue row)',   patch_leg_fmt)
        ws.write(legend_row + 1, 1, 'Patch via RMM — install confirmed after CVE first detected', legend_cell_fmt)
        ws.set_row(legend_row + 1, None, patch_leg_fmt)
        ws.write(legend_row + 2, 0, '  (orange row)', exploit_leg_fmt)
        ws.write(legend_row + 2, 1, 'Known active exploit — unresolved, prioritise immediately', legend_cell_fmt)
        ws.set_row(legend_row + 2, None, exploit_leg_fmt)
        ws.write(legend_row + 3, 0, '  (yellow row)', no_patch_leg_fmt)
        ws.write(legend_row + 3, 1, 'No patch evidence in RMM — vendor lag or product not scanned for patches', legend_cell_fmt)
        ws.set_row(legend_row + 3, None, no_patch_leg_fmt)
        ws.write(legend_row + 4, 0, '  (white row)',  normal_leg_fmt)
        ws.write(legend_row + 4, 1, 'Unresolved — patch available but not yet applied', legend_cell_fmt)

def build_stale_excluded_sheet(writer, stale_df):
    if stale_df.empty: return
    df = stale_df[['Name', 'Last Response', 'Device Type']].drop_duplicates(subset=['Name']).copy()
    df = df.sort_values('Last Response').rename(columns={'Name': 'Device Name'})
    df.to_excel(writer, sheet_name='Stale Excluded Devices', index=False)
    ws = writer.sheets['Stale Excluded Devices']
    ws.set_column('A:A', 35); ws.set_column('B:B', 25); ws.set_column('C:C', 20)
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)

def build_raw_data_sheet(writer, raw_df):
    df = _drop_internal(raw_df)
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    writer.sheets['Raw Data'].autofilter(0, 0, len(df), len(df.columns) - 1)

def build_patch_sheets(writer, overview_df, full_df, patch_df):
    for df, name in ((overview_df, 'Patch Match Overview'),
                     (full_df,     'Patch Match Full Data'),
                     (patch_df,    'Patch Report (Full)')):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)


# ==============================================================================
# ORCHESTRATOR / THREADING
# ==============================================================================

def execute_processing_thread(vuln_path, rmm_path, skip_rmm,
                               patch_path, include_patch,
                               prev_report_path, include_trend,
                               threshold_str, date_str, show_all_dates,
                               progress):
    try:
        threshold = float(threshold_str)

        output_file = filedialog.asksaveasfilename(
            defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')]
        )
        if not output_file:
            root.after(0, lambda: [progress.stop(), progress.destroy()])
            return

        # ── Load & merge CVE + RMM ────────────────────────────────────────────
        df_vuln   = load_vulnerability_data(vuln_path)
        df_rmm    = None if skip_rmm else load_rmm_data(rmm_path)
        merged_df = merge_data(df_vuln, df_rmm, skip_rmm)

        raw_df         = merged_df.copy()
        stale_excluded = pd.DataFrame()

        if not show_all_dates:
            cutoff = pd.to_datetime(date_str)
            high   = merged_df[merged_df['Vulnerability Score'] >= threshold]
            stale_excluded = high[
                (high['_Sort_Time'] < cutoff) &
                (high['Last Response'] != 'Not Found in RMM')
            ].copy()
            merged_df = merged_df[
                (merged_df['_Sort_Time'] >= cutoff) |
                (merged_df['Last Response'] == 'Not Found in RMM')
            ]

        if merged_df.empty:
            root.after(0, lambda: [
                progress.stop(), progress.destroy(),
                messagebox.showwarning('No Data', 'No vulnerability records found after applying filters.'),
            ])
            return

        filtered_df = merged_df[merged_df['Vulnerability Score'] >= threshold].copy()
        triage_df   = filtered_df[filtered_df['Last Response'] != 'Not Found in RMM'].copy()

        # Sheet name for the main detections tab: "April Detections", "March Detections", etc.
        overview_sheet_name = datetime.now().strftime('%B') + ' Detections'

        reserved = {
            'trend summary', overview_sheet_name.lower(), 'all detections', 'raw data',
            'stale excluded devices', 'new this month', 'resolved', 'persisting cves',
            'patch match overview', 'patch match full data', 'patch report (full)',
        }
        used_names       = set(reserved)
        product_to_sheet = {}
        for product, _ in triage_df.groupby('Base Product'):
            product_to_sheet[product] = clean_sheet_name(product, used_names)

        # ── Optional: patch match (pre-compute before writer opens) ───────────
        patch_data = None
        if include_patch and patch_path:
            p_ov, p_full, p_raw, tot_r, filt_r = process_patch_match(
                patch_path, merged_df.copy(), min_score=threshold)
            patch_data = (p_ov, p_full, p_raw, tot_r, filt_r)

        # ── Optional: trend comparison ────────────────────────────────────────
        trend_data       = None
        prev_report_name = ''
        if include_trend and prev_report_path:
            prev_df          = load_previous_report(prev_report_path)
            prev_report_name = prev_report_path.replace('\\', '/').split('/')[-1]
            trend_data       = compute_trends(merged_df, prev_df, threshold)

        # ── Extract client name from CVE data ─────────────────────────────────
        customer_name = ''
        for col in ('Customer', 'Customer Name', 'Client', 'Client Name'):
            if col in merged_df.columns:
                vals = merged_df[col].dropna().astype(str).str.strip()
                vals = vals[vals.str.len() > 0]
                if not vals.empty:
                    customer_name = vals.iloc[0]
                    break

        # ── Write workbook ────────────────────────────────────────────────────
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            wb = writer.book

            link_fmt   = wb.add_format({'font_color': 'blue', 'underline': True})
            header_fmt = wb.add_format({'bold': True, 'font_size': 12,
                                        'bg_color': '#D9D9D9', 'border': 1})
            miss_fmt   = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            # Sheet order: Trend Summary → [Month] Detections → All Detections →
            #              New This Month → product tabs → Resolved → Persisting →
            #              stale → raw → patch
            if trend_data:
                build_trend_summary_sheet(wb, trend_data, threshold,
                                          prev_report_name, header_fmt,
                                          customer_name=customer_name)

            # ── Build patch-resolved lookup (needed by both Detections and product sheets)
            patch_resolved_pairs = set()
            no_patch_pairs       = set()
            if patch_data:
                p_full = patch_data[1]
                for _, row in p_full.iterrows():
                    nk = normalize_device_name(str(row.get('Name', '')))
                    ck = extract_cve_id(str(row.get('Vulnerability Name', '')))
                    res = str(row.get('Resolved (from Patch Report)', '')).strip()
                    pmr = str(row.get('Patch Match Result', '')).strip()
                    if res == 'Resolved':
                        patch_resolved_pairs.add((nk, ck))
                    elif pmr in ('Not found in patch report',
                                 'Device in patch report - product not found'):
                        no_patch_pairs.add((nk, ck))

            # ── Patch-confirmed count for Resolution Status ───────────────────
            patch_confirmed_count = 0
            if patch_resolved_pairs:
                triage_keys = set(zip(
                    triage_df['Name'].apply(normalize_device_name),
                    triage_df['Vulnerability Name'].apply(extract_cve_id),
                ))
                patch_confirmed_count = len(patch_resolved_pairs & triage_keys)

            # ── Re-detected after patch count ─────────────────────────────────
            redetected_count = 0
            if trend_data and prev_report_path:
                try:
                    xl_prev = pd.ExcelFile(prev_report_path)
                    if 'Resolved' in xl_prev.sheet_names:
                        prev_res = xl_prev.parse('Resolved')
                        if 'Vulnerability Name' in prev_res.columns:
                            prev_res_cves = set(
                                prev_res['Vulnerability Name'].apply(extract_cve_id)
                            )
                            cur_cves_scope = set(
                                triage_df['Vulnerability Name'].apply(extract_cve_id)
                            )
                            redetected_count = len(prev_res_cves & cur_cves_scope)
                except Exception:
                    pass

            build_overview_sheet(wb, merged_df, filtered_df, triage_df, threshold,
                                  product_to_sheet, header_fmt, link_fmt,
                                  customer_name=customer_name,
                                  patch_confirmed_count=patch_confirmed_count,
                                  redetected_count=redetected_count,
                                  sheet_name=overview_sheet_name,
                                  trend_metrics=trend_data['metrics'] if trend_data else None)

            # Sheet order: [Month] Detections → New This Month → Persisting CVEs →
            #              product tabs → Resolved → All Detections →
            #              Stale → Raw Data → Patch sheets
            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'New This Month', 'Persisting CVEs'})

            build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                                  patch_resolved_pairs=patch_resolved_pairs,
                                  no_patch_pairs=no_patch_pairs)

            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'Resolved'})

            build_all_detections_sheet(writer, merged_df, link_fmt, miss_fmt)

            if not stale_excluded.empty:
                build_stale_excluded_sheet(writer, stale_excluded)
            build_raw_data_sheet(writer, raw_df)

            if patch_data:
                build_patch_sheets(writer, patch_data[0], patch_data[1], patch_data[2])

        # ── Success message ───────────────────────────────────────────────────
        msg = f'Dashboard saved to:\n{output_file}'
        if trend_data:
            m = trend_data['metrics']
            msg += (f"\n\nTrend vs previous report:"
                    f"\n  ▲ {m['new_cve_count']:,} new CVE types   "
                    f"▼ {m['resolved_cve_count']:,} resolved   "
                    f"⏳ {m['persisting_cve_count']:,} persisting")
        if patch_data:
            msg += (f"\n\nPatch Match: "
                    f"{patch_data[3]:,} rows → {patch_data[4]:,} after score filter")

        root.after(0, lambda: [
            progress.stop(), progress.destroy(),
            messagebox.showinfo('Success', msg),
        ])

    except Exception as e:
        import traceback
        err = f'Processing failed:\n{e}\n\n{traceback.format_exc()}'
        root.after(0, lambda: [
            progress.stop(), progress.destroy(),
            messagebox.showerror('Error', err),
        ])


def process_reports():
    vuln_path        = vuln_var.get()
    rmm_path         = rmm_var.get()
    skip_rmm         = skip_rmm_var.get()
    include_patch    = include_patch_var.get()
    patch_path       = patch_var.get()
    include_trend    = include_trend_var.get()
    prev_report_path = prev_report_var.get()

    if not vuln_path:
        messagebox.showerror('Error', 'Please select the Vulnerability Report.'); return
    if not skip_rmm and not rmm_path:
        messagebox.showerror('Error', 'Please select the Device Inventory / RMM Report.'); return
    if include_patch and not patch_path:
        messagebox.showerror('Error', 'Patch Report matching is enabled but no file selected.\n'
                                       'Please browse for a Patch Report or uncheck the option.'); return
    if include_trend and not prev_report_path:
        messagebox.showerror('Error', 'Trend tracking is enabled but no previous report selected.\n'
                                       'Please browse for a previous dashboard or uncheck the option.'); return

    progress = ttk.Progressbar(root, mode='indeterminate')
    progress.pack(pady=5)
    progress.start()

    threading.Thread(
        target=execute_processing_thread,
        args=(vuln_path, rmm_path, skip_rmm,
              patch_path, include_patch,
              prev_report_path, include_trend,
              score_var.get(), date_var.get(), show_all_dates_var.get(),
              progress),
        daemon=True,
    ).start()


# ==============================================================================
# GUI TOGGLE HELPERS
# ==============================================================================

def toggle_rmm_state():
    if skip_rmm_var.get():
        rmm_entry.config(state=tk.DISABLED)
        rmm_button.config(state=tk.DISABLED)
        show_all_dates_var.set(True)
        show_all_dates_cb.config(state=tk.DISABLED)
        cal.config(state='disabled')
    else:
        rmm_entry.config(state=tk.NORMAL)
        rmm_button.config(state=tk.NORMAL)
        show_all_dates_cb.config(state=tk.NORMAL)
        toggle_date_state()

def toggle_date_state():
    cal.config(state='disabled' if show_all_dates_var.get() else 'normal')

def toggle_patch_state():
    s = tk.NORMAL if include_patch_var.get() else tk.DISABLED
    patch_entry.config(state=s)
    patch_button.config(state=s)

def toggle_trend_state():
    s = tk.NORMAL if include_trend_var.get() else tk.DISABLED
    prev_entry.config(state=s)
    prev_button.config(state=s)


# ==============================================================================
# GUI SETUP
# ==============================================================================

root = tk.Tk()
root.title('N-able CVE Dashboard & Triage Tool')
root.geometry('570x800')
root.resizable(False, True)

vuln_var           = tk.StringVar()
rmm_var            = tk.StringVar()
patch_var          = tk.StringVar()
prev_report_var    = tk.StringVar()
score_var          = tk.StringVar(value='9.0')
skip_rmm_var       = tk.BooleanVar(value=False)
include_patch_var  = tk.BooleanVar(value=False)
include_trend_var  = tk.BooleanVar(value=False)
date_var           = tk.StringVar()
show_all_dates_var = tk.BooleanVar(value=True)

pad = dict(pady=(5, 0))

# ── 1. Vulnerability Report ───────────────────────────────────────────────────
tk.Label(root, text='1. Vulnerability Report (CSV or XLSX)',
         font=('Arial', 10, 'bold')).pack(**pad)
tk.Entry(root, textvariable=vuln_var, width=68).pack()
tk.Button(root, text='Browse', command=lambda: select_file(vuln_var)).pack()

# ── 2. RMM / Device Inventory ─────────────────────────────────────────────────
tk.Label(root, text='2. Device Inventory / RMM Report (CSV or XLSX)',
         font=('Arial', 10, 'bold')).pack(**pad)
rmm_entry = tk.Entry(root, textvariable=rmm_var, width=68)
rmm_entry.pack()
rmm_button = tk.Button(root, text='Browse', command=lambda: select_file(rmm_var))
rmm_button.pack()
tk.Checkbutton(root,
               text="Skip Device Report (disables 'Last Response' and date filtering)",
               variable=skip_rmm_var, command=toggle_rmm_state).pack()

# ── 3. Patch Report (optional) ────────────────────────────────────────────────
patch_frame = tk.LabelFrame(root, text='3. Patch Report — Optional',
                             font=('Arial', 10, 'bold'), padx=8, pady=4)
patch_frame.pack(fill='x', padx=10, pady=(6, 2))
tk.Checkbutton(patch_frame,
               text="Include Patch Report matching (adds 'Patch Match' sheets to output)",
               variable=include_patch_var, command=toggle_patch_state).pack(anchor='w')
patch_entry = tk.Entry(patch_frame, textvariable=patch_var, width=62, state=tk.DISABLED)
patch_entry.pack(pady=(2, 0))
patch_button = tk.Button(patch_frame, text='Browse',
                          command=lambda: select_file(patch_var), state=tk.DISABLED)
patch_button.pack()

# ── 4. Previous Report / Trend Tracking (optional) ───────────────────────────
trend_frame = tk.LabelFrame(root,
                             text='4. Trend Tracking — Optional  (Month-over-Month)',
                             font=('Arial', 10, 'bold'), padx=8, pady=4)
trend_frame.pack(fill='x', padx=10, pady=(6, 2))
tk.Checkbutton(
    trend_frame,
    text=('Compare against a previous report\n'
          'Adds: Trend Summary · New This Month · Resolved · Persisting CVEs'),
    variable=include_trend_var, command=toggle_trend_state,
    justify='left',
).pack(anchor='w')
tk.Label(trend_frame,
         text='Browse for any dashboard previously generated by this tool (.xlsx)',
         fg='#555555', font=('Arial', 8)).pack(anchor='w')
prev_entry = tk.Entry(trend_frame, textvariable=prev_report_var, width=62, state=tk.DISABLED)
prev_entry.pack(pady=(2, 0))
prev_button = tk.Button(
    trend_frame, text='Browse',
    command=lambda: select_file(prev_report_var,
                                filetypes=[('Excel Dashboard', '*.xlsx'), ('All Files', '*.*')]),
    state=tk.DISABLED,
)
prev_button.pack()

# ── 5. Score Threshold ────────────────────────────────────────────────────────
tk.Label(root, text='5. Score Threshold (minimum score shown in product tabs)',
         font=('Arial', 10, 'bold')).pack(**pad)
tk.Entry(root, textvariable=score_var, width=10).pack()

# ── 6. RMM Check-in Cutoff Date ───────────────────────────────────────────────
tk.Label(root, text='6. RMM Check-in Cutoff Date',
         font=('Arial', 10, 'bold')).pack(**pad)
date_frame = tk.Frame(root)
date_frame.pack(pady=(2, 0))
cal = DateEntry(date_frame, selectmode='day', textvariable=date_var,
                date_pattern='yyyy-mm-dd', width=12)
cal.pack(side=tk.LEFT, padx=5)
show_all_dates_cb = tk.Checkbutton(date_frame, text='Show All Dates',
                                    variable=show_all_dates_var, command=toggle_date_state)
show_all_dates_cb.pack(side=tk.LEFT)
toggle_date_state()

# ── Generate ──────────────────────────────────────────────────────────────────
tk.Button(root, text='GENERATE COMPLETE DASHBOARD',
          command=process_reports,
          bg='#0078D7', fg='white',
          font=('Arial', 10, 'bold'), height=2).pack(pady=14)

root.mainloop()
