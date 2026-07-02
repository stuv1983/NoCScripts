"""
device_sheets.py — device/raw-data sheets: All Detections, Stale Excluded
Devices, CVEs on Stale Devices, Device Inventory, and Raw Data.

Note: build_all_detections_sheet and build_raw_data_sheet are currently
dead code — neither is imported by orchestrator.py.

Author : Stu Villanti <s.villanti@kenstra.com>
"""
import logging

import pandas as pd

from data_pipeline import extract_cve_id, get_col_letter, _drop_internal
from sheet_helpers import write_nvd_links as _write_nvd_links

log = logging.getLogger(__name__)

def build_all_detections_sheet(writer, merged_df, link_fmt, missing_row_fmt):
    df = _drop_internal(merged_df)
    df['NVD'] = ''

    cols = df.columns.tolist()
    if 'Device Type' in cols and 'Name' in cols:
        cols.insert(cols.index('Name') + 1, cols.pop(cols.index('Device Type')))
        df = df[cols]

    df = df.sort_values(by=['Vulnerability Score', 'Name'], ascending=[False, True])
    df.to_excel(writer, sheet_name='All Detections', index=False)

    ws = writer.sheets['All Detections']
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    cl = df.columns.tolist()

    if 'Vulnerability Name' in cl:
        vn_idx = cl.index('Vulnerability Name')
        ws.set_column(vn_idx, vn_idx, 25, link_fmt)
        # CVE text already written by to_excel(); set_column applies blue colour
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
    if 'N Days Exposed' in cl:
        nde_idx  = cl.index('N Days Exposed')
        nde_col  = get_col_letter(nde_idx)
        ws.set_column(nde_idx, nde_idx, 15)
        wb = writer.book
        # >180 days unpatched — critical (dark red)
        ws.conditional_format(1, nde_idx, len(df), nde_idx, {
            'type': 'cell', 'criteria': '>=', 'value': 180,
            'format': wb.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True}),
        })
        # 91–179 days — high (red)
        ws.conditional_format(1, nde_idx, len(df), nde_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 91, 'maximum': 179,
            'format': wb.add_format({'bg_color': '#FCE4D6'}),
        })
        # 31–90 days — amber
        ws.conditional_format(1, nde_idx, len(df), nde_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 31, 'maximum': 90,
            'format': wb.add_format({'bg_color': '#FFF2CC'}),
        })
        # 0–30 days — green (within acceptable window)
        ws.conditional_format(1, nde_idx, len(df), nde_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 0, 'maximum': 30,
            'format': wb.add_format({'bg_color': '#E2EFDA'}),
        })

def build_stale_excluded_sheet(writer, stale_df, not_in_rmm_df=None) -> None:
    """
    'Stale Excluded Devices' — one flat filterable table.
    Date-stale rows = amber, Not-Found-in-RMM rows = red highlight.
    A 'Reason' column distinguishes the two categories.
    """
    has_stale = stale_df is not None and not stale_df.empty
    has_nirm  = not_in_rmm_df is not None and not not_in_rmm_df.empty
    if not has_stale and not has_nirm:
        return

    cols_src = ['Name', 'Username', 'Last Response', 'Days Since Last Response', 'Device Type']
    wb = writer.book
    ws = wb.add_worksheet('Stale Excluded Devices')

    hdr_fmt  = wb.add_format({'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white', 'border': 1})
    row_stale= wb.add_format({'bg_color': '#FFFDE7', 'border': 1})
    row_nirm = wb.add_format({'bg_color': '#FFEBEE', 'font_color': '#9C0006', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})

    headers = ['Device Name', 'Username', 'Last Response', 'Days Since Last Response',
               'Device Type', 'Reason']
    col_widths = [35, 25, 25, 25, 18, 30]
    for ci, w in enumerate(col_widths):
        ws.set_column(ci, ci, w)

    # Build unified DataFrame
    frames = []
    if has_stale:
        _s = stale_df[[c for c in cols_src if c in stale_df.columns]].drop_duplicates(subset=['Name']).copy()
        _s['Reason'] = '⏱  Date-Stale'
        frames.append(_s)
    if has_nirm:
        _n = not_in_rmm_df[[c for c in cols_src if c in not_in_rmm_df.columns]].drop_duplicates(subset=['Name']).copy()
        _n['Reason'] = '🚫  Not Found in RMM'
        frames.append(_n)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.rename(columns={'Name': 'Device Name'})
    combined = combined.sort_values(['Reason', 'Last Response'] if 'Last Response' in combined.columns else ['Reason'])

    # Header row
    for ci, h in enumerate(headers):
        ws.write(0, ci, h, hdr_fmt)

    # Data rows — use positional access to avoid itertuples mangling column names with spaces
    col_positions = {col: i for i, col in enumerate(combined.columns)}
    for ri, row_vals in enumerate(combined.values.tolist(), start=1):
        _reason = str(row_vals[col_positions['Reason']]) if 'Reason' in col_positions else ''
        _fmt = row_nirm if 'Not Found' in _reason else row_stale
        for ci, h in enumerate(headers):
            src_col = h  # headers align with combined columns (Device Name, Username, etc.)
            pos = col_positions.get(src_col)
            val = row_vals[pos] if pos is not None else ''
            ws.write(ri, ci, str(val) if val is not None and not (isinstance(val, float) and (val != val)) else '', _fmt)

    # Autofilter on header row
    ws.autofilter(0, 0, len(combined), len(headers) - 1)

    note_row = len(combined) + 2
    ws.write(note_row, 0,
             'ℹ  Date-Stale: last seen before the cutoff — may still be live. '
             'Not-in-RMM (🚫 red): device absent from RMM inventory — '
             'verify decommission status (shadow IT / orphaned agent).', note_fmt)
    ws.set_row(note_row, 30)


def build_stale_cves_sheet(writer, df, link_fmt, not_in_rmm_cves_df=None) -> None:
    """
    'CVEs on Stale Devices' — one flat filterable table.
    Date-stale rows = light grey, Not-in-RMM rows = red.
    A 'Reason' column distinguishes the two; autofilter on the header.
    """
    has_stale = df is not None and not df.empty
    has_nirm  = not_in_rmm_cves_df is not None and not not_in_rmm_cves_df.empty
    if not has_stale and not has_nirm:
        return

    cols_src = ['Name', 'Username', 'Device Type', 'Vulnerability Name', 'Vulnerability Score',
                'Vulnerability Severity', 'Affected Products',
                'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response']
    headers  = cols_src + ['NVD', 'Reason']
    col_widths = {
        'Name': 25, 'Username': 22, 'Device Type': 15, 'Vulnerability Name': 25,
        'Vulnerability Score': 18, 'Vulnerability Severity': 20,
        'Affected Products': 30, 'Has Known Exploit': 16, 'CISA KEV': 12,
        'Last Response': 20, 'Days Since Last Response': 22, 'NVD': 10, 'Reason': 28,
    }

    wb = writer.book
    ws = wb.add_worksheet('CVEs on Stale Devices')

    hdr_fmt    = wb.add_format({'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white', 'border': 1})
    row_stale  = wb.add_format({'bg_color': '#F5F5F5', 'border': 1})
    row_nirm   = wb.add_format({'bg_color': '#FFEBEE', 'font_color': '#9C0006', 'border': 1})
    link_stale = wb.add_format({'bg_color': '#F5F5F5', 'border': 1, 'font_color': '#0563C1', 'underline': True})
    link_nirm  = wb.add_format({'bg_color': '#FFEBEE', 'border': 1, 'font_color': '#9C0006', 'underline': True})
    note_fmt   = wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})

    for ci, col_nm in enumerate(headers):
        ws.set_column(ci, ci, col_widths.get(col_nm, 15))

    # Build unified DataFrame
    frames = []
    if has_stale:
        _s = df[[c for c in cols_src if c in df.columns]].copy()
        _s['NVD'] = ''; _s['Reason'] = '⏱  Date-Stale'
        frames.append(_s)
    if has_nirm:
        _n = not_in_rmm_cves_df[[c for c in cols_src if c in not_in_rmm_cves_df.columns]].copy()
        _n['NVD'] = ''; _n['Reason'] = '🚫  Not Found in RMM'
        frames.append(_n)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values(
        by=['Reason', 'Name', 'Vulnerability Score'],
        ascending=[True, True, False]
    )
    cl = list(combined.columns)
    vn_idx  = cl.index('Vulnerability Name') if 'Vulnerability Name' in cl else None
    nvd_idx = headers.index('NVD')

    # Header
    for ci, h in enumerate(headers):
        ws.write(0, ci, h, hdr_fmt)

    # Data rows
    for ri, row in enumerate(combined.itertuples(index=False), start=1):
        _reason = str(row[-1]) if hasattr(row, '_fields') else ''
        _is_nirm = 'Not Found' in _reason
        _rfmt  = row_nirm  if _is_nirm else row_stale
        _lfmt  = link_nirm if _is_nirm else link_stale
        row_vals = list(row)
        for ci, col_nm in enumerate(headers):
            _ci_src = cl.index(col_nm) if col_nm in cl else None
            val = row_vals[_ci_src] if _ci_src is not None else ''
            safe = val if not (isinstance(val, float) and pd.isna(val)) else ''
            if col_nm == 'Vulnerability Name' and vn_idx is not None:
                ws.write(ri, ci, str(safe), _rfmt)
            elif col_nm == 'NVD':
                cve_val = row_vals[vn_idx] if vn_idx is not None else ''
                cve_id  = extract_cve_id(str(cve_val))
                ws.write(ri, nvd_idx, 'NVD ↗' if cve_id else '', _rfmt)
            else:
                ws.write(ri, ci, safe, _rfmt)

    # Autofilter on header row — works immediately on open
    ws.autofilter(0, 0, len(combined), len(headers) - 1)

    note_row = len(combined) + 2
    ws.write(note_row, 0,
             'ℹ  Date-Stale (grey): device excluded — Last Response before cutoff. '
             'Not-in-RMM (🚫 red): device absent from RMM inventory — '
             'verify decommission status (shadow IT / orphaned agent). '
             'Use the Reason filter to view each category separately.',
             note_fmt)
    ws.set_row(note_row, 36)


def build_device_report_sheet(writer, df_rmm: 'pd.DataFrame') -> None:
    """
    Device Inventory sheet — one row per device with check-in recency.
    Colour-codes Days Since Last Response so stale devices are immediately visible.
    """
    if df_rmm is None or df_rmm.empty:
        return

    wb  = writer.book
    hdr = wb.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1})
    grn = wb.add_format({'bg_color': '#E2EFDA'})
    amb = wb.add_format({'bg_color': '#FFF2CC', 'font_color': '#7F6000'})
    red = wb.add_format({'bg_color': '#FCE4D6'})
    crt = wb.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})

    # Select and rename columns for the sheet
    _col_map = {
        'Device':                   'Device Name',
        'Device Type':              'Type',
        'Last Response':            'Last Response',
        'Days Since Last Response': 'Days Since Last Response',
        'Username':                 'Username',
        'Site':                     'Site',
        'Client':                   'Client',
        'OS':                       'OS',
        'Description':              'Description',
    }
    _present = [c for c in _col_map if c in df_rmm.columns]
    out = df_rmm[_present].rename(columns=_col_map).copy()

    # Compute Days Since Last Response if not already present
    if 'Days Since Last Response' not in out.columns and 'Last Response' in out.columns:
        _lr = pd.to_datetime(out['Last Response'], errors='coerce')
        out['Days Since Last Response'] = (pd.Timestamp.now() - _lr).dt.days

    # Sort: most stale first
    if 'Days Since Last Response' in out.columns:
        out = out.sort_values('Days Since Last Response', ascending=False, na_position='last')

    out.to_excel(writer, sheet_name='Device Inventory', index=False)
    ws = writer.sheets['Device Inventory']
    ws.autofilter(0, 0, len(out), len(out.columns) - 1)
    ws.set_row(0, None, hdr)

    # Column widths
    _widths = {'Device Name': 30, 'Type': 13, 'Last Response': 22,
               'Days Since Last Response': 22, 'Username': 22,
               'Site': 18, 'Client': 18, 'OS': 30, 'Description': 30}
    for ci, col in enumerate(out.columns):
        ws.set_column(ci, ci, _widths.get(col, 16))

    # Colour-code Days Since Last Response
    if 'Days Since Last Response' in out.columns:
        _d_idx = out.columns.tolist().index('Days Since Last Response')
        ws.conditional_format(1, _d_idx, len(out), _d_idx, {
            'type': 'cell', 'criteria': '>=', 'value': 60, 'format': crt})  # >60 days: critical
        ws.conditional_format(1, _d_idx, len(out), _d_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 30, 'maximum': 59,
            'format': red})   # 30-59 days: stale
        ws.conditional_format(1, _d_idx, len(out), _d_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 14, 'maximum': 29,
            'format': amb})   # 14-29 days: approaching stale
        ws.conditional_format(1, _d_idx, len(out), _d_idx, {
            'type': 'cell', 'criteria': 'between', 'minimum': 0, 'maximum': 13,
            'format': grn})   # <14 days: healthy

    ws.write(len(out) + 2, 0,
             f'{len(out)} device(s).  '
             f'Green < 14 days  |  Amber 14-29 days  |  Red 30-59 days  |  Dark red ≥ 60 days.',
             note_fmt)
    log.debug('Device Inventory sheet written: %d devices', len(out))


def build_raw_data_sheet(writer, raw_df):
    df = _drop_internal(raw_df)
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    writer.sheets['Raw Data'].autofilter(0, 0, len(df), len(df.columns) - 1)