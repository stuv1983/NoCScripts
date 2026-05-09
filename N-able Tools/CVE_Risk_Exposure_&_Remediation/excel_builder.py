"""
excel_builder.py — all xlsxwriter sheet-building functions.
No pandas data loading. No GUI. Receives DataFrames, writes sheets.
"""

import logging
from datetime import datetime
from typing import Optional, Set, Tuple, Dict
import re
import pandas as pd
from config import CVE_PATTERN, INSTALLED_STATUSES
from data_pipeline import (
    normalize_device_name, extract_cve_id, get_base_product,
    clean_sheet_name, _drop_internal, parse_last_response, get_col_letter,
)

log = logging.getLogger(__name__)

# EXCEL SHEET BUILDERS
# ==============================================================================

# ── Link helpers ──────────────────────────────────────────────────────────────
# IMPORTANT: do NOT use =HYPERLINK(...) formula strings via df.to_excel().
# xlsxwriter writes them as formulas with no cached result → openpyxl/pandas
# reads the cell back as 0.  Use write_url() instead: it stores the display
# text as the cached value so the cell is correctly readable by any reader.


# ==============================================================================
# WORKBOOK STYLE REGISTRY
# ==============================================================================

def get_workbook_styles(wb) -> dict:
    """
    Build all shared xlsxwriter Format objects once per workbook.

    Calling wb.add_format() is cheap but xlsxwriter tracks every format
    object for the workbook lifetime.  Creating the same format in multiple
    sheet-builder functions generates duplicate entries and makes colour
    inconsistencies easy to introduce.  A single registry call at workbook
    open time guarantees consistency and makes palette changes a one-liner.

    Usage:
        styles = get_workbook_styles(wb)
        ws.write(row, col, text, styles['header'])
    """
    return {
        # ── Title / header ────────────────────────────────────────────────────
        'title':        wb.add_format({'bold': True, 'font_size': 14,
                                       'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1}),
        'header':       wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#D9D9D9', 'border': 1}),
        'sub_header':   wb.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1}),
        'section':      wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1}),

        # ── Status / alert ────────────────────────────────────────────────────
        'alert':        wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#C00000', 'font_color': 'white'}),
        'warn':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#ED7D31', 'font_color': 'white'}),
        'info':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#375623', 'font_color': 'white'}),

        # ── Text / label ──────────────────────────────────────────────────────
        'bold':         wb.add_format({'bold': True}),
        'note':         wb.add_format({'italic': True, 'font_color': '#595959'}),
        'note_sm':      wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9}),
        'note_amber':   wb.add_format({'italic': True, 'font_color': '#7F6000', 'font_size': 8,
                                       'bg_color': '#FFFFE0', 'border': 1, 'text_wrap': True}),
        'link':         wb.add_format({'font_color': 'blue', 'underline': True}),

        # ── Trend colours ─────────────────────────────────────────────────────
        'up':           wb.add_format({'font_color': '#C00000', 'bold': True}),   # worse
        'down':         wb.add_format({'font_color': '#375623', 'bold': True}),   # better
        'same':         wb.add_format({'font_color': '#595959'}),

        # ── Row highlight ─────────────────────────────────────────────────────
        'row_red':      wb.add_format({'bg_color': '#FCE4D6'}),
        'row_green':    wb.add_format({'bg_color': '#E2EFDA'}),
        'row_amber':    wb.add_format({'bg_color': '#FFF2CC'}),
        'row_blue':     wb.add_format({'bg_color': '#DEEAF1'}),
        'row_pink':     wb.add_format({'bg_color': '#F2CEEF'}),
        'row_teal':     wb.add_format({'bg_color': '#D9F0F4'}),   # teal — patch installing
        'row_missing':  wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),

        # ── Health / score ────────────────────────────────────────────────────
        'score_good':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#375623'}),
        'score_warn':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#7F6000'}),
        'score_bad':    wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#9C0006'}),
    }

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
    row += 1; ws.write(row, 0, f'  ✅  Resolved          →  {m["resolved_cve_count"]} CVE types resolved (no longer detected or all devices ☑)')
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
        ('Resolved (Patch Confirmed)',        trend['resolved_df'],   res_bg,
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
                          sheet_name='Detections', trend_metrics=None,
                          evidence_summary: Optional[dict] = None,
                          recommended_actions: Optional[list] = None,
                          has_prev_report: bool = False):
    ws = workbook.add_worksheet(sheet_name)

    # ── Format definitions ────────────────────────────────────────────────────
    title_fmt = workbook.add_format({
        'bold': True, 'font_size': 14,
        'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1,
    })
    alert_fmt = workbook.add_format({
        'bold': True, 'font_size': 12,
        'bg_color': '#C00000', 'font_color': 'white',
    })
    warn_fmt = workbook.add_format({
        'bold': True, 'font_size': 12,
        'bg_color': '#ED7D31', 'font_color': 'white',
    })
    info_fmt = workbook.add_format({
        'bold': True, 'font_size': 12,
        'bg_color': '#375623', 'font_color': 'white',
    })
    count_fmt = workbook.add_format({'bold': True, 'font_size': 22, 'align': 'center'})
    lbl_sm    = workbook.add_format({'font_size': 9, 'align': 'center', 'text_wrap': True})
    note_fmt  = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})

    # ── Title row ─────────────────────────────────────────────────────────────
    title_text = (
        f'{customer_name}  —  CVE Risk Dashboard  (Score ≥ {threshold})  —  '
        f'{datetime.now().strftime("%B %Y")}' if customer_name else
        f'CVE Risk Dashboard  (Score ≥ {threshold})  —  {datetime.now().strftime("%B %Y")}'
    )
    ws.merge_range(0, 0, 0, 9, title_text, title_fmt)
    row_offset = 2

    # ── Patch Status Summary (dominant top section — this is what L1/management reads) ──
    # Five big tiles across the top. Impact-first order: worst → best.
    is_kev     = filtered_df['CISA KEV'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])
    is_exploit = filtered_df['Has Known Exploit'].astype(str).str.strip().str.lower().isin(['yes', 'true', '1', 'y'])

    kev_cves    = filtered_df[is_kev]['Vulnerability Name'].nunique()
    kev_devices = filtered_df[is_kev]['Name'].nunique()
    expl_cves   = filtered_df[is_exploit]['Vulnerability Name'].nunique()
    total_det   = filtered_df['Vulnerability Name'].nunique()
    uniq_dev    = filtered_df['Name'].nunique()
    avg_per_dev = round(total_det / uniq_dev, 1) if uniq_dev > 0 else 0
    total_srv   = merged_df[merged_df['Device Type'] == 'Server']['Name'].nunique()
    srv_aff     = filtered_df[filtered_df['Device Type'] == 'Server']['Name'].nunique()
    srv_pct     = f'{round((srv_aff / total_srv) * 100, 1)}%' if total_srv > 0 else '0%'

    missing_df      = filtered_df[filtered_df['Last Response'] == 'Not Found in RMM'].copy()
    missing_devices = sorted(missing_df['Name'].unique())

    # Build patch status counts from evidence_summary (if available)
    # Map display labels to their tile category
    _TILE_ORDER = [
        ('Patch required',                           alert_fmt, 'Patch Required'),
        ('Device missing from patch report',          warn_fmt,  'Missing from Patch Report'),
        ('Patched but still detected (rescan required)', warn_fmt, 'Patched / Rescan Needed'),
        ('Patched but still vulnerable (rescan required)', warn_fmt, 'Patched / Rescan Needed'),
        ('Product not tracked',                       warn_fmt,  'Product Not Tracked'),
        ('Installed but version unknown',             warn_fmt,  'Version Unknown'),
        ('No patch baseline defined',                 info_fmt,  'No Baseline Defined'),
    ]

    if evidence_summary:
        ws.merge_range(row_offset, 0, row_offset, 9,
                       'Patch Status Summary', header_fmt)
        ws.write(row_offset + 1, 0,
                 'Based on patch report correlation — indicates likely follow-up areas, '
                 'not confirmed root cause.', note_fmt)

        tile_col = 0
        for label, tile_colour, tile_title in _TILE_ORDER:
            count = evidence_summary.get(label, 0)
            if count == 0:
                continue
            ws.merge_range(row_offset + 2, tile_col, row_offset + 2, tile_col + 1,
                           tile_title, tile_colour)
            ws.merge_range(row_offset + 3, tile_col, row_offset + 3, tile_col + 1,
                           count, count_fmt)
            ws.merge_range(row_offset + 4, tile_col, row_offset + 4, tile_col + 1,
                           'devices', lbl_sm)
            tile_col += 2
            if tile_col > 8:
                break

        # Recommended Actions immediately below tiles
        if recommended_actions:
            act_row = row_offset + 6
            act_hdr_fmt = workbook.add_format({
                'bold': True, 'font_size': 11,
                'bg_color': '#2E4057', 'font_color': 'white', 'border': 1,
            })
            act_num_fmt = workbook.add_format({'bold': True, 'font_color': '#2E4057'})
            act_txt_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
            ws.merge_range(act_row, 0, act_row, 9, 'Recommended Actions', act_hdr_fmt)
            for i, act in enumerate(recommended_actions, start=1):
                r = act_row + i
                ws.write(r, 0, f'{i}.', act_num_fmt)
                ws.merge_range(r, 1, r, 7, act['action'], act_txt_fmt)
                ws.write(r, 8, act['count'], workbook.add_format({'align': 'center', 'bold': True}))
                ws.write(r, 9, 'devices', lbl_sm)
                ws.set_row(r, 28)
            row_offset = act_row + len(recommended_actions) + 2
        else:
            row_offset = row_offset + 8
    else:
        row_offset = row_offset  # no patch data — keep at 2

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

    # ── Patch Evidence Summary (plain counts — client-safe) ─────────────────
    # The synthetic Patch Reliability Score (0-100) uses internally-defined
    # weighted penalties that are not standards-based.  For client-facing
    # reports, plain counts are clearer and more defensible.
    if evidence_summary:
        summ_fmt  = workbook.add_format({'bold': True, 'font_size': 10})
        summ_note = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})
        ws.write(r0,   7, 'Patch Evidence Summary', header_fmt)
        sr = r0 + 1
        for label, count in sorted(evidence_summary.items(), key=lambda x: -x[1]):
            ws.write(sr, 7, f'{count}  {label}', summ_fmt)
            sr += 1
        ws.write(sr, 7, 'See Patch Evidence Notes sheet for per-device detail', summ_note)
        ws.set_column('H:H', 48)

    # ── N-able Patch Report note — Pending ≠ Installed ───────────────────────
    # This note addresses a common source of confusion: the "Discovered / Install
    # Date" column in the patch report has different meanings depending on Status.
    # Surfacing it here prevents "why isn't this resolved?" questions.
    if evidence_summary:
        pending_note_fmt = workbook.add_format({
            'italic': True, 'font_color': '#7F6000', 'font_size': 8,
            'bg_color': '#FFFFE0', 'border': 1, 'text_wrap': True,
        })
        pr = sr + 2
        ws.merge_range(pr, 7, pr + 2, 9,
            'N-able Patch Report note: '
            'For Status = Pending, the "Discovered / Install Date" is the date the patch was '
            'detected as available — not the date it was installed. '
            'Pending rows are not treated as remediated. '
            'Only Status = Installed or Reboot Required is accepted as patch evidence.',
            pending_note_fmt,
        )
        ws.set_row(pr, 14)
        ws.set_row(pr + 1, 14)
        ws.set_row(pr + 2, 14)

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
    # Compute unique pair counts directly from triage_df and patch data.
    # COUNTIF across product sheets inflates counts when one CVE affects
    # multiple products on the same device (it appears in each product sheet
    # and is counted once per sheet, not once per unique device-CVE pair).
    # Using the pre-computed sets gives the correct deduplicated totals.
    from data_pipeline import normalize_device_name, extract_cve_id
    triage_pairs_set = set(zip(
        triage_df['Name'].apply(normalize_device_name),
        triage_df['Vulnerability Name'].apply(extract_cve_id),
    ))
    n_total    = len(triage_pairs_set)
    # patch_confirmed_count is already deduplicated (intersection with triage)
    n_resolved = patch_confirmed_count  # patch-confirmed ☑ (unique pairs)
    n_unres    = n_total - n_resolved

    # Keep COUNTIF formula for live user editing feedback (shown alongside)
    if product_to_sheet:
        f_res   = ' + '.join([f"COUNTIF('{s}'!A:A, \"☑\")" for s in product_to_sheet.values()])
        f_unres = ' + '.join([f"COUNTIF('{s}'!A:A, \"☐\")" for s in product_to_sheet.values()])
    else:
        f_res, f_unres = '0', '0'

    ws.write(row_r + 1, 4, 'Resolved (☑)');
    ws.write(row_r + 1, 5, n_resolved)
    ws.write(row_r + 1, 6, 'patch-confirmed unique device × CVE pairs', note_fmt_small)
    ws.write(row_r + 2, 4, 'Unresolved (☐)');
    ws.write(row_r + 2, 5, n_unres)
    ws.write(row_r + 3, 4, 'Total (device × CVE pairs)');
    ws.write(row_r + 3, 5, n_total)
    ws.write(row_r + 3, 6, f'— {triage_df["Name"].nunique()} unique devices,  '
                            f'{triage_df["Vulnerability Name"].nunique()} unique CVE types', note_fmt_small)
    # Live checkbox tally (may count cross-product duplicates — use as tracking aid)
    if product_to_sheet:
        ws.write(row_r + 4, 4, '   ☑ ticked in sheets', sub_grey)
        ws.write_formula(row_r + 4, 5, f'={f_res}')
        ws.write(row_r + 4, 6, 'incl. cross-product duplicates — use Total above for reporting', note_fmt_small)

    extra_rows = 4   # rows used so far
    if patch_confirmed_count > 0:
        ws.write(row_r + 5, 4, '── Resolved breakdown ──', sub_grey)
        ws.write(row_r + 6, 4, '  Patch via RMM', sub_grey)
        ws.write(row_r + 6, 5, patch_confirmed_count)
        ws.write(row_r + 6, 6, 'pre-filled ☑ by patch report', note_fmt_small)
        extra_rows = 6

        if has_prev_report:
            ws.write(row_r + 7, 4, '  Manually Marked', sub_grey)
            ws.write_formula(row_r + 7, 5, f'={f_res} - {patch_confirmed_count}')
            ws.write(row_r + 7, 6, 'user-checked ☑', note_fmt_small)
            extra_rows = 7

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
                          patch_resolved_pairs=None,
                          patch_gap_pairs: Optional[Dict[Tuple[str, str], str]] = None):
    """
    patch_resolved_pairs: (device, cve) pairs confirmed resolved via patch report → blue
    patch_gap_pairs:      (device, cve) → gap_reason string, from classify_patch_gap():
                            'coverage_gap'   → yellow  (device absent from patch report)
                            'unmanaged_app'  → amber   (device present, product untracked)
                            'detection_mismatch' → pink (scanner/patch tool discrepancy)
    Known-exploit rows get orange. Priority: blue > orange > yellow/amber/pink > white.
    """
    if patch_resolved_pairs is None:
        patch_resolved_pairs = set()
    if patch_gap_pairs is None:
        patch_gap_pairs = {}

    cols_order = ['Resolved', 'Vulnerability Name', 'Name', 'Device Type',
                  'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index',
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Affected Products',
                  'Baseline Compliance', 'NVD']
    for product, group in triage_df.groupby('Base Product'):
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        group = group.sort_values(
            by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])

        # Derive canonical product key for this sheet's product so the resolved
        # lookup uses (device, cve, canonical_product) instead of (device, cve).
        # This prevents Edge patch evidence bleeding into Chrome rows for the
        # same CVE — a cross-product false-positive resolved tick.
        from data_pipeline import _detect_product as _dp_detect_prod
        # Use raw Affected Products values (not the base product name) so that
        # 'Microsoft Office 365' maps to canonical key 'office365', not 'office'.
        # get_base_product() strips '365' as if it were a version suffix, which
        # breaks the lookup for any product whose name ends in a number.
        _raw_pnames = group['Affected Products'].dropna().astype(str).unique().tolist()
        _sheet_pk = ''
        for _rpn in _raw_pnames:
            _pk_candidate = _dp_detect_prod(_rpn)
            if _pk_candidate:
                _sheet_pk = _pk_candidate
                break
        if not _sheet_pk:
            _sheet_pk = _dp_detect_prod(str(product))

        def _resolved_value(row):
            nk = normalize_device_name(row['Name'])
            ck = extract_cve_id(row['Vulnerability Name'])
            # Check product-scoped key first; fall back to product-free key
            # for backwards compatibility when patch_resolved_pairs was built
            # without product (e.g. loaded from an older dashboard workbook).
            if (nk, ck, _sheet_pk) in patch_resolved_pairs:
                return '☑'
            # If any 3-tuple key exists for (nk, ck) — new format present
            # If only 2-tuple keys exist — old format, use old check
            if patch_resolved_pairs and len(next(iter(patch_resolved_pairs))) == 2:
                return '☑' if (nk, ck) in patch_resolved_pairs else '☐'
            return '☐'

        group.insert(0, 'Resolved', group.apply(_resolved_value, axis=1))
        group['NVD'] = ''

        final_cols = [c for c in cols_order if c in group.columns]
        group[final_cols].to_excel(writer, sheet_name=sheet_name, index=False)

        ws  = writer.sheets[sheet_name]
        wb_ = writer.book
        ws.autofilter(0, 0, len(group), len(final_cols) - 1)

        # Row formats — use style registry to avoid per-sheet format recreation
        styles_           = get_workbook_styles(wb_)
        patch_res_fmt     = styles_['row_blue']
        exploit_fmt       = wb_.add_format({'bg_color': '#FFE0CC'})  # orange — no registry entry
        coverage_fmt      = styles_['row_amber']
        unmanaged_fmt     = styles_['row_red']
        mismatch_fmt      = styles_['row_pink']
        installing_fmt    = styles_['row_teal']

        _GAP_FMTS = {
            'coverage_gap':        coverage_fmt,
            'unmanaged_app':       unmanaged_fmt,
            'detection_mismatch':  mismatch_fmt,
            'patch_installing':    installing_fmt,
        }

        cl = final_cols
        if 'Resolved' in cl:
            ri = cl.index('Resolved')
            ws.data_validation(1, ri, len(group), ri, {'validate': 'list', 'source': ['☐', '☑']})
            ws.set_column(ri, ri, 10)

        # Row highlights — priority: blue > orange > gap-category > white
        _TRUE_VALS = {'yes', 'true', '1', 'y'}
        for row_i, (_, row) in enumerate(group[final_cols].iterrows(), start=1):
            nk = normalize_device_name(str(row.get('Name', '')))
            ck = extract_cve_id(str(row.get('Vulnerability Name', '')))
            _is_resolved = (
                (nk, ck, _sheet_pk) in patch_resolved_pairs
                or (len(patch_resolved_pairs) > 0
                    and len(next(iter(patch_resolved_pairs))) == 2
                    and (nk, ck) in patch_resolved_pairs)
            )
            if _is_resolved:
                ws.set_row(row_i, None, patch_res_fmt)
            elif str(row.get('Has Known Exploit', '')).strip().lower() in _TRUE_VALS:
                ws.set_row(row_i, None, exploit_fmt)
            else:
                gap = patch_gap_pairs.get((nk, ck))
                if gap and gap in _GAP_FMTS:
                    ws.set_row(row_i, None, _GAP_FMTS[gap])

        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25, link_fmt)
            _write_cve_links(ws, group['Vulnerability Name'], vn_idx, link_fmt)
        if 'NVD' in cl:
            nvd_idx = cl.index('NVD')
            ws.set_column(nvd_idx, nvd_idx, 10, link_fmt)
            _write_nvd_links(ws, group['Vulnerability Name'], nvd_idx, link_fmt)
        if 'Name'               in cl: ws.set_column(cl.index('Name'),               cl.index('Name'),               25)
        if 'Device Type'        in cl: ws.set_column(cl.index('Device Type'),        cl.index('Device Type'),        15)
        if 'Baseline Compliance' in cl: ws.set_column(cl.index('Baseline Compliance'), cl.index('Baseline Compliance'), 22)

        log.debug("Sheet '%s': %d rows written", sheet_name, len(group))

        # ── Legend ───────────────────────────────────────────────────────────
        legend_row = len(group) + 3
        l_title = wb_.add_format({'bold': True, 'font_size': 9, 'bg_color': '#F2F2F2', 'border': 1})
        l_cell  = wb_.add_format({'font_size': 9, 'border': 1})

        legend_entries = [
            ('#DEEAF1', 'blue row',   'Patch via RMM — install confirmed after CVE first detected'),
            ('#FFE0CC', 'orange row', 'Known active exploit — unresolved, prioritise immediately'),
            ('#FFF2CC', 'yellow row', 'Coverage gap — device not in patch report'),
            ('#FCE4D6', 'peach row',  'Unmanaged app — product not tracked in patch report'),
            ('#F2CEEF', 'pink row',   'Detection mismatch — CVE detected but no matching patch found'),
            ('#D9F0F4', 'teal row',   'Patch installing — patch is in progress, re-check after next RMM sync'),
            ('#FFFFFF', 'white row',  'Unresolved — patch available but not yet applied'),
        ]
        # Baseline note below legend
        ws.write(legend_row + len(legend_entries) + 2, 0,
                 'ℹ  Baseline Compliance column: shows whether the installed version meets the '
                 'current rolling product baseline (_baseline in config.json), '
                 'independently of CVE-specific patch status.',
                 wb_.add_format({'italic': True, 'font_color': '#595959', 'font_size': 8}))
        ws.write(legend_row, 0, 'Legend', l_title)
        for i, (colour, label, desc) in enumerate(legend_entries, start=1):
            fmt = wb_.add_format({'bg_color': colour, 'font_size': 9, 'border': 1})
            ws.write(legend_row + i, 0, f'  ({label})', fmt)
            ws.write(legend_row + i, 1, desc, l_cell)
            ws.set_row(legend_row + i, None, fmt)

def build_diagnostics_sheets(writer, diagnostics: dict) -> None:
    """
    Write diagnostic sheets: Patch Evidence Notes, Patch Lag, Version Drift.
    Skips sheets silently if DataFrame is empty.
    """
    wb = writer.book
    red  = wb.add_format({'bg_color': '#FCE4D6'})
    amb  = wb.add_format({'bg_color': '#FFF2CC'})
    grn  = wb.add_format({'bg_color': '#E2EFDA'})
    note = wb.add_format({'italic': True, 'font_color': '#595959'})

    # Colour per display label
    _LABEL_COLOUR = {
        'Patch required':                '#FCE4D6',   # red
        'Installed but still detected':  '#FCE4D6',   # red
        'No patch evidence':             '#FFF2CC',   # amber
        'Product not tracked':           '#FFF2CC',   # amber
        'No patch baseline defined':     '#FFF2CC',   # amber
        'Installed but version unknown': '#DEEAF1',   # blue
    }

    # ── Patch Evidence Notes ──────────────────────────────────────────────────
    rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
    if not rc_df.empty:
        # Only write columns intended for stakeholder view
        _SHOW_COLS = ['Device', 'Product', 'CVE', 'Patch Match Result',
                      'Resolved', 'Patch Evidence Notes', 'Baseline Compliance',
                      'Recommended Steps']
        out = rc_df[[c for c in _SHOW_COLS if c in rc_df.columns]].copy()
        out.to_excel(writer, sheet_name='Patch Evidence Notes', index=False)
        ws = writer.sheets['Patch Evidence Notes']
        ws.autofilter(0, 0, len(out), len(out.columns) - 1)
        ws.set_column('A:A', 28)   # Device
        ws.set_column('B:B', 30)   # Product
        ws.set_column('C:C', 20)   # CVE
        ws.set_column('D:D', 35)   # Patch Match Result
        ws.set_column('E:E', 12)   # Resolved
        ws.set_column('F:F', 32)   # Patch Evidence Notes
        ws.set_column('G:G', 22)   # Baseline Compliance
        ws.set_column('H:H', 55)   # Recommended Steps
        for i, label in enumerate(out.get('Patch Evidence Notes', []), start=1):
            colour = _LABEL_COLOUR.get(str(label), '#FFFFFF')
            ws.set_row(i, 30, wb.add_format({'bg_color': colour, 'text_wrap': True, 'valign': 'top'}))
        ws.write(len(out) + 2, 0,
                 'Patch Evidence Notes indicate likely follow-up areas based on CVE and '
                 'patch report correlation — not confirmed root cause.', note)
        log.debug("Patch Evidence Notes sheet: %d rows", len(out))

    # ── Patch Lag ─────────────────────────────────────────────────────────────
    lag_df = diagnostics.get('patch_lag_df', pd.DataFrame())
    if not lag_df.empty:
        lag_df.to_excel(writer, sheet_name='Patch Lag', index=False)
        ws = writer.sheets['Patch Lag']
        ws.autofilter(0, 0, len(lag_df), len(lag_df.columns) - 1)
        ws.set_column('A:A', 28); ws.set_column('B:B', 18); ws.set_column('C:C', 32)
        ws.set_column('F:F', 12)
        for i, lag in enumerate(lag_df.get('Lag (days)', []), start=1):
            fmt = red if lag is not None and (lag < 0 or lag > 60) else \
                  amb if lag is not None and lag > 14 else grn
            ws.set_row(i, None, fmt)
        ws.write(len(lag_df) + 2, 0,
                 'Negative lag = patch installed before CVE was first detected.', note)

    # ── Version Drift ─────────────────────────────────────────────────────────
    drift_df = diagnostics.get('version_drift_df', pd.DataFrame())
    if not drift_df.empty:
        drift_df.to_excel(writer, sheet_name='Version Drift', index=False)
        ws = writer.sheets['Version Drift']
        ws.autofilter(0, 0, len(drift_df), len(drift_df.columns) - 1)
        ws.set_column('A:A', 36); ws.set_column('C:C', 60)
        for i, spread in enumerate(drift_df.get('Distinct Versions', []), start=1):
            ws.set_row(i, None, red if spread >= 4 else amb if spread >= 2 else grn)
        ws.write(len(drift_df) + 2, 0,
                 'High distinct-version count = inconsistent update cadence across fleet.', note)
        # Note products with no version data (patch tool not tracking them)
        no_data = diagnostics.get('version_drift_no_data', [])
        if no_data:
            ws.write(len(drift_df) + 4, 0,
                     f'ℹ  No version data for: {", ".join(no_data)} — '
                     f'these products are not returning version numbers from the patch tool. '
                     f'Version drift cannot be assessed until they are tracked.',
                     wb.add_format({'italic': True, 'font_color': '#7F6000', 'text_wrap': True}))
    else:
        # Still create the sheet with an explanation
        ws = writer.book.add_worksheet('Version Drift')
        no_data = diagnostics.get('version_drift_no_data', [])
        if no_data:
            ws.write(0, 0,
                     f'No version data available for: {", ".join(no_data)}. '
                     f'These products are detected by N-able but the patch tool is not '
                     f'returning installed version numbers — they may not be in your patch '
                     f'policy scope. Version drift cannot be assessed.',
                     wb.add_format({'italic': True, 'font_color': '#7F6000', 'text_wrap': True}))
            ws.set_column('A:A', 80)
            ws.set_row(0, 50)


def build_patch_resolved_sheet(writer, patch_full_df: 'pd.DataFrame') -> None:
    """
    CVEs confirmed resolved via the patch report — patch installed, version
    compliant, install date post-dates first detection.

    Available on every run where a patch report is loaded. Does not require
    a previous dashboard (unlike the trend-based 'Resolved (Patch Confirmed)' sheet).
    Written to the 'Patch Confirmed' sheet to avoid naming collision with the trend sheet.
    """
    import pandas as pd

    resolved = patch_full_df[
        patch_full_df['Patch Evidence Status'] == 'Patch confirmed - pending rescan'
    ].copy()

    if resolved.empty:
        return

    wb       = writer.book
    grn      = wb.add_format({'bg_color': '#E2EFDA'})
    hdr      = wb.add_format({'bold': True, 'bg_color': '#375623',
                               'font_color': 'white', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})

    # Compute lag days
    if 'Patch Install Date' in resolved.columns and 'First detected' in resolved.columns:
        idt = pd.to_datetime(resolved['Patch Install Date'], errors='coerce')
        fdt = pd.to_datetime(resolved['First detected'],    errors='coerce')
        resolved['Lag (days)'] = (idt - fdt).dt.days
    else:
        resolved['Lag (days)'] = ''

    cols = [c for c in [
        'Name', 'Vulnerability Name', 'Affected Products',
        'Vulnerability Score', 'Matched Patch Version',
        'Patch Install Date', 'First detected', 'Lag (days)',
        'Product Baseline', 'Baseline Compliance',
    ] if c in resolved.columns]

    out = (resolved[cols]
           .drop_duplicates(subset=['Name', 'Vulnerability Name'])
           .sort_values(['Affected Products', 'Vulnerability Score'],
                        ascending=[True, False])
           .reset_index(drop=True))

    out.to_excel(writer, sheet_name='Resolved (Patch Confirmed)', index=False)
    ws = writer.sheets['Resolved (Patch Confirmed)']
    ws.autofilter(0, 0, len(out), len(out.columns) - 1)
    ws.set_row(0, None, hdr)

    ws.set_column('A:A', 28)   # Device
    ws.set_column('B:B', 22)   # CVE
    ws.set_column('C:C', 32)   # Product
    ws.set_column('D:D', 10)   # Score
    ws.set_column('E:E', 22)   # Version
    ws.set_column('F:G', 20)   # Dates
    ws.set_column('H:H', 12)   # Lag
    ws.set_column('I:I', 20)   # Product Baseline
    ws.set_column('J:J', 22)   # Baseline Compliance

    for i in range(1, len(out) + 1):
        ws.set_row(i, None, grn)

    # Summary note
    note_row = len(out) + 2
    unique_cves     = out['Vulnerability Name'].nunique()
    unique_devices  = out['Name'].nunique()
    ws.write(note_row, 0,
             f'{unique_cves} CVE type(s) resolved across {unique_devices} device(s) '
             f'via patch report. Install date confirmed after first detection date.',
             note_fmt)
    ws.merge_range(note_row, 0, note_row, len(out.columns) - 1,
                   f'{unique_cves} CVE type(s) confirmed patched across {unique_devices} '
                   f'device(s) via patch report. Install date confirmed after first detection date.',
                   note_fmt)

    log.debug("Patch Confirmed sheet: %d rows, %d CVEs, %d devices",
              len(out), unique_cves, unique_devices)


def build_products_not_tracked_sheet(writer,
                                      patch_full_df: 'pd.DataFrame') -> None:
    """
    Products detected by N-able with unresolved CVEs where the device IS in the
    patch report but this specific product is not being patched/tracked on it.

    This means the product needs to be added to the RMM patch policy for those
    devices — not a config.json issue, a patch scope issue.

    Also flags products not in config.json product_map at all (no canonical key),
    with a suggested entry to add.
    """
    import pandas as pd, re
    from data_pipeline import get_base_product, _detect_product, _norm_text

    wb       = writer.book
    red      = wb.add_format({'bg_color': '#FCE4D6'})
    amb      = wb.add_format({'bg_color': '#FFF2CC'})
    hdr      = wb.add_format({'bold': True, 'bg_color': '#1F4E79',
                               'font_color': 'white', 'border': 1})
    code_fmt = wb.add_format({'font_name': 'Courier New', 'font_size': 9,
                               'bg_color': '#F2F2F2'})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959',
                               'text_wrap': True})

    unmanaged = patch_full_df[
        patch_full_df['Patch Match Result'] == 'Device in patch report - product not found'
    ].copy()

    if unmanaged.empty:
        return

    unmanaged['_bp'] = unmanaged['Affected Products'].apply(get_base_product)
    unmanaged['_pk'] = unmanaged['Affected Products'].apply(
        lambda v: _detect_product(_norm_text(str(v))))

    # Aggregate per base product
    agg = (unmanaged.groupby(['_bp', '_pk'])
           .agg(
               devices       = ('Name',               'nunique'),
               cves          = ('Vulnerability Name', 'nunique'),
               sample_names  = ('Name', lambda x: ', '.join(sorted(x.unique())[:3])
                                         + (' ...' if x.nunique() > 3 else '')),
           )
           .reset_index()
           .sort_values('devices', ascending=False)
           .reset_index(drop=True))

    # Build suggested config.json entries
    def _suggest_entry(bp, pk):
        bp_clean = re.sub(r'\s+\d[\d.]+\s*$', '', str(bp).lower().strip())
        key      = pk if pk else bp_clean.replace(' ', '_')
        return f'["{bp_clean}", "{key}"]'

    agg['In product_map'] = agg['_pk'].apply(lambda v: '✓' if v else '✗')
    agg['Suggested config.json entry'] = agg.apply(
        lambda r: _suggest_entry(r['_bp'], r['_pk']), axis=1)

    out = agg.rename(columns={
        '_bp':          'Product (as detected by N-able)',
        '_pk':          'Internal Key',
        'devices':      'Devices Affected',
        'cves':         'CVE Count',
        'sample_names': 'Sample Devices',
    })[['Product (as detected by N-able)', 'Devices Affected', 'CVE Count',
        'Sample Devices', 'In product_map', 'Suggested config.json entry']]

    out.to_excel(writer, sheet_name='Products Not in Patch Scope', index=False)
    ws = writer.sheets['Products Not in Patch Scope']
    ws.autofilter(0, 0, len(out), len(out.columns) - 1)
    ws.set_row(0, None, hdr)

    ws.set_column('A:A', 40)   # Product name
    ws.set_column('B:B', 16)   # Devices
    ws.set_column('C:C', 11)   # CVEs
    ws.set_column('D:D', 45)   # Sample devices
    ws.set_column('E:E', 14)   # In product_map
    ws.set_column('F:F', 45)   # Suggested entry

    for i, row in enumerate(out.itertuples(), start=1):
        n = row._2  # Devices Affected
        ws.set_row(i, None, red if n >= 10 else amb)
        # Code-style formatting for the suggested entry column
        ws.write(i, 5, row._6, code_fmt)

    note_row = len(out) + 2
    ws.merge_range(note_row, 0, note_row, 5,
                   'These products are detected by N-able on devices that ARE in the patch report, '
                   'but this specific product is not included in the RMM patch policy for those devices. '
                   'To fix: add the product to your RMM patch policy scope. '
                   'If the product is also missing from config.json (✗ in "In product_map"), '
                   'add the suggested entry to config.json product_map as well.',
                   note_fmt)
    ws.set_row(note_row, 50)

    log.debug("Products Not in Patch Scope sheet: %d products", len(out))


def build_patch_failure_sheet(writer, failure_df: 'pd.DataFrame',
                              failure_lookup: dict,
                              cve_device_overlap: 'pd.DataFrame',
                              inventory_devices: 'set | None' = None) -> None:
    """
    inventory_devices: normalised device names from the RMM inventory.
    If supplied, devices absent from the inventory are excluded — they are
    decommissioned and patch failures on them are not actionable.
    """
    import pandas as pd
    wb  = writer.book
    red = wb.add_format({'bg_color': '#FCE4D6'})
    amb = wb.add_format({'bg_color': '#FFF2CC'})
    grn = wb.add_format({'bg_color': '#E2EFDA'})
    hdr = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})

    # ── Sheet 1: Device failure summary ──────────────────────────────────────
    # Filter to active (inventoried) devices only — decommissioned devices
    # appear in patch reports but are not actionable.
    active_lookup = failure_lookup
    excluded_count = 0
    if inventory_devices:
        active_lookup  = {d: info for d, info in failure_lookup.items()
                          if d in inventory_devices}
        excluded_count = len(failure_lookup) - len(active_lookup)
        if excluded_count:
            log.info("Patch Failures: excluded %d decommissioned device(s) "
                     "not in Device Inventory", excluded_count)

    rows = []
    for device, info in sorted(active_lookup.items(),
                               key=lambda x: -x[1]['failure_count']):
        rows.append({
            'Device':               device,
            'Total Failures':       info['failure_count'],
            'Unique KBs Failing':   info['unique_kbs'],
            'Primary Failure Type': info['top_category'].replace('_', ' ').title(),
            'Description':          info['top_description'],
            'All Categories':       ', '.join(f"{k.replace('_',' ').title()}: {v}"
                                             for k, v in info['categories'].items()),
        })
    if not rows:
        return

    summary_df = pd.DataFrame(rows)
    summary_df.to_excel(writer, sheet_name='Patch Failures', index=False)
    ws = writer.sheets['Patch Failures']
    ws.autofilter(0, 0, len(summary_df), len(summary_df.columns) - 1)
    ws.set_column('A:A', 26); ws.set_column('D:D', 30)
    ws.set_column('E:E', 55); ws.set_column('F:F', 55)

    # Colour by failure count severity
    for i, row in enumerate(rows, start=1):
        fc = row['Total Failures']
        ws.set_row(i, None, red if fc >= 20 else amb if fc >= 5 else grn)

    # Category totals (active devices only)
    active_devices_set = set(active_lookup.keys())
    active_fail = failure_df[failure_df['_device_norm'].isin(active_devices_set)]
    cat_totals = active_fail['_failure_cat'].value_counts()
    ws.write(len(summary_df) + 2, 0, 'Failure category totals (active devices):', hdr)
    for i, (cat, count) in enumerate(cat_totals.items()):
        ws.write(len(summary_df) + 3 + i, 0, f'  {cat.replace("_"," ").title()}')
        ws.write(len(summary_df) + 3 + i, 1, count)

    if excluded_count:
        ws.write(len(summary_df) + 3 + len(cat_totals) + 1, 0,
                 f'  ℹ  {excluded_count} device(s) excluded — not in Device Inventory '
                 f'(decommissioned)',
                 wb.add_format({'italic': True, 'font_color': '#595959'}))

    # ── CVEs on failing devices ───────────────────────────────────────────────
    if not cve_device_overlap.empty:
        out_cols = [c for c in ['Name', 'Vulnerability Name', 'Vulnerability Score',
                                'Affected Products', 'Has Known Exploit']
                    if c in cve_device_overlap.columns]
        overlap = cve_device_overlap[out_cols].drop_duplicates().sort_values(
            'Vulnerability Score', ascending=False) if 'Vulnerability Score' in out_cols else cve_device_overlap[out_cols]
        overlap.to_excel(writer, sheet_name='CVEs on Failing Devices', index=False)
        ws2 = writer.sheets['CVEs on Failing Devices']
        ws2.autofilter(0, 0, len(overlap), len(overlap.columns) - 1)
        ws2.set_column('A:A', 26); ws2.set_column('B:B', 22); ws2.set_column('D:D', 32)
        ws2.write(len(overlap) + 2, 0,
                  'These CVEs are on devices where patches are actively failing. '
                  'Resolving the patch delivery issue (see Patch Failures sheet) '
                  'should also clear these CVEs.', note_fmt)
        log.debug("CVEs on Failing Devices sheet: %d rows", len(overlap))

    log.debug("Patch Failures sheet: %d devices", len(rows))


def build_stale_excluded_sheet(writer, stale_df) -> None:
    if stale_df.empty:
        return
    df = stale_df[['Name', 'Last Response', 'Device Type']]\
           .drop_duplicates(subset=['Name']).copy()
    df = df.sort_values('Last Response').rename(columns={'Name': 'Device Name'})
    df.to_excel(writer, sheet_name='Stale Excluded Devices', index=False)
    ws = writer.sheets['Stale Excluded Devices']
    ws.set_column('A:A', 35); ws.set_column('B:B', 25); ws.set_column('C:C', 20)
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def build_client_summary_sheet(workbook, filtered_df, trend_data=None, customer_name='',
                               cutoff_date=None, stale_excluded_df=None,
                               not_in_rmm_count=0, not_in_rmm_cve_count=0):
    """
    Client-facing summary sheet — first sheet in the workbook.

    Contains:
      - Key metrics table (total rows, unique CVEs, CVSS 9+ counts)
      - CVSS score split table
      - Vulnerability Distribution pie chart (CVSS score spread)
      - Patching Effort bar chart (detections by CVSS score)
      - Month-over-month patching progress table + line chart (when trend available)
    """
    ws = workbook.add_worksheet('Client Summary')

    # ── Formats ───────────────────────────────────────────────────────────────
    title_fmt  = workbook.add_format({'bold': True, 'font_size': 15,
                                       'bg_color': '#1F4E79', 'font_color': 'white',
                                       'border': 1, 'valign': 'vcenter'})
    hdr_fmt    = workbook.add_format({'bold': True, 'bg_color': '#2E75B6',
                                       'font_color': 'white', 'border': 1, 'align': 'center'})
    sect_fmt   = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0',
                                       'border': 1, 'font_size': 11})
    lbl_fmt    = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    val_fmt    = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})
    val_pct    = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.0%'})
    red_fmt    = workbook.add_format({'bold': True, 'font_color': 'white',
                                       'bg_color': '#C00000', 'border': 1,
                                       'align': 'right', 'num_format': '#,##0'})
    grn_fmt    = workbook.add_format({'bold': True, 'font_color': 'white',
                                       'bg_color': '#375623', 'border': 1,
                                       'align': 'right', 'num_format': '#,##0'})
    note_fmt   = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})
    trend_up   = workbook.add_format({'bold': True, 'font_color': '#375623', 'border': 1,
                                       'align': 'right'})
    trend_dn   = workbook.add_format({'bold': True, 'font_color': '#C00000', 'border': 1,
                                       'align': 'right'})
    trend_eq   = workbook.add_format({'font_color': '#595959', 'border': 1, 'align': 'right'})

    ws.set_column('A:A', 38)
    ws.set_column('B:C', 18)
    ws.set_column('D:D', 14)

    # ── Title ─────────────────────────────────────────────────────────────────
    title_text = (f'{customer_name}  —  ' if customer_name else '') + 'CVE Risk Exposure Summary'
    ws.merge_range('A1:D1', title_text, title_fmt)
    ws.set_row(0, 28)
    ws.write('A2', f'Generated: {datetime.now().strftime("%d %B %Y")}', note_fmt)

    # ── Compute metrics ───────────────────────────────────────────────────────
    total_rows      = len(filtered_df)
    unique_cves     = int(filtered_df['Vulnerability Name'].nunique()) if 'Vulnerability Name' in filtered_df.columns else 0
    unique_devices  = int(filtered_df['Name'].nunique()) if 'Name' in filtered_df.columns else 0

    score_col = 'Vulnerability Score' if 'Vulnerability Score' in filtered_df.columns else None
    crit_mask = filtered_df[score_col] >= 9.0 if score_col else pd.Series([True] * len(filtered_df))
    crit_rows = int(crit_mask.sum())
    crit_cves = int(filtered_df.loc[crit_mask, 'Vulnerability Name'].nunique()) if score_col and 'Vulnerability Name' in filtered_df.columns else unique_cves

    exploit_col = 'Has Known Exploit' if 'Has Known Exploit' in filtered_df.columns else None
    exploit_count = int((filtered_df[exploit_col] == True).sum()) if exploit_col else 0  # noqa: E712

    kev_col = 'CISA KEV' if 'CISA KEV' in filtered_df.columns else None
    kev_count = int((filtered_df[kev_col] == True).sum()) if kev_col else 0  # noqa: E712

    device_type_col = 'Device Type' if 'Device Type' in filtered_df.columns else None
    server_count = 0
    if device_type_col:
        server_mask = filtered_df[device_type_col].astype(str).str.lower().str.contains('server', na=False)
        server_count = int(filtered_df.loc[server_mask, 'Name'].nunique()) if 'Name' in filtered_df.columns else 0

    # ── Key Metrics table ─────────────────────────────────────────────────────
    row = 3
    ws.merge_range(row, 0, row, 3, '  Key Metrics  (active devices only — excludes stale / not in RMM)', sect_fmt)
    row += 1

    metrics = [
        ('Total CVE detection rows',          total_rows,     val_fmt),
        ('Unique CVE types detected',          unique_cves,    val_fmt),
        ('Unique devices affected',            unique_devices, val_fmt),
        ('Detections at CVSS 9.0+',            crit_rows,      red_fmt),
        ('Unique CVEs at CVSS 9.0+',           crit_cves,      red_fmt),
        ('Detections with known exploit',      exploit_count,  red_fmt if exploit_count else val_fmt),
        ('Detections on CISA KEV list',        kev_count,      red_fmt if kev_count else val_fmt),
        ('Servers with CVE detections',        server_count,   val_fmt),
    ]
    for label, value, fmt in metrics:
        ws.write(row, 0, label, lbl_fmt)
        ws.merge_range(row, 1, row, 3, value, fmt)
        row += 1

    # ── Resolution Status breakdown ──────────────────────────────────────────
    _status_col = ('Threat Status' if 'Threat Status' in filtered_df.columns
                   else 'Status'   if 'Status'        in filtered_df.columns
                   else None)
    row += 1
    if _status_col:
        _res_mask   = filtered_df[_status_col].astype(str).str.strip().str.upper() == 'RESOLVED'
        _unres_mask = ~_res_mask
        _res_rows   = int(_res_mask.sum())
        _unres_rows = int(_unres_mask.sum())
        _tot_rows   = _res_rows + _unres_rows
        _res_cves   = int(filtered_df.loc[_res_mask,   'Vulnerability Name'].nunique()) if 'Vulnerability Name' in filtered_df.columns else 0
        _unres_cves = int(filtered_df.loc[_unres_mask, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in filtered_df.columns else 0

        ws.merge_range(row, 0, row, 3, '  Resolution Status  (CVSS 9.0+, active devices only)', sect_fmt)
        row += 1
        ws.write(row, 0, 'Status',           hdr_fmt)
        ws.write(row, 1, 'Detection Rows',   hdr_fmt)
        ws.write(row, 2, '% of Total',       hdr_fmt)
        ws.write(row, 3, 'Unique CVE Types', hdr_fmt)
        row += 1
        _res_pct   = _res_rows   / _tot_rows if _tot_rows else 0
        _unres_pct = _unres_rows / _tot_rows if _tot_rows else 0
        ws.write(row, 0, 'Resolved',   lbl_fmt); ws.write(row, 1, _res_rows,   grn_fmt); ws.write(row, 2, _res_pct,   val_pct); ws.write(row, 3, _res_cves,   grn_fmt); row += 1
        ws.write(row, 0, 'Unresolved', lbl_fmt); ws.write(row, 1, _unres_rows, red_fmt); ws.write(row, 2, _unres_pct, val_pct); ws.write(row, 3, _unres_cves, red_fmt); row += 1
        ws.write(row, 0, 'Total',      lbl_fmt); ws.write(row, 1, _tot_rows,   val_fmt); ws.write(row, 2, 1.0,        val_pct); ws.write(row, 3, (filtered_df['Vulnerability Name'].nunique() if 'Vulnerability Name' in filtered_df.columns else 0), val_fmt); row += 1
        ws.merge_range(row, 0, row, 3,
                       f'ℹ  CVSS 9.0+: {_res_rows:,} resolved  vs  {_unres_rows:,} unresolved'
                       + (f'  |  {_res_cves:,} CVE types resolved  vs  {_unres_cves:,} unresolved' if _res_cves or _unres_cves else ''),
                       note_fmt)
        row += 1

    # ── Excluded Detections breakdown ─────────────────────────────────────────
    _stale_rows    = int(len(stale_excluded_df))  if stale_excluded_df  is not None and not stale_excluded_df.empty  else 0
    _stale_devs    = int(stale_excluded_df['Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Name' in stale_excluded_df.columns else 0
    _stale_cves    = int(stale_excluded_df['Vulnerability Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Vulnerability Name' in stale_excluded_df.columns else 0
    _cutoff_label  = cutoff_date if cutoff_date else 'N/A (all dates included)'
    _has_exclusions = _stale_rows > 0 or not_in_rmm_count > 0

    if _has_exclusions:
        row += 1
        ws.merge_range(row, 0, row, 3, '  Excluded from Report', sect_fmt)
        row += 1
        ws.write(row, 0, 'Exclusion Reason',          hdr_fmt)
        ws.write(row, 1, 'Devices',                   hdr_fmt)
        ws.write(row, 2, 'CVE Detection Rows',        hdr_fmt)
        ws.write(row, 3, 'Unique CVE Types',          hdr_fmt)
        row += 1

        _amber_fmt = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'right', 'num_format': '#,##0'})
        _amber_lbl = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'bold': True})

        if _stale_rows > 0:
            _stale_label = f'Last Response before {_cutoff_label}'
            ws.write(row, 0, _stale_label, _amber_lbl)
            ws.write(row, 1, _stale_devs,  _amber_fmt)
            ws.write(row, 2, _stale_rows,  _amber_fmt)
            ws.write(row, 3, _stale_cves,  _amber_fmt)
            row += 1

        if not_in_rmm_count > 0:
            ws.write(row, 0, 'Device not found in RMM inventory', _amber_lbl)
            ws.write(row, 1, not_in_rmm_count,     _amber_fmt)
            ws.write(row, 2, not_in_rmm_cve_count, _amber_fmt)
            ws.write(row, 3, '',                    _amber_fmt)
            row += 1

        ws.merge_range(row, 0, row, 3,
                       f'ℹ  Data filtered from {_cutoff_label}. '
                       'Excluded devices are listed in the "Stale Excluded Devices" sheet.',
                       note_fmt)
        row += 1

    # ── CVSS Score Split table ────────────────────────────────────────────────
    row += 1
    ws.merge_range(row, 0, row, 3, '  CVSS Score Split  (detection rows)', sect_fmt)
    row += 1
    ws.write(row, 0, 'CVSS Score', hdr_fmt)
    ws.write(row, 1, 'Detection Rows', hdr_fmt)
    ws.write(row, 2, '% of Total', hdr_fmt)
    ws.write(row, 3, 'Unique CVEs', hdr_fmt)
    row += 1

    score_split_start = row   # remember for chart data range

    if score_col:
        score_groups = (
            filtered_df.groupby(filtered_df[score_col].round(1))
            .agg(rows=('Vulnerability Name', 'count'),
                 cves=('Vulnerability Name', 'nunique'))
            .sort_index(ascending=False)
        )
        score_split_data = []
        for score_val, srow in score_groups.iterrows():
            pct = srow['rows'] / total_rows if total_rows else 0
            ws.write(row, 0, float(score_val), lbl_fmt)
            ws.write(row, 1, int(srow['rows']),  val_fmt)
            ws.write(row, 2, pct,                val_pct)
            ws.write(row, 3, int(srow['cves']),  val_fmt)
            score_split_data.append((float(score_val), int(srow['rows']), int(srow['cves'])))
            row += 1
    else:
        score_split_data = []

    score_split_end = row - 1  # last data row (inclusive)

    # ── MoM Patching Progress table ───────────────────────────────────────────
    mom_start_row = None
    mom_data = []

    if trend_data:
        m = trend_data['metrics']
        row += 1
        ws.merge_range(row, 0, row, 3, '  Month-over-Month Patching Progress', sect_fmt)
        row += 1
        mom_start_row = row

        mom_metrics = [
            ('CVE types resolved / patched',     m.get('resolved_cve_count', 0),   True),
            ('CVE types newly introduced',        m.get('new_cve_count', 0),        False),
            ('CVE types persisting (unpatched)',  m.get('persisting_cve_count', 0), False),
            ('Devices fully remediated',          m.get('remediated_devices', 0),   True),
            ('New devices with CVEs',             m.get('new_devices', 0),          False),
        ]

        ws.write(row, 0, 'Metric', hdr_fmt)
        ws.write(row, 1, 'Count', hdr_fmt)
        ws.write(row, 2, 'Direction', hdr_fmt)
        ws.write(row, 3, '', hdr_fmt)
        row += 1

        for label, value, good_if_nonzero in mom_metrics:
            if good_if_nonzero:
                v_fmt  = grn_fmt if value > 0 else val_fmt
                d_str  = f'▼  {value:,}  (improvement)' if value > 0 else '—  no change'
                d_fmt  = trend_up if value > 0 else trend_eq
            else:
                v_fmt  = red_fmt if value > 0 else val_fmt
                d_str  = f'▲  {value:,}  (increase)' if value > 0 else '—  no change'
                d_fmt  = trend_dn if value > 0 else trend_eq
            ws.write(row, 0, label, lbl_fmt)
            ws.write(row, 1, value, v_fmt)
            ws.merge_range(row, 2, row, 3, d_str, d_fmt)
            mom_data.append((label, value))
            row += 1

        mom_end_row = row - 1

    # ── Note ──────────────────────────────────────────────────────────────────
    row += 1
    ws.write(row, 0,
             'ℹ  All metrics exclude devices not found in RMM and devices outside the '
             'active date window. CVSS 9.0+ detections are highlighted in red.',
             note_fmt)

    # ══════════════════════════════════════════════════════════════════════════
    # CHARTS  (placed to the right of the data tables, starting col E)
    # ══════════════════════════════════════════════════════════════════════════

    if score_split_data and len(score_split_data) >= 2:

        # ── Chart 1: Vulnerability Distribution Pie ───────────────────────────
        pie = workbook.add_chart({'type': 'pie'})
        pie.add_series({
            'name':       'Detection Rows',
            'categories': ['Client Summary', score_split_start, 0, score_split_end, 0],
            'values':     ['Client Summary', score_split_start, 1, score_split_end, 1],
            'data_labels': {'percentage': True, 'category': True,
                            'font': {'size': 9}},
        })
        pie.set_title({'name': 'Vulnerability Distribution by CVSS Score'})
        pie.set_style(10)
        pie.set_size({'width': 380, 'height': 260})
        ws.insert_chart('F4', pie, {'x_offset': 5, 'y_offset': 5})

        # ── Chart 2: Patching Effort Bar ─────────────────────────────────────
        bar = workbook.add_chart({'type': 'bar'})
        bar.add_series({
            'name':       'Detection Rows',
            'categories': ['Client Summary', score_split_start, 0, score_split_end, 0],
            'values':     ['Client Summary', score_split_start, 1, score_split_end, 1],
            'fill':       {'color': '#2E75B6'},
            'data_labels': {'value': True, 'font': {'size': 9}},
        })
        bar.add_series({
            'name':       'Unique CVEs',
            'categories': ['Client Summary', score_split_start, 0, score_split_end, 0],
            'values':     ['Client Summary', score_split_start, 3, score_split_end, 3],
            'fill':       {'color': '#ED7D31'},
            'data_labels': {'value': True, 'font': {'size': 9}},
        })
        bar.set_title({'name': 'Patching Effort by CVSS Score'})
        bar.set_x_axis({'name': 'Count'})
        bar.set_y_axis({'name': 'CVSS Score'})
        bar.set_legend({'position': 'bottom'})
        bar.set_style(10)
        bar.set_size({'width': 380, 'height': 260})
        ws.insert_chart('F20', bar, {'x_offset': 5, 'y_offset': 5})

    # ── Chart 3: MoM Patching Progress Bar ───────────────────────────────────
    if trend_data and mom_data and mom_start_row is not None:
        mom_chart_start = mom_start_row + 1   # skip header row
        mom_chart_end   = mom_chart_start + len(mom_data) - 1

        mom_bar = workbook.add_chart({'type': 'bar'})
        mom_bar.add_series({
            'name':       'Count',
            'categories': ['Client Summary', mom_chart_start, 0, mom_chart_end, 0],
            'values':     ['Client Summary', mom_chart_start, 1, mom_chart_end, 1],
            'fill':       {'color': '#375623'},
            'data_labels': {'value': True, 'font': {'size': 9}},
        })
        mom_bar.set_title({'name': 'Month-over-Month Patching Progress'})
        mom_bar.set_x_axis({'name': 'Count'})
        mom_bar.set_legend({'none': True})
        mom_bar.set_style(10)
        mom_bar.set_size({'width': 380, 'height': 260})
        ws.insert_chart('F36', mom_bar, {'x_offset': 5, 'y_offset': 5})

    log.debug("Client Summary sheet written")


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