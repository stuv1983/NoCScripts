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
                          sheet_name='Detections', trend_metrics=None,
                          health_score: Optional[dict] = None,
                          evidence_summary: Optional[dict] = None):
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
        ('Patch required',                               alert_fmt, 'Patch Required',           'patch required'),
        ('Device missing from patch report',             warn_fmt,  'Missing from Patch Report', 'missing from report'),
        ('Patched but still detected (rescan required)', warn_fmt,  'Patched / Rescan Needed',   'rescan needed'),
        ('Patched but still vulnerable (rescan required)', warn_fmt, 'Patched / Rescan Needed',  'rescan needed'),
        ('Product not tracked',                          warn_fmt,  'Product Not Tracked',       'not tracked'),
        ('Installed but version unknown',                warn_fmt,  'Version Unknown',           'version unknown'),
        ('No patch baseline defined',                    info_fmt,  'No Baseline Defined',       'no baseline'),
    ]

    if evidence_summary:
        ws.merge_range(row_offset, 0, row_offset, 9,
                       'Patch Status Summary', header_fmt)
        ws.write(row_offset + 1, 0,
                 'Based on patch report correlation — indicates likely follow-up areas, '
                 'not confirmed root cause.', note_fmt)

        tile_col = 0
        for label, tile_colour, tile_title, sub_label in _TILE_ORDER:
            count = evidence_summary.get(label, 0)
            if count == 0:
                continue
            ws.merge_range(row_offset + 2, tile_col, row_offset + 2, tile_col + 1,
                           tile_title, tile_colour)
            ws.merge_range(row_offset + 3, tile_col, row_offset + 3, tile_col + 1,
                           count, count_fmt)
            ws.merge_range(row_offset + 4, tile_col, row_offset + 4, tile_col + 1,
                           sub_label, lbl_sm)
            tile_col += 2
            if tile_col > 8:
                break

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

    # ── Patch Reliability Score ───────────────────────────────────────────────
    if health_score and health_score.get('score') is not None:
        hs   = health_score['score']
        grad = health_score['grade']
        interp = health_score['interpretation']
        hs_colour = ('#375623' if hs >= 75 else '#7F6000' if hs >= 60 else '#9C0006')
        hs_fmt  = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': hs_colour})
        hs_note = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9})
        ws.write(r0,   7, 'Patch Reliability Score', header_fmt)
        ws.write(r0+1, 7, f'{hs} / 100  ({grad})', hs_fmt)
        ws.write(r0+2, 7, interp, hs_note)
        ws.set_column('H:H', 42)

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
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Affected Products', 'NVD']
    for product, group in triage_df.groupby('Base Product'):
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        group = group.sort_values(
            by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])

        def _resolved_value(row):
            nk = normalize_device_name(row['Name'])
            ck = extract_cve_id(row['Vulnerability Name'])
            return '☑' if (nk, ck) in patch_resolved_pairs else '☐'

        group.insert(0, 'Resolved', group.apply(_resolved_value, axis=1))
        group['NVD'] = ''

        final_cols = [c for c in cols_order if c in group.columns]
        group[final_cols].to_excel(writer, sheet_name=sheet_name, index=False)

        ws  = writer.sheets[sheet_name]
        wb_ = writer.book
        ws.autofilter(0, 0, len(group), len(final_cols) - 1)

        # Row formats — one per gap category + resolved + exploit
        patch_res_fmt     = wb_.add_format({'bg_color': '#DEEAF1'})  # blue   — patch via RMM
        exploit_fmt       = wb_.add_format({'bg_color': '#FFE0CC'})  # orange — known exploit
        coverage_fmt      = wb_.add_format({'bg_color': '#FFF2CC'})  # yellow — coverage_gap
        unmanaged_fmt     = wb_.add_format({'bg_color': '#FCE4D6'})  # peach  — unmanaged_app
        mismatch_fmt      = wb_.add_format({'bg_color': '#F2CEEF'})  # pink   — detection_mismatch

        _GAP_FMTS = {
            'coverage_gap':        coverage_fmt,
            'unmanaged_app':       unmanaged_fmt,
            'detection_mismatch':  mismatch_fmt,
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
            if (nk, ck) in patch_resolved_pairs:
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
        if 'Name'        in cl: ws.set_column(cl.index('Name'),        cl.index('Name'),        25)
        if 'Device Type' in cl: ws.set_column(cl.index('Device Type'), cl.index('Device Type'), 15)

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
            ('#FFFFFF', 'white row',  'Unresolved — patch available but not yet applied'),
        ]
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
                      'Resolved', 'Patch Evidence Notes', 'Recommended Steps']
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
        ws.set_column('G:G', 55)   # Recommended Steps
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
