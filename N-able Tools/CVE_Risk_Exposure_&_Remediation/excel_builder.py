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

def get_workbook_styles(wb) -> dict:
    return {
        'title':        wb.add_format({'bold': True, 'font_size': 14,
                                       'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1}),
        'header':       wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#D9D9D9', 'border': 1}),
        'sub_header':   wb.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1}),
        'section':      wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1}),
        'alert':        wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#C00000', 'font_color': 'white'}),
        'warn':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#ED7D31', 'font_color': 'white'}),
        'info':         wb.add_format({'bold': True, 'font_size': 12,
                                       'bg_color': '#375623', 'font_color': 'white'}),
        'bold':         wb.add_format({'bold': True}),
        'note':         wb.add_format({'italic': True, 'font_color': '#595959'}),
        'note_sm':      wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9}),
        'note_amber':   wb.add_format({'italic': True, 'font_color': '#7F6000', 'font_size': 8,
                                       'bg_color': '#FFFFE0', 'border': 1, 'text_wrap': True}),
        'link':         wb.add_format({'font_color': 'blue', 'underline': True}),
        'up':           wb.add_format({'font_color': '#C00000', 'bold': True}), 
        'down':         wb.add_format({'font_color': '#375623', 'bold': True}), 
        'same':         wb.add_format({'font_color': '#595959'}),
        'row_red':      wb.add_format({'bg_color': '#FCE4D6'}),
        'row_green':    wb.add_format({'bg_color': '#E2EFDA'}),
        'row_amber':    wb.add_format({'bg_color': '#FFF2CC'}),
        'row_blue':     wb.add_format({'bg_color': '#DEEAF1'}),
        'row_pink':     wb.add_format({'bg_color': '#F2CEEF'}),
        'row_teal':     wb.add_format({'bg_color': '#D9F0F4'}), 
        'row_missing':  wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),
        'score_good':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#375623'}),
        'score_warn':   wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#7F6000'}),
        'score_bad':    wb.add_format({'bold': True, 'font_size': 18, 'font_color': '#9C0006'}),
    }

def _write_cve_links(ws, vuln_name_series, col_idx, link_fmt):
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
    ws = workbook.add_worksheet('Trend Summary')
    m  = trend['metrics']

    title_fmt = workbook.add_format({
        'bold': True, 'font_size': 14,
        'bg_color': '#1F4E79', 'font_color': 'white', 'border': 1,
    })
    sub_fmt   = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
    lbl_fmt   = workbook.add_format({'bold': True})
    up_fmt    = workbook.add_format({'font_color': '#C00000', 'bold': True}) 
    down_fmt  = workbook.add_format({'font_color': '#375623', 'bold': True})  
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
            ws.write(row, 1, int(prow['Previous']))
            ws.write(row, 2, int(prow['Current']))
            dv_str, dv_fmt = _ch(int(prow['Change']), prod_up_fmt, prod_dn_fmt, prod_eq_fmt)
            ws.write(row, 3, dv_str, dv_fmt)
            ws.write(row, 4, '')   
            ws.write(row, 5, int(prow['CVE_Previous']))
            ws.write(row, 6, int(prow['CVE_Current']))
            cv_str, cv_fmt = _ch(int(prow['CVE_Change']), prod_up_fmt, prod_dn_fmt, prod_eq_fmt)
            ws.write(row, 7, cv_str, cv_fmt)

    row += 2; ws.merge_range(row, 0, row, 3, '  Detail Sheets in This Workbook', sect_fmt)
    row += 1; ws.write(row, 0, f'  📋  New This Month    →  {m["new_cve_count"]} new CVE types × all affected devices')
    row += 1; ws.write(row, 0, f'  ⏳  Persisting CVEs   →  {m["persisting_cve_count"]} CVE types carried over from previous report')


# ── Trend Detail Sheets ───────────────────────────────────────────────────────

def build_trend_detail_sheets(writer, workbook, trend, link_fmt, sheets_subset=None):
    new_bg  = workbook.add_format({'bg_color': '#FCE4D6'})  
    per_bg  = workbook.add_format({'bg_color': '#FFF2CC'}) 

    detail_cols = ['Name', 'Device Type', 'Vulnerability Name', 'Vulnerability Score',
                   'Vulnerability Severity', 'Affected Products',
                   'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response']

    all_sheets = [
        ('New This Month',  trend['new_df'],        new_bg,
         'New CVEs not seen in the previous report — investigate and prioritise.'),
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
        df['NVD'] = ''

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
                          has_prev_report: bool = False,
                          stale_excluded_df: Optional[pd.DataFrame] = None,
                          report_month: str = ''):
    ws = workbook.add_worksheet(sheet_name)
    if not report_month:
        report_month = datetime.now().strftime("%B %Y")

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

    title_text = (
        f'{customer_name}  —  CVE Risk Dashboard  (Score ≥ {threshold})  —  {report_month}' if customer_name else
        f'CVE Risk Dashboard  (Score ≥ {threshold})  —  {report_month}'
    )
    ws.merge_range(0, 0, 0, 9, title_text, title_fmt)
    row_offset = 2

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
        row_offset = row_offset  

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

    if trend_metrics:
        m = trend_metrics
        ctx_row = r0 + 6
        ctx_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 11,
            'bg_color': '#2E4057', 'font_color': 'white', 'border': 1,
        })
        new_fmt  = workbook.add_format({'bold': True, 'font_color': '#C00000'}) 
        res_fmt  = workbook.add_format({'bold': True, 'font_color': '#375623'}) 
        per_fmt  = workbook.add_format({'bold': True, 'font_color': '#7F6000'})  
        note_ctx = workbook.add_format({'font_color': '#595959', 'italic': True, 'font_size': 9})

        nc = m['new_cve_count']
        rc = m['resolved_cve_count']
        pc = m['persisting_cve_count']
        prev_c = m['prev_cves']
        cur_c  = m['cur_cves']

        scope_delta_cur  = cur_c  - (nc + pc)
        scope_delta_prev = prev_c - (rc + pc)

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

    row_t = r0 + 7
    ws.write(row_t, 0, f'Unique CVEs by Severity (Score {threshold}+)', header_fmt)
    sev_counts = filtered_df.drop_duplicates(subset=['Vulnerability Name'])['Vulnerability Severity'].value_counts()
    r = row_t + 1
    for sev, cnt in sev_counts.items():
        ws.write(r, 0, str(sev)); ws.write(r, 1, cnt); r += 1

    row_p = max(r + 2, r0 + 14)
    hdr_small = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    ws.write(row_p, 0, f'Top 10 Products (Score {threshold}+)', header_fmt)
    ws.write(row_p, 1, 'Unique Devices', hdr_small)
    ws.write(row_p, 2, 'Unique CVE Types', hdr_small)

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
    dt_counts = filtered_df.groupby('Device Type', observed=True)['Name'].nunique()
    r2 = row_t + 1
    for dt, cnt in dt_counts.items():
        ws.write(r2, 4, str(dt)); ws.write(r2, 5, cnt); r2 += 1

    row_r = max(r2 + 2, r0 + 14)
    ws.write(row_r, 4, f'Resolution Status (Score {threshold}+)', header_fmt)
    sub_grey = workbook.add_format({'font_color': '#595959', 'indent': 1})
    note_fmt_small = workbook.add_format({'font_color': '#595959', 'italic': True, 'font_size': 9})
    grn_tile = workbook.add_format({'font_color': '#375623', 'bold': True})
    red_tile = workbook.add_format({'font_color': '#C00000', 'bold': True})

    # Use N-able Status column as source of truth — same as Client Summary.
    # Count unique (device, cve) pairs per status so these figures are
    # comparable to Client Summary rows but deduplicated across products.
    _ov_sc = ('Threat Status' if 'Threat Status' in triage_df.columns
              else 'Status'   if 'Status'        in triage_df.columns
              else None)
    if _ov_sc:
        _ov_res_rows = triage_df[triage_df[_ov_sc].astype(str).str.strip().str.upper() == 'RESOLVED']
        _ov_unr_rows = triage_df[triage_df[_ov_sc].astype(str).str.strip().str.upper() == 'UNRESOLVED']
        _ov_res_pairs = set(zip(_ov_res_rows['Name'], _ov_res_rows['Vulnerability Name']))
        _ov_unr_pairs = set(zip(_ov_unr_rows['Name'], _ov_unr_rows['Vulnerability Name']))
        _ov_all_pairs = set(zip(triage_df['Name'],    triage_df['Vulnerability Name']))
        n_res_pairs  = len(_ov_res_pairs)
        n_unr_pairs  = len(_ov_unr_pairs)
        n_total      = len(_ov_all_pairs)
        n_overlap    = len(_ov_res_pairs & _ov_unr_pairs)
    else:
        _ov_all_pairs = set(zip(triage_df['Name'], triage_df['Vulnerability Name']))
        n_res_pairs = n_unr_pairs = n_overlap = 0
        n_total = len(_ov_all_pairs)

    if product_to_sheet:
        f_res   = ' + '.join([f"COUNTIF('{s}'!A:A, \"☑\")" for s in product_to_sheet.values()])
        f_unres = ' + '.join([f"COUNTIF('{s}'!A:A, \"☐\")" for s in product_to_sheet.values()])
    else:
        f_res, f_unres = '0', '0'

    ws.write(row_r + 1, 4, 'Resolved')
    ws.write(row_r + 1, 5, n_res_pairs, grn_tile)
    ws.write(row_r + 1, 6, 'unique device × CVE pairs with Status = RESOLVED in N-able', note_fmt_small)
    ws.write(row_r + 2, 4, 'Unresolved')
    ws.write(row_r + 2, 5, n_unr_pairs, red_tile)
    ws.write(row_r + 2, 6, 'unique device × CVE pairs still showing UNRESOLVED in N-able', note_fmt_small)
    ws.write(row_r + 3, 4, 'Total unique pairs')
    ws.write(row_r + 3, 5, n_total)
    ws.write(row_r + 3, 6, f'— {triage_df["Name"].nunique()} devices, '
                            f'{triage_df["Vulnerability Name"].nunique()} CVE types', note_fmt_small)
    if n_overlap > 0:
        ws.write(row_r + 4, 4, f'  ↕ {n_overlap:,} pair(s) in both', sub_grey)
        ws.write(row_r + 4, 6,
                 'same CVE resolved on some devices, unresolved on others — resolved + unresolved > total is expected',
                 note_fmt_small)

    extra_rows = 4
    if patch_confirmed_count > 0:
        ws.write(row_r + 5, 4, '── Patch tool breakdown ──', sub_grey)
        ws.write(row_r + 6, 4, '  Patch-confirmed (☑ pre-filled)', sub_grey)
        ws.write(row_r + 6, 5, patch_confirmed_count)
        ws.write(row_r + 6, 6, 'unique device × CVE pairs confirmed via patch report', note_fmt_small)
        extra_rows = 6

        if has_prev_report:
            ws.write(row_r + 7, 4, '  ☑ ticked in sheets', sub_grey)
            ws.write_formula(row_r + 7, 5, f'={f_res}')
            ws.write(row_r + 7, 6, 'incl. cross-product duplicates — for reference only', note_fmt_small)
            ws.write(row_r + 8, 4, '  Manually marked', sub_grey)
            ws.write_formula(row_r + 8, 5, f'={f_res} - {patch_confirmed_count}')
            ws.write(row_r + 8, 6, 'user-checked ☑', note_fmt_small)
            extra_rows = 8

    if redetected_count > 0:
        rr = row_r + extra_rows + 1
        ws.write(rr, 4, '⚠ Re-detected After Patch')
        ws.write(rr, 5, redetected_count)
        ws.write(rr, 6, 'CVEs manually marked resolved last report but still present — investigate', note_fmt_small)
        extra_rows += 1

    row_m = row_r + extra_rows + 2
    stale_devs = stale_excluded_df['Name'].unique().tolist() if stale_excluded_df is not None else []
    
    ws.write(row_m, 4, f'Devices Not Found in RMM ({len(missing_devices)}) / Excluded Stale ({len(stale_devs)}) (Score {threshold}+)', header_fmt)
    ws.write(row_m, 5, 'Last Response', hdr_small)
    ws.write(row_m, 6, 'Days Since Last Response', hdr_small)

    mi = row_m + 1
    if not missing_devices and not stale_devs:
        ws.write(mi, 4, 'All devices synced and active')
    else:
        for dev in missing_devices:
            dev_rows = filtered_df[filtered_df['Name'] == dev]
            lr_vals  = dev_rows['Last Response'].dropna().unique()
            lr_val   = lr_vals[0] if len(lr_vals) else 'Not Found in RMM'
            
            days_vals = dev_rows['Days Since Last Response'].dropna().unique() if 'Days Since Last Response' in dev_rows.columns else []
            days_val = days_vals[0] if len(days_vals) else '—'

            ws.write(mi, 4, str(dev))
            ws.write(mi, 5, str(lr_val))
            ws.write(mi, 6, str(days_val))
            mi += 1
            
        for dev in stale_devs:
            dev_rows = stale_excluded_df[stale_excluded_df['Name'] == dev]
            lr_vals  = dev_rows['Last Response'].dropna().unique()
            lr_val   = lr_vals[0] if len(lr_vals) else '—'
            
            days_vals = dev_rows['Days Since Last Response'].dropna().unique() if 'Days Since Last Response' in dev_rows.columns else []
            days_val = days_vals[0] if len(days_vals) else '—'

            ws.write(mi, 4, f"{dev} (Stale)")
            ws.write(mi, 5, str(lr_val))
            ws.write(mi, 6, str(days_val))
            mi += 1

    ws.set_column('A:A', 38)
    ws.set_column('B:C', 14)
    ws.set_column('E:E', 48)
    ws.set_column('F:F', 22)
    ws.set_column('G:G', 24)

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
    if patch_resolved_pairs is None:
        patch_resolved_pairs = set()
    if patch_gap_pairs is None:
        patch_gap_pairs = {}

    cols_order = ['Resolved', 'Vulnerability Name', 'Name', 'Device Type',
                  'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index',
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response', 'Affected Products',
                  'Baseline Compliance', 'NVD']
    for product, group in triage_df.groupby('Base Product'):
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        group = group.sort_values(
            by=['Vulnerability Score', '_Sort_Time', 'Name'], ascending=[False, False, True])

        from data_pipeline import _detect_product as _dp_detect_prod
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
            if (nk, ck, _sheet_pk) in patch_resolved_pairs:
                return '☑'
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

        styles_           = get_workbook_styles(wb_)
        patch_res_fmt     = styles_['row_blue']
        exploit_fmt       = wb_.add_format({'bg_color': '#FFE0CC'})
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
    wb = writer.book
    red  = wb.add_format({'bg_color': '#FCE4D6'})
    amb  = wb.add_format({'bg_color': '#FFF2CC'})
    grn  = wb.add_format({'bg_color': '#E2EFDA'})
    note = wb.add_format({'italic': True, 'font_color': '#595959'})

    _LABEL_COLOUR = {
        'Patch required':                '#FCE4D6',
        'Installed but still detected':  '#FCE4D6',
        'No patch evidence':             '#FFF2CC',
        'Product not tracked':           '#FFF2CC',
        'No patch baseline defined':     '#FFF2CC',
        'Installed but version unknown': '#DEEAF1',
    }

    rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
    if not rc_df.empty:
        _SHOW_COLS = ['Device', 'Product', 'CVE', 'Patch Match Result',
                      'Resolved', 'Patch Evidence Notes', 'Baseline Compliance',
                      'Recommended Steps']
        out = rc_df[[c for c in _SHOW_COLS if c in rc_df.columns]].copy()
        out.to_excel(writer, sheet_name='Patch Evidence Notes', index=False)
        ws = writer.sheets['Patch Evidence Notes']
        ws.autofilter(0, 0, len(out), len(out.columns) - 1)
        ws.set_column('A:A', 28)
        ws.set_column('B:B', 30)
        ws.set_column('C:C', 20)
        ws.set_column('D:D', 35)
        ws.set_column('E:E', 12)
        ws.set_column('F:F', 32)
        ws.set_column('G:G', 22)
        ws.set_column('H:H', 55)
        for i, label in enumerate(out.get('Patch Evidence Notes', []), start=1):
            colour = _LABEL_COLOUR.get(str(label), '#FFFFFF')
            ws.set_row(i, 30, wb.add_format({'bg_color': colour, 'text_wrap': True, 'valign': 'top'}))
        ws.write(len(out) + 2, 0,
                 'Patch Evidence Notes indicate likely follow-up areas based on CVE and '
                 'patch report correlation — not confirmed root cause.', note)

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

    drift_df = diagnostics.get('version_drift_df', pd.DataFrame())
    if not drift_df.empty:
        drift_df.to_excel(writer, sheet_name='Version Drift', index=False)
        ws = writer.sheets['Version Drift']
        ws.autofilter(0, 0, len(drift_df), len(drift_df.columns) - 1)
        ws.set_column('A:A', 36); ws.set_column('C:C', 60)
        # Audit Note column (may or may not be present)
        _cols = list(drift_df.columns)
        if 'Audit Note' in _cols:
            _an_idx = _cols.index('Audit Note')
            ws.set_column(_an_idx, _an_idx, 50)
        for i, spread in enumerate(drift_df.get('Distinct Versions', []), start=1):
            ws.set_row(i, None, red if spread >= 4 else amb if spread >= 2 else grn)
        ws.write(len(drift_df) + 2, 0,
                 'High distinct-version count = inconsistent update cadence across fleet. '
                 'Audit Note: per-user/AppData installs bypass GPO — remove and replace with system-scope. '
                 '32-bit installs on 64-bit OS should be replaced.',
                 wb.add_format({'italic': True, 'font_color': '#595959', 'text_wrap': True}))
        ws.set_row(len(drift_df) + 2, 36)
        no_data = diagnostics.get('version_drift_no_data', [])
        if no_data:
            ws.write(len(drift_df) + 4, 0,
                     f'ℹ  No version data for: {", ".join(no_data)} — '
                     f'these products are not returning version numbers from the patch tool. '
                     f'Version drift cannot be assessed until they are tracked.',
                     wb.add_format({'italic': True, 'font_color': '#7F6000', 'text_wrap': True}))
    else:
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

    ws.set_column('A:A', 28)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 32)
    ws.set_column('D:D', 10)
    ws.set_column('E:E', 22)
    ws.set_column('F:G', 20)
    ws.set_column('H:H', 12)
    ws.set_column('I:I', 20)
    ws.set_column('J:J', 22)

    for i in range(1, len(out) + 1):
        ws.set_row(i, None, grn)

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


def build_products_not_tracked_sheet(writer,
                                      patch_full_df: 'pd.DataFrame') -> None:
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

    ws.set_column('A:A', 40)
    ws.set_column('B:B', 16)
    ws.set_column('C:C', 11)
    ws.set_column('D:D', 45)
    ws.set_column('E:E', 14)
    ws.set_column('F:F', 45)

    for i, row in enumerate(out.itertuples(), start=1):
        n = row._2
        ws.set_row(i, None, red if n >= 10 else amb)
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


def build_patch_failure_sheet(writer, failure_df: 'pd.DataFrame',
                              failure_lookup: dict,
                              cve_device_overlap: 'pd.DataFrame',
                              inventory_devices: 'set | None' = None) -> None:
    import pandas as pd
    wb  = writer.book
    red = wb.add_format({'bg_color': '#FCE4D6'})
    amb = wb.add_format({'bg_color': '#FFF2CC'})
    grn = wb.add_format({'bg_color': '#E2EFDA'})
    hdr = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
    hdr_red  = wb.add_format({'bold': True, 'bg_color': '#C00000', 'font_color': 'white', 'border': 1})
    note_fmt = wb.add_format({'italic': True, 'font_color': '#595959'})
    title_fmt= wb.add_format({'bold': True, 'font_size': 12, 'bg_color': '#1F4E79',
                               'font_color': 'white', 'border': 1})

    active_lookup = failure_lookup
    excluded_count = 0
    if inventory_devices:
        active_lookup  = {d: info for d, info in failure_lookup.items()
                          if d in inventory_devices}
        excluded_count = len(failure_lookup) - len(active_lookup)

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

    # ── Summary stats at top before the table ───────────────────────────────
    ws = wb.add_worksheet('Patch Failures')
    stat_fmt  = wb.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    stat_val  = wb.add_format({'border': 1, 'align': 'right'})

    total_failures  = sum(r['Total Failures']     for r in rows)
    total_devices   = len(rows)
    total_kbs       = sum(r['Unique KBs Failing'] for r in rows)
    active_fail_df  = failure_df[failure_df['_device_norm'].isin(set(active_lookup.keys()))]
    cat_totals      = active_fail_df['_failure_cat'].value_counts()
    top_cat_label   = cat_totals.index[0].replace('_', ' ').title() if not cat_totals.empty else '—'
    top_cat_count   = int(cat_totals.iloc[0]) if not cat_totals.empty else 0

    ws.merge_range(0, 0, 0, 5, 'Patch Failure Analysis', title_fmt)
    stats = [
        ('Devices with failures', total_devices),
        ('Total failure events',  total_failures),
        ('Distinct KBs failing',  total_kbs),
        ('Most common failure',   f'{top_cat_label} ({top_cat_count} events)'),
    ]
    if excluded_count:
        stats.append(('Excluded (not in inventory)', excluded_count))
    for si, (label, val) in enumerate(stats):
        ws.write(1 + si, 0, label, stat_fmt)
        ws.write(1 + si, 1, val,   stat_val)
    ws.set_column('A:A', 26); ws.set_column('D:D', 30)
    ws.set_column('E:E', 55); ws.set_column('F:F', 55)
    ws.set_column('B:C', 18)

    # ── Device failure table ─────────────────────────────────────────────────
    tbl_start = len(stats) + 3
    for ci, col in enumerate(summary_df.columns):
        ws.write(tbl_start, ci, col, hdr)
    for i, row in enumerate(rows, start=tbl_start + 1):
        fc = row['Total Failures']
        ws.set_row(i, None, red if fc >= 20 else amb if fc >= 5 else grn)
        for ci, col in enumerate(summary_df.columns):
            ws.write(i, ci, row[col])
    ws.autofilter(tbl_start, 0, tbl_start + len(rows), len(summary_df.columns) - 1)

    # ── Category totals below table ──────────────────────────────────────────
    note_start = tbl_start + len(rows) + 2
    ws.write(note_start, 0, 'Failure category totals (active devices):', hdr)
    for i, (cat, count) in enumerate(cat_totals.items()):
        ws.write(note_start + 1 + i, 0, f'  {cat.replace("_"," ").title()}')
        ws.write(note_start + 1 + i, 1, count)

    # ── CVEs on Failing Devices — enriched ──────────────────────────────────
    if not cve_device_overlap.empty:
        # Add primary failure type to each row from lookup
        _fail_info = {d: info for d, info in active_lookup.items()}
        _norm_name = cve_device_overlap['Name'].astype(str).apply(
            lambda n: n.strip().upper().split('\\')[-1].split('.')[0])

        cve_out = cve_device_overlap.copy()
        cve_out['_nk'] = _norm_name
        cve_out['Primary Failure Type'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['top_category'].replace('_', ' ').title()
                       if nk in _fail_info else '—'
        )
        cve_out['Total Device Failures'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['failure_count'] if nk in _fail_info else 0
        )
        cve_out['Failure Description'] = cve_out['_nk'].map(
            lambda nk: _fail_info[nk]['top_description'] if nk in _fail_info else '—'
        )
        cve_out = cve_out.drop(columns=['_nk'], errors='ignore')

        out_cols = [c for c in [
            'Name', 'Vulnerability Name', 'Vulnerability Score', 'Affected Products',
            'Has Known Exploit', 'Primary Failure Type', 'Total Device Failures',
            'Failure Description'
        ] if c in cve_out.columns]

        overlap = (cve_out[out_cols]
                   .drop_duplicates(subset=['Name', 'Vulnerability Name'])
                   .sort_values(['Total Device Failures', 'Vulnerability Score'],
                                ascending=[False, False])
                   .reset_index(drop=True))

        overlap.to_excel(writer, sheet_name='CVEs on Failing Devices', index=False)
        ws2 = writer.sheets['CVEs on Failing Devices']
        ws2.autofilter(0, 0, len(overlap), len(overlap.columns) - 1)
        ws2.set_column('A:A', 26); ws2.set_column('B:B', 22)
        ws2.set_column('D:D', 32); ws2.set_column('F:F', 24)
        ws2.set_column('G:G', 20); ws2.set_column('H:H', 55)
        ws2.set_row(0, None, hdr_red)

        # Colour rows by failure severity
        for i, row in enumerate(overlap.itertuples(index=False), start=1):
            fc = getattr(row, 'Total_Device_Failures', 0) or 0
            ws2.set_row(i, None, red if fc >= 20 else amb if fc >= 5 else grn)

        ws2.write(len(overlap) + 2, 0,
                  f'⚠  {len(overlap)} CVE detection(s) on {overlap["Name"].nunique()} device(s) '
                  f'where patches are actively failing. '
                  f'Resolving the delivery failure (Primary Failure Type) will unblock patching. '
                  f'See Patch Failures sheet for per-device remediation steps.', note_fmt)
        ws2.set_row(len(overlap) + 2, 50)

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

    # Data rows
    for ri, row in enumerate(combined.itertuples(index=False), start=1):
        _reason = str(row[-1]) if hasattr(row, '_fields') else ''
        _fmt = row_nirm if 'Not Found' in _reason else row_stale
        for ci, h in enumerate(headers):
            col_map = {'Device Name': 'Device Name', 'Username': 'Username',
                       'Last Response': 'Last Response',
                       'Days Since Last Response': 'Days Since Last Response',
                       'Device Type': 'Device Type', 'Reason': 'Reason'}
            try:
                val = getattr(row, col_map[h].replace(' ', '_').replace('(', '').replace(')', ''))
            except AttributeError:
                val = ''
            ws.write(ri, ci, str(val) if val is not None and not (isinstance(val, float) and pd.isna(val)) else '', _fmt)

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
                cve_id = extract_cve_id(str(safe))
                url = f'https://nvd.nist.gov/vuln/detail/{cve_id}' if cve_id else ''
                if url: ws.write_url(ri, ci, url, _lfmt, str(safe))
                else:   ws.write(ri, ci, str(safe), _rfmt)
            elif col_nm == 'NVD':
                cve_val = row_vals[vn_idx] if vn_idx is not None else ''
                cve_id  = extract_cve_id(str(cve_val))
                url = f'https://nvd.nist.gov/vuln/detail/{cve_id}' if cve_id else ''
                if url: ws.write_url(ri, nvd_idx, url, _lfmt, 'NVD ↗')
                else:   ws.write(ri, nvd_idx, '', _rfmt)
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


def build_client_summary_sheet(workbook, filtered_df, triage_df, threshold,
                               trend_data=None, customer_name='',
                               cutoff_date=None, stale_excluded_df=None,
                               not_in_rmm_count=0, not_in_rmm_cve_count=0,
                               not_in_rmm_unique_cves=0,
                               report_month=''):
    """
    Client Summary sheet.

    filtered_df  — score-filtered rows including not-in-RMM & stale (waterfall baseline).
    triage_df    — active scope only (stale + not-in-RMM removed). All Key Metrics use this.
    threshold    — CVSS floor shown in the waterfall header.
    """
    ws = workbook.add_worksheet('Client Summary')
    if not report_month:
        report_month = datetime.now().strftime("%B %Y")

    title_fmt = workbook.add_format({'bold': True, 'font_size': 15, 'bg_color': '#1F4E79',
                                      'font_color': 'white', 'border': 1, 'valign': 'vcenter'})
    hdr_fmt   = workbook.add_format({'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
                                      'border': 1, 'align': 'center'})
    sect_fmt  = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1, 'font_size': 11})
    lbl_fmt   = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    val_fmt   = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})
    val_pct   = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.0%'})
    red_fmt   = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#C00000',
                                      'border': 1, 'align': 'right', 'num_format': '#,##0'})
    grn_fmt   = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#375623',
                                      'border': 1, 'align': 'right', 'num_format': '#,##0'})
    note_fmt  = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9, 'text_wrap': True})
    trend_up  = workbook.add_format({'bold': True, 'font_color': '#375623', 'border': 1, 'align': 'right'})
    trend_dn  = workbook.add_format({'bold': True, 'font_color': '#C00000',  'border': 1, 'align': 'right'})
    trend_eq  = workbook.add_format({'font_color': '#595959', 'border': 1, 'align': 'right'})
    wf_plus   = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    wf_minus  = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1})
    wf_mval   = workbook.add_format({'font_color': '#C00000', 'bg_color': '#FFF2CC',
                                      'border': 1, 'align': 'right', 'num_format': '#,##0'})
    wf_eq_lbl = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1})
    wf_eq_val = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0', 'border': 1,
                                      'align': 'right', 'num_format': '#,##0'})

    ws.set_column('A:A', 44); ws.set_column('B:D', 18)
    title_text = (f'{customer_name}  \u2014  ' if customer_name else '') + 'CVE Risk Exposure Summary'
    ws.merge_range('A1:D1', title_text, title_fmt); ws.set_row(0, 28)
    ws.write('A2', f'Report Month: {report_month}  |  Generated: {datetime.now().strftime("%d %b %Y")}',
             workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9}))

    # Key Metrics — all from triage_df (active scope only)
    total_rows     = len(triage_df)
    unique_cves    = int(triage_df['Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_df.columns else 0
    unique_devices = int(triage_df['Name'].nunique())               if 'Name'               in triage_df.columns else 0
    score_col      = 'Vulnerability Score' if 'Vulnerability Score' in triage_df.columns else None
    crit_mask      = triage_df[score_col] >= 9.0 if score_col else pd.Series([True]*len(triage_df))
    crit_rows      = int(crit_mask.sum())
    crit_cves      = int(triage_df.loc[crit_mask,'Vulnerability Name'].nunique()) if score_col and 'Vulnerability Name' in triage_df.columns else unique_cves
    exploit_col    = 'Has Known Exploit' if 'Has Known Exploit' in triage_df.columns else None
    exploit_count  = int((triage_df[exploit_col]==True).sum()) if exploit_col else 0
    kev_col        = 'CISA KEV' if 'CISA KEV' in triage_df.columns else None
    kev_count      = int((triage_df[kev_col]==True).sum()) if kev_col else 0
    server_count   = 0
    if 'Device Type' in triage_df.columns and 'Name' in triage_df.columns:
        srv_mask     = triage_df['Device Type'].astype(str).str.lower().str.contains('server', na=False)
        server_count = int(triage_df.loc[srv_mask,'Name'].nunique())

    row = 3
    ws.merge_range(row,0,row,3,'  Key Metrics  (active devices only \u2014 excludes stale / not in RMM)',sect_fmt); row+=1
    for label, value, fmt in [
        ('Total CVE detection rows',        total_rows,    val_fmt),
        ('Unique CVE types detected',        unique_cves,   val_fmt),
        ('Unique devices affected',          unique_devices,val_fmt),
        ('Detections at CVSS 9.0+',          crit_rows,     red_fmt),
        ('Unique CVEs at CVSS 9.0+',         crit_cves,     red_fmt),
        ('Detections with known exploit',    exploit_count, red_fmt if exploit_count else val_fmt),
        ('Detections on CISA KEV list',      kev_count,     red_fmt if kev_count     else val_fmt),
        ('Servers with CVE detections',      server_count,  val_fmt),
    ]:
        ws.write(row,0,label,lbl_fmt); ws.merge_range(row,1,row,3,value,fmt); row+=1

    # Resolution Status
    _sc = ('Threat Status' if 'Threat Status' in triage_df.columns
           else 'Status'   if 'Status'        in triage_df.columns else None)
    row+=1
    if _sc:
        _res = triage_df[_sc].astype(str).str.strip().str.upper()=='RESOLVED'
        _unr = ~_res
        _rr  = int(_res.sum()); _ur = int(_unr.sum()); _tot = _rr+_ur
        _rc  = int(triage_df.loc[_res,'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_df.columns else 0
        _uc  = int(triage_df.loc[_unr,'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_df.columns else 0
        ws.merge_range(row,0,row,3,'  Resolution Status  (active devices only)',sect_fmt); row+=1
        ws.write(row,0,'Status',hdr_fmt); ws.write(row,1,'Detection Rows',hdr_fmt)
        ws.write(row,2,'% of Total',hdr_fmt); ws.write(row,3,'Unique CVE Types',hdr_fmt); row+=1
        _rp=_rr/_tot if _tot else 0; _up=_ur/_tot if _tot else 0
        ws.write(row,0,'Resolved',  lbl_fmt); ws.write(row,1,_rr,grn_fmt); ws.write(row,2,_rp,val_pct); ws.write(row,3,_rc,grn_fmt); row+=1
        ws.write(row,0,'Unresolved',lbl_fmt); ws.write(row,1,_ur,red_fmt); ws.write(row,2,_up,val_pct); ws.write(row,3,_uc,red_fmt); row+=1
        ws.write(row,0,'Total',     lbl_fmt); ws.write(row,1,_tot,val_fmt); ws.write(row,2,1.0,val_pct)
        ws.write(row,3,triage_df['Vulnerability Name'].nunique() if 'Vulnerability Name' in triage_df.columns else 0,val_fmt); row+=1
        ws.merge_range(row,0,row,3,
                       f'\u2139  CVSS 9.0+: {_rr:,} resolved  vs  {_ur:,} unresolved'
                       +(f'  |  {_rc:,} CVE types resolved  vs  {_uc:,} unresolved' if _rc or _uc else ''),note_fmt); row+=1

    # Data Filtering Reconciliation waterfall
    _stale_rows = int(len(stale_excluded_df)) if stale_excluded_df is not None and not stale_excluded_df.empty else 0
    _stale_devs = int(stale_excluded_df['Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Name' in stale_excluded_df.columns else 0
    _stale_cves = int(stale_excluded_df['Vulnerability Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Vulnerability Name' in stale_excluded_df.columns else 0
    _cutoff_lbl = cutoff_date if cutoff_date else 'N/A (all dates included)'

    _combined = pd.concat([d for d in (filtered_df, stale_excluded_df) if d is not None and not d.empty], ignore_index=True)
    _t_rows = len(_combined)
    _t_devs = int(_combined['Name'].nunique())               if 'Name'               in _combined.columns else 0
    _t_cves = int(_combined['Vulnerability Name'].nunique()) if 'Vulnerability Name' in _combined.columns else 0

    row+=1
    ws.merge_range(row,0,row,3,f'  Data Filtering Reconciliation  (CVSS \u2265 {threshold})',sect_fmt); row+=1
    ws.write(row,0,'Filter Step',hdr_fmt); ws.write(row,1,'Unique Devices',hdr_fmt)
    ws.write(row,2,'Detection Rows',hdr_fmt); ws.write(row,3,'Unique CVE Types',hdr_fmt); row+=1
    ws.write(row,0,'[+]  Total raw detections (all devices, CVSS \u2265 threshold)',wf_plus)
    ws.write(row,1,_t_devs,val_fmt); ws.write(row,2,_t_rows,val_fmt); ws.write(row,3,_t_cves,val_fmt); row+=1
    if _stale_rows>0:
        ws.write(row,0,f'[-]  Excluded: stale devices  (Last Response before {_cutoff_lbl})',wf_minus)
        ws.write(row,1,_stale_devs,wf_mval); ws.write(row,2,_stale_rows,wf_mval); ws.write(row,3,_stale_cves,wf_mval); row+=1
    if not_in_rmm_count>0:
        ws.write(row,0,'[-]  Excluded: device not found in RMM',wf_minus)
        ws.write(row,1,not_in_rmm_count,wf_mval); ws.write(row,2,not_in_rmm_cve_count,wf_mval); ws.write(row,3,not_in_rmm_unique_cves,wf_mval); row+=1
    ws.write(row,0,'[=]  Active tracked scope  (Key Metrics above)',wf_eq_lbl)
    ws.write(row,1,unique_devices,wf_eq_val); ws.write(row,2,total_rows,wf_eq_val); ws.write(row,3,unique_cves,wf_eq_val); row+=1
    ws.merge_range(row,0,row,3,
                   '\u2139  Row counts subtract precisely. Unique Device and CVE Type counts may not subtract '
                   'perfectly \u2014 a CVE type on both an excluded and an active device is counted in both groups. '
                   'Stale devices are listed in the "Stale Excluded Devices" sheet.',note_fmt)
    ws.set_row(row,42); row+=1

    # ── Top At-Risk Devices ──────────────────────────────────────────────────────
    # Priority: 1) every server with ≥1 unresolved CVE  2) any device with a
    # known-exploit CVE  3) remainder by highest unresolved CVE count, up to 10.
    row += 1
    ws.merge_range(row, 0, row, 4,
                   '🚨  Top At-Risk Devices  (unresolved CVEs only)', sect_fmt)
    row += 1

    _has_uname   = 'Username'          in triage_df.columns
    _has_exploit = 'Has Known Exploit'  in triage_df.columns
    _has_dt      = 'Device Type'        in triage_df.columns
    _tar_sc      = ('Threat Status' if 'Threat Status' in triage_df.columns
                    else 'Status'   if 'Status'        in triage_df.columns else None)

    _th   = workbook.add_format({'bold': True, 'bg_color': '#2E75B6',
                                  'font_color': 'white', 'border': 1, 'align': 'center'})
    _td   = workbook.add_format({'border': 1})
    _td_r = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})
    _td_srv   = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC'})
    _td_srv_r = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC',
                                      'align': 'right', 'num_format': '#,##0'})
    _td_exp   = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6'})
    _td_exp_r = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6',
                                      'align': 'right', 'num_format': '#,##0'})

    ws.write(row, 0, '💻 Device Name', _th)
    ws.write(row, 1, '👤 Username',    _th)
    ws.write(row, 2, '⚠️ Unresolved CVEs', _th)
    ws.write(row, 3, '💣 Has Exploit', _th)
    ws.write(row, 4, '🖥️ Device Type', _th)
    row += 1

    if not triage_df.empty and 'Name' in triage_df.columns:
        _unr_df = (
            triage_df[triage_df[_tar_sc].astype(str).str.strip().str.upper() == 'UNRESOLVED'].copy()
            if _tar_sc else triage_df.copy()
        )
        if not _unr_df.empty:
            _agg = _unr_df.groupby('Name', as_index=False).agg(
                cve_count   =('Vulnerability Name', 'nunique'),
                username    =('Username', lambda s: next(
                    (v for v in s.astype(str) if v.strip() and v.lower() != 'nan'), ''))
                    if _has_uname else ('Name', lambda s: ''),
                has_exploit =('Has Known Exploit', lambda s:
                    'Yes' if s.astype(str).str.strip().str.lower()
                    .isin(['yes','true','1','y']).any() else 'No')
                    if _has_exploit else ('Name', lambda s: 'No'),
                device_type =('Device Type', 'first') if _has_dt else ('Name', lambda s: 'Unknown'),
            )
            _is_srv  = _agg['device_type'].astype(str).str.lower().str.contains('server', na=False)
            _is_exp  = _agg['has_exploit'].astype(str).str.strip().str.lower() == 'yes'
            _priority = set(_agg.loc[_is_srv | _is_exp, 'Name'].tolist())
            _sorted  = _agg.sort_values('cve_count', ascending=False)
            _ordered = (
                list(_sorted.loc[_sorted['Name'].isin(_priority)].itertuples(index=False))
                + list(_sorted.loc[~_sorted['Name'].isin(_priority)].itertuples(index=False))
            )
            _seen: set = set(); _top: list = []
            for _r in _ordered:
                if _r.Name not in _seen:
                    _seen.add(_r.Name); _top.append(_r)
                if len(_top) >= 10: break

            for _r in _top:
                _srv = 'server' in str(_r.device_type).lower()
                _exp = str(_r.has_exploit).strip().lower() == 'yes'
                _bf  = _td_exp   if _exp else (_td_srv   if _srv else _td)
                _nf  = _td_exp_r if _exp else (_td_srv_r if _srv else _td_r)
                ws.write(row, 0, str(_r.Name),        _bf)
                ws.write(row, 1, str(_r.username),    _bf)
                ws.write(row, 2, int(_r.cve_count),   _nf)
                ws.write(row, 3, str(_r.has_exploit), _bf)
                ws.write(row, 4, str(_r.device_type), _bf)
                row += 1

            ws.merge_range(row, 0, row, 4,
                'ℹ  🟡 Amber = Server (always listed if unresolved CVEs exist).  '
                '🟥 Red = device has CVEs with known exploit.  '
                'Up to 10 devices shown. Counts are unresolved detections only.',
                note_fmt)
            ws.set_row(row, 36); row += 1
        else:
            ws.merge_range(row, 0, row, 4, 'No unresolved CVE data.', note_fmt); row += 1
    else:
        ws.merge_range(row, 0, row, 4, 'No active device data.', note_fmt); row += 1

    ws.set_column('A:A', 44); ws.set_column('B:B', 24)
    ws.set_column('C:C', 20); ws.set_column('D:D', 18); ws.set_column('E:E', 16)

    # CVSS Score Split
    row+=1
    ws.merge_range(row,0,row,3,'  CVSS Score Split  (active detection rows)',sect_fmt); row+=1
    ws.write(row,0,'CVSS Score',hdr_fmt); ws.write(row,1,'Detection Rows',hdr_fmt)
    ws.write(row,2,'% of Total',hdr_fmt); ws.write(row,3,'Unique CVEs',hdr_fmt); row+=1
    score_split_start=row; score_split_data=[]
    if score_col:
        sg = triage_df.groupby(triage_df[score_col].round(1)).agg(rows=('Vulnerability Name','count'),cves=('Vulnerability Name','nunique')).sort_index(ascending=False)
        for sv,sr in sg.iterrows():
            pct=sr['rows']/total_rows if total_rows else 0
            ws.write(row,0,float(sv),lbl_fmt); ws.write(row,1,int(sr['rows']),val_fmt)
            ws.write(row,2,pct,val_pct); ws.write(row,3,int(sr['cves']),val_fmt)
            score_split_data.append((float(sv),int(sr['rows']),int(sr['cves']))); row+=1
    score_split_end=row-1

    # Month-over-Month
    mom_start_row=None; mom_data=[]
    if trend_data:
        m=trend_data['metrics']
        row+=1
        ws.merge_range(row,0,row,3,'  Month-over-Month Patching Progress',sect_fmt); row+=1
        mom_start_row=row
        ws.write(row,0,'Metric',hdr_fmt); ws.write(row,1,'Count',hdr_fmt)
        ws.write(row,2,'Direction',hdr_fmt); ws.write(row,3,'',hdr_fmt); row+=1
        for label,value,good in [
            ('CVE types resolved / patched',    m.get('resolved_cve_count',0),   True),
            ('CVE types newly introduced',       m.get('new_cve_count',0),        False),
            ('CVE types persisting (unpatched)', m.get('persisting_cve_count',0), False),
            ('Devices fully remediated',         m.get('remediated_devices',0),   True),
            ('New devices with CVEs',            m.get('new_devices',0),          False),
        ]:
            if good:
                vf=grn_fmt if value>0 else val_fmt; ds=f'\u25bc  {value:,}  (improvement)' if value>0 else '\u2014  no change'; df2=trend_up if value>0 else trend_eq
            else:
                vf=red_fmt if value>0 else val_fmt; ds=f'\u25b2  {value:,}  (increase)'    if value>0 else '\u2014  no change'; df2=trend_dn if value>0 else trend_eq
            ws.write(row,0,label,lbl_fmt); ws.write(row,1,value,vf); ws.merge_range(row,2,row,3,ds,df2)
            mom_data.append((label,value)); row+=1

    row+=1
    ws.write(row,0,'\u2139  All Key Metrics exclude stale devices and devices not found in RMM. '
                   'See the reconciliation table above for the full filtering breakdown.',note_fmt)

    # Charts
    if score_split_data and len(score_split_data)>=2:
        pie=workbook.add_chart({'type':'pie'})
        pie.add_series({'name':'Detection Rows',
                        'categories':['Client Summary',score_split_start,0,score_split_end,0],
                        'values':    ['Client Summary',score_split_start,1,score_split_end,1],
                        'data_labels':{'percentage':True,'category':True,'font':{'size':9}}})
        pie.set_title({'name':'Vulnerability Distribution by CVSS Score'}); pie.set_style(10)
        pie.set_size({'width':380,'height':260}); ws.insert_chart('F4',pie,{'x_offset':5,'y_offset':5})
        bar=workbook.add_chart({'type':'bar'})
        bar.add_series({'name':'Detection Rows',
                        'categories':['Client Summary',score_split_start,0,score_split_end,0],
                        'values':    ['Client Summary',score_split_start,1,score_split_end,1],
                        'fill':{'color':'#2E75B6'},'data_labels':{'value':True,'font':{'size':9}}})
        bar.add_series({'name':'Unique CVEs',
                        'categories':['Client Summary',score_split_start,0,score_split_end,0],
                        'values':    ['Client Summary',score_split_start,3,score_split_end,3],
                        'fill':{'color':'#ED7D31'},'data_labels':{'value':True,'font':{'size':9}}})
        bar.set_title({'name':'Patching Effort by CVSS Score'}); bar.set_x_axis({'name':'Count'})
        bar.set_y_axis({'name':'CVSS Score'}); bar.set_legend({'position':'bottom'}); bar.set_style(10)
        bar.set_size({'width':380,'height':260}); ws.insert_chart('F20',bar,{'x_offset':5,'y_offset':5})
    if trend_data and mom_data and mom_start_row is not None:
        mcs=mom_start_row+1; mce=mcs+len(mom_data)-1
        mb=workbook.add_chart({'type':'bar'})
        mb.add_series({'name':'Count',
                       'categories':['Client Summary',mcs,0,mce,0],
                       'values':    ['Client Summary',mcs,1,mce,1],
                       'fill':{'color':'#375623'},'data_labels':{'value':True,'font':{'size':9}}})
        mb.set_title({'name':'Month-over-Month Patching Progress'}); mb.set_x_axis({'name':'Count'})
        mb.set_legend({'none':True}); mb.set_style(10)
        mb.set_size({'width':380,'height':260}); ws.insert_chart('F36',mb,{'x_offset':5,'y_offset':5})

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