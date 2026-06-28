"""
excel_builder.py — xlsxwriter sheet-building functions.
No pandas data loading. No GUI. Receives DataFrames, writes sheets.

Author : Stu Villanti <s.villanti@kenstra.com>
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

# ── Patching Health Score ──────────────────────────────────────────────────────

def compute_patching_health_score(
    score_scope_dedup: 'pd.DataFrame',
    is_res: 'pd.Series',
    is_unr: 'pd.Series',
    trend_data: 'Optional[dict]' = None,
    has_patch_report: bool = False,
    score_scope_threshold: float = 7.0,
) -> dict:
    """
    Compute a 0–100 Patching Health Score.

    score_scope_dedup is the health-score scope (CVSS ≥ 7.0 by default), which is broader
    than the display threshold (usually CVSS ≥ 9.0).  This keeps Component 1 (resolution
    rate) and Component 2 (critical CVSS ≥ 9 coverage) genuinely distinct.

    Scoring model
    ───────────────────────────────────────────────────────────────────────────
    Component                            Weight   Basis
    ──────────────────────────────────────── ──────   ─────────────────────────────────────
    Resolution rate                        60 pts  % of CVSS ≥ 7.0 scope rows Resolved
    Critical (CVSS ≥ 9) coverage           20 pts  % of CVSS 9+ rows Resolved
    Known-exploit coverage                 20 pts  % of known-exploit rows Resolved
    ──────────────────────────────────────── ──────
    Subtotal (before penalties)           100 pts

    Penalties (deducted, floor = 0)
    ─────────────────────────────────────────────
    Persisting CVEs                     up to –5 pts  (–0.5 per persisting CVE type, cap –5)
    Unresolved CISA KEV CVEs            up to –5 pts  (–1 per unresolved KEV CVE type,  cap –5)

    Hard grade caps (applied to numeric score before grade assignment)
    ─────────────────────────────────────────────
    KEV ≥ 3, or KEV > 0 with no patch report  → max score 74 (C ceiling)
    KEV > 0                                    → max score 89 (B ceiling)

    Grade bands
    ───────────
    A  90–100  Excellent — nearly all CVEs remediated
    B  75–89   Good      — strong coverage, minor gaps
    C  60–74   Fair      — meaningful progress, notable gaps remain
    D  40–59   Poor      — significant unresolved exposure
    F   0–39   Critical  — majority of CVEs unpatched, immediate action needed
    """
    total_rows = len(score_scope_dedup)
    if total_rows == 0:
        return {
            'score': 0, 'grade': 'N/A', 'grade_colour': '#D9D9D9',
            'components': {}, 'penalties': {}, 'resolution_rate': 0.0,
            'confidence': {'has_patch_report': has_patch_report,
                           'score_scope_threshold': score_scope_threshold},
        }

    # Component 1: Resolution rate (60 pts) across the full health scope
    res_count = int(is_res.sum())
    res_rate  = res_count / total_rows
    pts_res   = round(res_rate * 60, 2)

    # Component 2: Critical CVE coverage — CVSS 9+ subset (20 pts)
    score_col = 'Vulnerability Score' if 'Vulnerability Score' in score_scope_dedup.columns else None
    if score_col:
        _sc_num    = pd.to_numeric(score_scope_dedup[score_col], errors='coerce')
        crit_mask  = _sc_num >= 9.0
        crit_total = int(crit_mask.sum())
        crit_res   = int((crit_mask & is_res).sum())
        crit_rate  = crit_res / crit_total if crit_total else 1.0
    else:
        crit_total = 0; crit_res = 0; crit_rate = 1.0
    pts_crit = round(crit_rate * 20, 2)

    # Component 3: Known-exploit coverage (20 pts)
    exp_col = 'Has Known Exploit' if 'Has Known Exploit' in score_scope_dedup.columns else None
    if exp_col:
        exp_mask   = score_scope_dedup[exp_col].astype(str).str.strip().str.lower().isin(
            ['yes', 'true', '1', 'y'])
        exp_total  = int(exp_mask.sum())
        exp_res    = int((exp_mask & is_res).sum())
        exp_rate   = exp_res / exp_total if exp_total else 1.0
    else:
        exp_total = 0; exp_res = 0; exp_rate = 1.0
    pts_exp = round(exp_rate * 20, 2)

    subtotal = pts_res + pts_crit + pts_exp

    # Penalty 1: Persisting CVE types from trend comparison
    persisting_count = 0
    if trend_data and isinstance(trend_data.get('metrics'), dict):
        persisting_count = int(trend_data['metrics'].get('persisting_cve_count', 0))
    pen_persisting = min(persisting_count * 0.5, 5.0)

    # Penalty 2: Unresolved CISA KEV CVEs — counted on the health-score scope (CVSS ≥ 7.0).
    # Using the broader scope is intentional: a KEV at CVSS 7.x is still active exploitation
    # risk and should cap the grade regardless of the display threshold.
    kev_unres_cves = 0
    kev_col = 'CISA KEV' if 'CISA KEV' in score_scope_dedup.columns else None
    if kev_col and 'Vulnerability Name' in score_scope_dedup.columns:
        kev_mask       = score_scope_dedup[kev_col].astype(str).str.strip().str.lower().isin(
            ['yes', 'true', '1', 'y'])
        kev_unres_df   = score_scope_dedup[kev_mask & is_unr]
        kev_unres_cves = int(kev_unres_df['Vulnerability Name'].nunique())
    pen_kev = min(kev_unres_cves * 1.0, 5.0)

    total_penalty = pen_persisting + pen_kev
    raw_score     = max(0.0, subtotal - total_penalty)
    score         = int(round(raw_score))

    # Hard grade caps — applied to the numeric score so displayed number and grade always agree.
    if kev_unres_cves >= 3 or (kev_unres_cves > 0 and not has_patch_report):
        score = min(score, 74)   # KEV exposure without patch evidence → C ceiling
    elif kev_unres_cves > 0:
        score = min(score, 89)   # any unresolved KEV → B ceiling

    if score >= 90:
        grade = 'A'; grade_colour = '#375623'
    elif score >= 75:
        grade = 'B'; grade_colour = '#70AD47'
    elif score >= 60:
        grade = 'C'; grade_colour = '#ED7D31'
    elif score >= 40:
        grade = 'D'; grade_colour = '#C00000'
    else:
        grade = 'F'; grade_colour = '#7B0000'

    return {
        'score':           score,
        'grade':           grade,
        'grade_colour':    grade_colour,
        'resolution_rate': res_rate,
        'components': {
            'resolution': {
                'rate': res_rate, 'pts': pts_res, 'weight': 60,
                'resolved': res_count, 'total': total_rows,
            },
            'critical_coverage': {
                'rate': crit_rate, 'pts': pts_crit, 'weight': 20,
                'resolved': crit_res, 'total': crit_total,
            },
            'exploit_coverage': {
                'rate': exp_rate, 'pts': pts_exp, 'weight': 20,
                'resolved': exp_res, 'total': exp_total,
            },
        },
        'penalties': {
            'persisting_cves': {'count': persisting_count, 'pts': pen_persisting},
            'kev_unresolved':  {'count': kev_unres_cves,   'pts': pen_kev},
        },
        'confidence': {
            'has_patch_report':      has_patch_report,
            'score_scope_threshold': score_scope_threshold,
        },
    }


def compute_score_lift(
    row: dict,
    total_rows: int,
    crit_total: int,
    exp_total: int,
    kev_unresolved_rows_by_cve: 'dict[str, int]',
    persisting_cves: 'set[str]',
) -> float:
    """
    Estimate health-score points recoverable by resolving this single row.

    Returns 0.0 for already-resolved rows.  All denominators come from the
    fleet-wide health scope (pre-computed in build_product_sheets) so values
    are comparable across all product sheets.

    Score lift components
    ─────────────────────────────────────────────
    Base resolution      60 / total_rows       every unresolved row
    Critical coverage   +20 / crit_total       only CVSS ≥ 9 rows
    Known exploit       +20 / exp_total        only rows with Has Known Exploit = Yes
    KEV penalty         +1.0                   only if this clears the LAST unresolved
                                                instance of this KEV CVE type
    Persisting penalty  +0.5                   only for CVE IDs in the persisting set
    """
    if str(row.get('Resolved', '')).strip() == '☑':
        return 0.0

    cve_id = extract_cve_id(str(row.get('Vulnerability Name', '')))
    lift   = 0.0

    if total_rows:
        lift += 60.0 / total_rows

    score = pd.to_numeric(row.get('Vulnerability Score', 0), errors='coerce')
    if crit_total and pd.notna(score) and score >= 9.0:
        lift += 20.0 / crit_total

    if exp_total and str(row.get('Has Known Exploit', '')).strip().lower() in ('yes', 'y', 'true', '1'):
        lift += 20.0 / exp_total

    if str(row.get('CISA KEV', '')).strip().lower() in ('yes', 'y', 'true', '1'):
        if kev_unresolved_rows_by_cve.get(cve_id, 0) == 1:
            lift += 1.0

    if cve_id in persisting_cves:
        lift += 0.5

    return round(lift, 2)


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
        'row_blue':     wb.add_format({'bg_color': '#BDD7EE', 'font_color': '#1F3864'}),  # deeper blue — resolved/confirmed
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
    # Plain text instead of hyperlinks: xlsxwriter has a hard limit of 65,530
    # URLs per worksheet; large sheets (Chrome/Edge 40k+ rows) exceed it.
    for row_i, val in enumerate(vuln_name_series, start=1):
        if CVE_PATTERN.search(str(val)):
            ws.write(row_i, col_idx, 'NVD ↗', link_fmt)

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
            # CVE text already written by to_excel(); set_column applies blue colour
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
                          report_month: str = '',
                          approaching_stale_names: Optional[Set[str]] = None,
                          stale_warning_days: int = 14):
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

    # ── N-day Exposure Age summary ────────────────────────────────────────────
    if 'N Days Exposed' in filtered_df.columns:
        _nde_num = pd.to_numeric(filtered_df['N Days Exposed'], errors='coerce')
        _unresolved_nde = _nde_num.dropna()  # '✓ Patched' and '—' become NaN, excluded

        # Deduplicate by CVE so each unique vulnerability is counted once
        _nde_per_cve = (
            filtered_df[['Vulnerability Name', 'N Days Exposed']]
            .drop_duplicates(subset=['Vulnerability Name'])
            .copy()
        )
        _nde_per_cve['_n'] = pd.to_numeric(_nde_per_cve['N Days Exposed'], errors='coerce')
        _nde_vals = _nde_per_cve['_n'].dropna()

        _band_180  = int((_nde_vals >= 180).sum())
        _band_91   = int(((_nde_vals >= 91)  & (_nde_vals < 180)).sum())
        _band_31   = int(((_nde_vals >= 31)  & (_nde_vals < 91)).sum())
        _band_0    = int((_nde_vals < 31).sum())
        _avg_age   = round(_nde_vals.mean(), 0) if not _nde_vals.empty else 0
        _max_age   = int(_nde_vals.max()) if not _nde_vals.empty else 0

        _nday_hdr  = workbook.add_format({
            'bold': True, 'bg_color': '#2E4057', 'font_color': 'white', 'border': 1,
        })
        _nday_crit = workbook.add_format({
            'bold': True, 'bg_color': '#C00000', 'font_color': 'white',
        })
        _nday_high = workbook.add_format({'bg_color': '#FCE4D6', 'bold': True})
        _nday_amb  = workbook.add_format({'bg_color': '#FFF2CC', 'bold': True})
        _nday_ok   = workbook.add_format({'bg_color': '#E2EFDA', 'bold': True})
        _nday_lbl  = workbook.add_format({'font_size': 9, 'italic': True, 'font_color': '#595959'})

        nr = r0 + 6
        ws.merge_range(nr, 0, nr, 5, 'N-Day Exposure Age  (unique CVEs, unpatched only)', _nday_hdr)
        ws.write(nr+1, 0, '≥ 180 days',        _nday_crit); ws.write(nr+1, 1, _band_180, _nday_crit)
        ws.write(nr+1, 2, 'Critical — far outside patch SLA', _nday_lbl)
        ws.write(nr+2, 0, '91 – 179 days',     _nday_high); ws.write(nr+2, 1, _band_91,  _nday_high)
        ws.write(nr+2, 2, 'High — breach 90-day remediation target', _nday_lbl)
        ws.write(nr+3, 0, '31 – 90 days',      _nday_amb);  ws.write(nr+3, 1, _band_31,  _nday_amb)
        ws.write(nr+3, 2, 'Amber — approaching or past 30-day target', _nday_lbl)
        ws.write(nr+4, 0, '0 – 30 days',       _nday_ok);   ws.write(nr+4, 1, _band_0,   _nday_ok)
        ws.write(nr+4, 2, 'Within acceptable window', _nday_lbl)
        ws.write(nr+5, 0, 'Avg exposure age');  ws.write(nr+5, 1, f'{int(_avg_age)} days')
        ws.write(nr+5, 3, 'Max exposure age');  ws.write(nr+5, 4, f'{_max_age} days')
        ws.write(nr+6, 0,
                 'ℹ  N = days since CVE.org/NVD Date Published (falls back to First Detected '
                 'if no publish date available). "✓ Patched" rows are excluded.',
                 workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 8}))
        ws.set_row(nr+6, 22)

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

def _build_patch_confirmed_sheet(writer, sheet_name: str, product: str,
                                  out_df: 'pd.DataFrame', col_names: list) -> None:
    """
    Write a lightweight "Patch Confirmed" sheet for a product where every
    detected CVE is already resolved (all rows are checkmark).

    Layout is intentionally minimal:
      - Same column order as a full triage sheet (col A = Resolved = checkmark)
        so Summary COUNTIF / COUNTIFS formulas count correctly.
      - Green header banner with a clear "all resolved" message.
      - Data rows written read-only (no dropdown validation — nothing to change).
      - No Score Lift column, no unresolved colouring, no legend.

    This keeps the workbook truthful and formula-safe without the bulk of a
    full triage sheet.
    """
    wb  = writer.book
    ws  = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # Banner
    banner_fmt = wb.add_format({
        'bold': True, 'font_size': 13,
        'bg_color': '#375623', 'font_color': 'white',
        'border': 1, 'align': 'left', 'valign': 'vcenter',
    })
    note_fmt = wb.add_format({
        'italic': True, 'font_size': 9,
        'font_color': '#375623', 'bg_color': '#E2EFDA', 'border': 1,
    })
    n_cols = min(len(col_names) - 1, 9)
    ws.merge_range(0, 0, 0, n_cols,
                   f'\u2705  {product}  \u2014  All CVEs Patch Confirmed', banner_fmt)
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, n_cols,
                   'Every detected CVE for this product has patch evidence or is marked RESOLVED '
                   'in N-able.  No action required.  Rows are read-only (\u2611 pre-filled).  '
                   'This sheet is included so Summary health-score formulas count these '
                   'resolutions correctly.',
                   note_fmt)
    ws.set_row(1, 28)

    # Column header row (row index 2)
    hdr_fmt = wb.add_format({
        'bold': True, 'bg_color': '#375623', 'font_color': 'white', 'border': 1,
    })
    for ci, col in enumerate(col_names):
        ws.write(2, ci, col, hdr_fmt)

    # Data rows (start at row index 3)
    grn_fmt = wb.add_format({'bg_color': '#E2EFDA', 'border': 1})
    grn_chk = wb.add_format({'bg_color': '#E2EFDA', 'border': 1,
                              'bold': True, 'align': 'center'})
    link_fmt_cell = wb.add_format({
        'font_color': '#0563C1', 'underline': True,
        'bg_color': '#E2EFDA', 'border': 1,
    })

    resolved_idx = col_names.index('Resolved') if 'Resolved' in col_names else None
    vuln_idx     = col_names.index('Vulnerability Name') if 'Vulnerability Name' in col_names else None

    for ri, row_tuple in enumerate(out_df.itertuples(index=False, name=None), start=3):
        for ci, val in enumerate(row_tuple):
            if ci == resolved_idx:
                ws.write(ri, ci, '\u2611', grn_chk)
            elif ci == vuln_idx:
                val_str = str(val) if val is not None else ''
                m = CVE_PATTERN.search(val_str)
                if m:
                    cve_id  = m.group(1).upper()
                    display = val_str[:255] if len(val_str) <= 255 else val_str[:252] + '...'
                    ws.write_url(ri, ci,
                                 f'https://www.cve.org/CVERecord?id={cve_id}',
                                 link_fmt_cell, string=display)
                else:
                    ws.write(ri, ci, val_str, grn_fmt)
            else:
                ws.write(ri, ci, val if val is not None else '', grn_fmt)

    n_data = len(out_df)

    # Column widths (mirrors full triage sheet)
    _widths = {
        'Resolved':                  10,
        'Score Lift':                10,
        'Vulnerability Name':        25,
        'Name':                      25,
        'Device Type':               15,
        'Vulnerability Severity':    20,
        'Vulnerability Score':        8,
        'Risk Severity Index':       18,
        'Has Known Exploit':         18,
        'CISA KEV':                  12,
        'Last Response':             22,
        'Days Since Last Response':  22,
        'Affected Products':         30,
        'Baseline Compliance':       22,
        'NVD':                       10,
    }
    for ci, col in enumerate(col_names):
        ws.set_column(ci, ci, _widths.get(col, 16))

    # Footer note
    foot_fmt = wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 8})
    ws.write(n_data + 4, 0,
             f'\u2139  {n_data} row(s) \u2014 all patch confirmed.  '
             'Sheet is read-only; use the full triage sheets for '
             'products with remaining unresolved CVEs.',
             foot_fmt)

    log.debug("Patch Confirmed sheet written for '%s': %d row(s)", product, n_data)


def build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                          patch_resolved_pairs=None,
                          patch_gap_pairs: Optional[Dict[Tuple[str, str], str]] = None,
                          approaching_stale_names: Optional[Set[str]] = None,
                          stale_warning_days: int = 14,
                          health_triage_df: 'Optional[pd.DataFrame]' = None,
                          trend_data: Optional[dict] = None):
    if patch_resolved_pairs is None:
        patch_resolved_pairs = set()
    if patch_gap_pairs is None:
        patch_gap_pairs = {}
    if approaching_stale_names is None:
        approaching_stale_names = set()

    # ── Fleet-level Score Lift context ────────────────────────────────────────────────────────
    # Pre-compute totals from the health scope once before the per-product loop
    # so every sheet divides by the same fleet-wide denominators.
    _sl_scope = health_triage_df if (health_triage_df is not None and not health_triage_df.empty) else triage_df
    _sl_dedup = _sl_scope.drop_duplicates(subset=['Name', 'Vulnerability Name'])
    _sl_sc_col    = 'Vulnerability Score' if 'Vulnerability Score' in _sl_dedup.columns else None
    _sl_total     = len(_sl_dedup)
    _sl_crit_total = int((pd.to_numeric(_sl_dedup[_sl_sc_col], errors='coerce') >= 9.0).sum()) if _sl_sc_col else 0
    _sl_exp_col   = 'Has Known Exploit' if 'Has Known Exploit' in _sl_dedup.columns else None
    _sl_exp_total = int(_sl_dedup[_sl_exp_col].astype(str).str.strip().str.lower().isin(['yes','y','true','1']).sum()) if _sl_exp_col else 0

    # Unresolved KEV row counts per CVE ID for penalty-recovery lift
    _sl_status_col = ('Threat Status' if 'Threat Status' in _sl_dedup.columns
                      else 'Status'   if 'Status'        in _sl_dedup.columns else None)
    _sl_is_unr = (~(_sl_dedup[_sl_status_col].astype(str).str.strip().str.upper() == 'RESOLVED')
                  if _sl_status_col
                  else pd.Series([True] * len(_sl_dedup), index=_sl_dedup.index))
    _kev_unres_by_cve: dict = {}
    if 'CISA KEV' in _sl_dedup.columns and 'Vulnerability Name' in _sl_dedup.columns:
        _kev_mask = _sl_dedup['CISA KEV'].astype(str).str.strip().str.lower().isin(['yes','y','true','1'])
        for _cve_raw in _sl_dedup.loc[_kev_mask & _sl_is_unr, 'Vulnerability Name']:
            _cid = extract_cve_id(str(_cve_raw))
            _kev_unres_by_cve[_cid] = _kev_unres_by_cve.get(_cid, 0) + 1

    _persisting_cves: set = set()
    if trend_data is not None:
        _persisting_cves = trend_data.get('persisting_cve_ids', set()) or set()

    # ── Pre-split patch_resolved_pairs by product key ───────────────────────────
    # When patch evidence is present, splitting once here means each product sheet
    # only checks its own small subset rather than the full global set.  This keeps
    # membership checks cheap even if patch_resolved_pairs grows large in future.
    _patch_2d:  set  = set()                  # 2-tuple pairs (device_key, cve_id)
    _patch_3d:  dict = {}                     # product_key → {(device_key, cve_id)}
    for _p in patch_resolved_pairs:
        if len(_p) == 3:
            _patch_3d.setdefault(_p[2], set()).add((_p[0], _p[1]))
        else:
            _patch_2d.add((_p[0], _p[1]))

    cols_order = ['Resolved', 'Score Lift', 'Vulnerability Name', 'Name', 'Device Type',
                  'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index',
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response', 'Affected Products',
                  'Baseline Compliance', 'NVD']

    def _chromium_sort_key(p: str) -> str:
        """Sort Chrome first, Edge immediately after (both are Chromium-based).
        All other products sort by their own name alphabetically."""
        pl = str(p).lower()
        if 'chrome' in pl and 'edge' not in pl:
            return 'google chrome\x00' + pl   # Chrome: sorts at its natural 'G' position
        if 'edge' in pl:
            return 'google chrome\x01' + pl   # Edge: immediately after Chrome
        return pl

    # Build groups once via groupby, then iterate in Chromium-aware order.
    # Fully-confirmed products are deferred so their sheets appear after all
    # active (partially-unresolved) product sheets and before the stale sheets.
    _product_groups = {p: g for p, g in triage_df.groupby('Base Product')}
    _deferred_confirmed: list = []   # (sheet_name, product, out_df, final_cols)

    for product in sorted(_product_groups.keys(), key=_chromium_sort_key):
        group = _product_groups[product]
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        # Primary sort applied after Score Lift is computed (below)

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

        # ── Performance: pre-compute normalised keys ONCE per group ──────────────
        # These are reused by both the Resolved column and the sparse set_row loop,
        # eliminating the duplicate regex work that apply + iterrows previously caused.
        _nk_list = [normalize_device_name(str(n)) for n in group['Name']]
        _ck_list = [extract_cve_id(str(v))        for v in group['Vulnerability Name']]

        # ── Resolved column — two-source priority ──────────────────────────────────────────
        # Source 1: patch evidence — uses pre-split sets so each group only checks its
        #           own product subset (_patch_3d[_sheet_pk]) or the shared 2-tuple set
        #           (_patch_2d).  Both are small regardless of overall patch_resolved_pairs size.
        # Source 2: Threat Status / Status == 'RESOLVED' — N-able export status.
        # Source 3: approaching-stale devices are always forced to ☐.
        _status_col = ('Threat Status' if 'Threat Status' in group.columns
                       else 'Status'   if 'Status'        in group.columns else None)

        _product_patch = _patch_3d.get(_sheet_pk, set())
        if _product_patch or _patch_2d:
            _res_bool = [
                (_product_patch and (nk, ck) in _product_patch) or (bool(_patch_2d) and (nk, ck) in _patch_2d)
                for nk, ck in zip(_nk_list, _ck_list)
            ]
        else:
            _res_bool = [False] * len(group)

        if _status_col:
            _status_resolved = (group[_status_col].astype(str)
                                .str.strip().str.upper().eq('RESOLVED').tolist())
            _res_bool = [_res_bool[i] or _status_resolved[i] for i in range(len(_res_bool))]

        if approaching_stale_names:
            _nm_list = group['Name'].tolist()
            _res_bool = [False if _nm_list[i] in approaching_stale_names else _res_bool[i]
                         for i in range(len(_res_bool))]

        _res_list = ['☑' if x else '☐' for x in _res_bool]

        # ── Fully-patched: defer to end of product sheets ──────────────────────
        # Accumulate confirmed products and write them after all active (partially-
        # unresolved) sheets so the tab order is: active products → confirmed
        # products → stale/NIRM sheets.  Sheet registration is deferred so
        # xlsxwriter writes tabs in the correct order.
        if all(v == '☑' for v in _res_list):
            group.insert(0, 'Resolved', _res_list)
            group['NVD'] = ''
            final_cols = [c for c in cols_order if c in group.columns]
            _out = group[final_cols]
            _deferred_confirmed.append((sheet_name, product, _out, final_cols))
            continue

        group.insert(0, 'Resolved', _res_list)

        # ── Score Lift ─────────────────────────────────────────────────────────────────
        _group_rows = group.to_dict('records')
        _sl_list = [
            compute_score_lift(r, _sl_total, _sl_crit_total, _sl_exp_total,
                               _kev_unres_by_cve, _persisting_cves)
            for r in _group_rows
        ]
        group.insert(1, 'Score Lift', _sl_list)

        group = group.sort_values(
            by=['Score Lift', 'Vulnerability Score', '_Sort_Time', 'Name'],
            ascending=[False, False, False, True],
        )
        group['NVD'] = ''

        final_cols = [c for c in cols_order if c in group.columns]
        _out = group[final_cols]

        # Direct write_row bypasses pandas to_excel overhead (~1.6× faster).
        # Register the sheet in writer.sheets so all subsequent set_column /
        # conditional_format / autofilter calls work exactly as before.
        wb_ = writer.book
        ws  = wb_.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws
        ws.write_row(0, 0, final_cols)
        for _ri, _row in enumerate(_out.itertuples(index=False, name=None), start=1):
            ws.write_row(_ri, 0, _row)

        ws.autofilter(0, 0, len(group), len(final_cols) - 1)

        styles_           = get_workbook_styles(wb_)
        patch_res_fmt     = styles_['row_blue']
        exploit_fmt       = wb_.add_format({'bg_color': '#FFE0CC'})
        coverage_fmt      = styles_['row_amber']
        unmanaged_fmt     = styles_['row_red']
        mismatch_fmt      = styles_['row_pink']
        installing_fmt    = styles_['row_teal']
        approaching_fmt   = wb_.add_format({'bg_color': '#FFF3E0', 'font_color': '#7B3F00'})  # orange-tinted — approaching stale

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

        _last = len(group)   # last data row (1-indexed)
        _TRUE_VALS = {'yes', 'true', '1', 'y'}

        unresolved_fmt = wb_.add_format({'bg_color': '#FFCCCC', 'font_color': '#8B0000'})  # coral red — unresolved, clearly distinct from blue

        # ── Bulk colouring via conditional_format (Excel XML — no Python loop needed) ──
        # Rules are evaluated in the order added — first added = highest priority in Excel.
        # IMPORTANT: range starts at Excel row 2 (xlsxwriter row index 1), so formula
        # must reference row 2 ($A2) not row 1 ($A1). Using $A1 causes an off-by-one:
        # each row gets coloured based on the PREVIOUS row's value, not its own.
        # Priority 1: Resolved (☑ in col A) → blue.
        ws.conditional_format(1, 0, _last, len(cl) - 1, {
            'type':     'formula',
            'criteria': '=$A2="☑"',
            'format':   patch_res_fmt,
        })
        # Priority 2: Unresolved (☐ in col A) → light red. Makes unresolved rows
        # immediately visible against the blue resolved rows.
        ws.conditional_format(1, 0, _last, len(cl) - 1, {
            'type':     'formula',
            'criteria': '=$A2="☐"',
            'format':   unresolved_fmt,
        })
        # Priority 3: Known exploit → darker orange (overrides unresolved red).
        _exp_col = 'Has Known Exploit'
        if _exp_col in cl:
            _ec = chr(ord('A') + cl.index(_exp_col))
            ws.conditional_format(1, 0, _last, len(cl) - 1, {
                'type':     'formula',
                'criteria': f'=OR(${_ec}2=TRUE,UPPER(TEXT(${_ec}2,"@"))="YES")',
                'format':   exploit_fmt,
            })

        # ── Sparse colouring via set_row for Python-computed states ──
        # Only approaching-stale and patch-gap rows need set_row; everything else is
        # handled by the conditional_format rules above.  Skip the loop entirely when
        # neither condition is active — that is the common case and saves ~80k iterations.
        _approaching = approaching_stale_names or set()
        if _approaching or patch_gap_pairs:
            # Pre-build column arrays so we avoid per-row dict look-ups from iterrows
            _name_arr    = group['Name'].tolist()    if 'Name'          in group.columns else [''] * len(group)
            _exp_arr     = (group[_exp_col].astype(str).str.strip().str.lower().tolist()
                            if _exp_col in group.columns else [''] * len(group))
            for _ri, (_nk, _ck, _nm, _rv, _ev) in enumerate(
                zip(_nk_list, _ck_list, _name_arr, _res_list, _exp_arr), start=1
            ):
                # Priority 1: approaching-stale overrides everything
                if _nm in _approaching:
                    _gap = patch_gap_pairs.get((_nk, _ck))
                    ws.set_row(_ri, None, _GAP_FMTS.get(_gap, approaching_fmt)
                               if _gap and _gap in _GAP_FMTS else approaching_fmt)
                    continue
                # Priority 2: resolved rows → handled by conditional_format
                if _rv == '☑':
                    continue
                # Priority 3: known exploit → handled by conditional_format
                if _ev in _TRUE_VALS:
                    continue
                # Priority 4: patch gap types
                _gap = patch_gap_pairs.get((_nk, _ck))
                if _gap and _gap in _GAP_FMTS:
                    ws.set_row(_ri, None, _GAP_FMTS[_gap])

        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25, link_fmt)
            # CVE text already written by to_excel(); set_column applies blue colour
        if 'NVD' in cl:
            nvd_idx = cl.index('NVD')
            ws.set_column(nvd_idx, nvd_idx, 10, link_fmt)
            _write_nvd_links(ws, group['Vulnerability Name'], nvd_idx, link_fmt)
        if 'Name'               in cl: ws.set_column(cl.index('Name'),               cl.index('Name'),               25)
        if 'Device Type'        in cl: ws.set_column(cl.index('Device Type'),        cl.index('Device Type'),        15)
        if 'Baseline Compliance' in cl: ws.set_column(cl.index('Baseline Compliance'), cl.index('Baseline Compliance'), 22)
        if 'Vulnerability Score' in cl:
            _vs_idx = cl.index('Vulnerability Score')
            _vs_col = get_col_letter(_vs_idx)
            ws.set_column(_vs_idx, _vs_idx, 8)
            # CVSS score colour coding — added AFTER row-level CFs so row colour
            # takes precedence on resolved/exploit rows; score colour shows on others.
            _crit_fmt = wb_.add_format({'bg_color': '#C00000', 'font_color': 'white',
                                        'bold': True,  'num_format': '0.0', 'align': 'center'})
            _high_fmt = wb_.add_format({'bg_color': '#ED7D31', 'font_color': 'white',
                                        'num_format': '0.0', 'align': 'center'})
            _med_fmt  = wb_.add_format({'bg_color': '#FFF2CC', 'font_color': '#7F6000',
                                        'num_format': '0.0', 'align': 'center'})
            _low_fmt  = wb_.add_format({'bg_color': '#E2EFDA',
                                        'num_format': '0.0', 'align': 'center'})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': '>=', 'value': 9.0, 'format': _crit_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 7.0, 'maximum': 8.9,
                'format': _high_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 4.0, 'maximum': 6.9,
                'format': _med_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 0.1, 'maximum': 3.9,
                'format': _low_fmt})

        legend_row = len(group) + 3
        l_title = wb_.add_format({'bold': True, 'font_size': 9, 'bg_color': '#F2F2F2', 'border': 1})
        l_cell  = wb_.add_format({'font_size': 9, 'border': 1})

        legend_entries = [
            ('#BDD7EE', 'blue row',   'Patch via RMM — install confirmed after CVE first detected'),
            ('#FFE0CC', 'orange row', 'Known active exploit — unresolved, prioritise immediately'),
            ('#FFF3E0', 'amber-orange row', f'Approaching stale — device offline \u2265 {stale_warning_days}d; patch confirmation unreliable (overrides blue)'),
            ('#FFF2CC', 'yellow row', 'Coverage gap — device not in patch report'),
            ('#FCE4D6', 'peach row',  'Unmanaged app — product not tracked in patch report'),
            ('#F2CEEF', 'pink row',   'Detection mismatch — CVE detected but no matching patch found'),
            ('#D9F0F4', 'teal row',   'Patch installing — patch is in progress, re-check after next RMM sync'),
            ('#FFCCCC', 'red row',    'Unresolved — patch not yet applied'),
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

    # ── Flush deferred confirmed sheets (after all active product sheets) ──────
    # These appear at the end of the product-sheet group, immediately before the
    # stale/NIRM sheets that the orchestrator writes next.
    for _sn, _prod, _out_df, _fcols in _deferred_confirmed:
        _build_patch_confirmed_sheet(writer, _sn, _prod, _out_df, _fcols)
    if _deferred_confirmed:
        log.debug(
            "Patch Confirmed sheets written: %d product(s) — %s",
            len(_deferred_confirmed),
            ', '.join(p for _, p, __, ___ in _deferred_confirmed),
        )

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
        'Installed but version unknown': '#BDD7EE',
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
        # Find the Lag column index for conditional_format
        _lag_cols = list(lag_df.columns)
        _lag_idx  = _lag_cols.index('Lag (days)') if 'Lag (days)' in _lag_cols else None
        _n = len(lag_df)
        if _lag_idx is not None:
            _lag_letter = chr(ord('A') + _lag_idx)
            # Green: 0–14 days — patched promptly
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_lag_letter}2>=0,${_lag_letter}2<=14)',
                'format': grn,
            })
            # Amber: 15–60 days — acceptable but slow
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_lag_letter}2>14,${_lag_letter}2<=60)',
                'format': amb,
            })
            # Red: >60 days or negative (patch pre-dates detection — data anomaly)
            ws.conditional_format(1, 0, _n, len(_lag_cols) - 1, {
                'type': 'formula', 'criteria': f'=OR(${_lag_letter}2>60,${_lag_letter}2<0)',
                'format': red,
            })
        ws.write(_n + 2, 0,
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
        # Colour by Distinct Versions count using conditional_format
        _dv_cols = list(drift_df.columns)
        _dv_idx  = _dv_cols.index('Distinct Versions') if 'Distinct Versions' in _dv_cols else None
        _dn = len(drift_df)
        if _dv_idx is not None:
            _dv_letter = chr(ord('A') + _dv_idx)
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=${_dv_letter}2=1',
                'format': grn,
            })
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=AND(${_dv_letter}2>=2,${_dv_letter}2<4)',
                'format': amb,
            })
            ws.conditional_format(1, 0, _dn, len(_dv_cols) - 1, {
                'type': 'formula', 'criteria': f'=${_dv_letter}2>=4',
                'format': red,
            })
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

    # Replace per-row set_row loop with a single conditional_format rule.
    # set_row() called N times = N xlsxwriter calls; one conditional_format
    # rule = one XML element regardless of row count — O(1) vs O(n).
    if len(out):
        ws.conditional_format(1, 0, len(out), len(out.columns) - 1, {
            'type':     'formula',
            'criteria': '=$A2<>""',
            'format':   grn,
        })

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


def build_client_summary_sheet(workbook, filtered_df, triage_df, threshold,
                               trend_data=None, customer_name='',
                               cutoff_date=None, stale_excluded_df=None,
                               not_in_rmm_count=0, not_in_rmm_cve_count=0,
                               not_in_rmm_unique_cves=0,
                               report_month='',
                               approaching_stale_names: Optional[Set[str]] = None,
                               stale_warning_days: int = 14,
                               product_to_sheet: Optional[dict] = None,
                               include_health_score: bool = False,
                               patch_resolved_pairs: Optional[set] = None,
                               health_triage_df: 'Optional[pd.DataFrame]' = None,
                               health_score_threshold: float = 7.0,
                               has_patch_report: bool = False):
    """
    Client Summary sheet.

    filtered_df            — score-filtered rows including not-in-RMM & stale (waterfall baseline).
    triage_df              — active scope only (stale + not-in-RMM removed). All Key Metrics use this.
    threshold              — CVSS floor shown in the waterfall header.
    patch_resolved_pairs   — same set passed to build_product_sheets so Summary and product tabs agree.
    health_triage_df       — broader scope (CVSS ≥ health_score_threshold) for health score only.
    health_score_threshold — actual CVSS floor used for health_triage_df (passed explicitly to keep
                             the footnote honest when the caller used a threshold below 7.0).
    has_patch_report       — tightens the KEV grade cap when no patch evidence is available.
    """
    ws = workbook.add_worksheet('Summary')
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

    # ── Key Metrics ─────────────────────────────────────────────────────────────
    # ALL counts from triage_df: active scope only (stale + not-in-RMM removed).
    # Deduplicate exactly as build_product_sheets does: for each Base Product sheet,
    # drop_duplicates(['Name','Vulnerability Name']), then concat.
    # This makes Key Metrics row counts match the live COUNTIF totals exactly.
    _p2s_keys = set((product_to_sheet or {}).keys())
    if 'Base Product' in triage_df.columns and _p2s_keys:
        _dedup_frames = [
            grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
            for bp, grp in triage_df.groupby('Base Product')
            if bp in _p2s_keys
        ]
        triage_dedup = pd.concat(_dedup_frames, ignore_index=True) if _dedup_frames else triage_df.copy()
    elif 'Base Product' in triage_df.columns:
        _dedup_frames = [
            grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
            for _, grp in triage_df.groupby('Base Product')
        ]
        triage_dedup = pd.concat(_dedup_frames, ignore_index=True) if _dedup_frames else triage_df.copy()
    else:
        triage_dedup = triage_df.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()

    _approaching_set = approaching_stale_names or set()

    # ── Compute resolved/unresolved by replaying the exact same ☑/☐ logic
    # that build_product_sheets writes into column A of each product sheet.
    # This guarantees the cached values supplied to write_formula() match what
    # the live COUNTIF formulas will compute.
    #
    # Resolution sources (mirror build_product_sheets priority):
    #   1. patch_resolved_pairs — explicit patch evidence overrides everything
    #   2. Threat Status / Status column == 'RESOLVED' — N-able export status
    #   3. Neither → ☐ (unresolved)
    # approaching-stale devices are always forced to ☐ regardless of source.
    from data_pipeline import _detect_product as _dp_detect_prod_sum, normalize_device_name as _ndv_sum, extract_cve_id as _eci_sum
    _p2s_sum   = product_to_sheet or {}
    _sc_col    = ('Threat Status' if 'Threat Status' in triage_dedup.columns
                  else 'Status'   if 'Status'        in triage_dedup.columns else None)
    _res_flags = []   # True = ☑ for each row in triage_dedup

    if 'Base Product' in triage_dedup.columns and _p2s_sum:
        for _bp, _grp in triage_dedup.groupby('Base Product', sort=False):
            if _bp not in _p2s_sum:
                _res_flags.extend([False] * len(_grp))
                continue
            # Determine the product key this sheet uses (mirrors build_product_sheets logic)
            _raw_pnames = _grp['Affected Products'].dropna().astype(str).unique().tolist() if 'Affected Products' in _grp.columns else []
            _sheet_pk = ''
            for _rpn in _raw_pnames:
                _pk = _dp_detect_prod_sum(_rpn)
                if _pk:
                    _sheet_pk = _pk
                    break
            if not _sheet_pk:
                _sheet_pk = _dp_detect_prod_sum(str(_bp))

            _nk = [_ndv_sum(str(n)) for n in _grp['Name']]
            _ck = [_eci_sum(str(v)) for v in _grp['Vulnerability Name']]
            _nm = _grp['Name'].tolist()

            # Source 1: patch_resolved_pairs (explicit patch evidence)
            if patch_resolved_pairs:
                _sample = next(iter(patch_resolved_pairs))
                if len(_sample) == 3:
                    _flags = [(nk, ck, _sheet_pk) in patch_resolved_pairs for nk, ck in zip(_nk, _ck)]
                else:
                    _flags = [(nk, ck) in patch_resolved_pairs for nk, ck in zip(_nk, _ck)]
            else:
                _flags = [False] * len(_grp)

            # Source 2: Threat Status / Status column — rows not already resolved by patch data
            if _sc_col and _sc_col in _grp.columns:
                _status_resolved = _grp[_sc_col].astype(str).str.strip().str.upper().eq('RESOLVED').tolist()
                _flags = [_flags[i] or _status_resolved[i] for i in range(len(_flags))]

            # Approaching-stale devices are always ☐
            if _approaching_set:
                _flags = [False if _nm[i] in _approaching_set else _flags[i] for i in range(len(_flags))]
            _res_flags.extend(_flags)
    else:
        # No product grouping — fall back to Threat Status column only
        _is_approaching_row = (triage_dedup['Name'].isin(_approaching_set)
                               if 'Name' in triage_dedup.columns
                               else pd.Series([False] * len(triage_dedup), index=triage_dedup.index))
        _raw_resolved = (triage_dedup[_sc_col].astype(str).str.strip().str.upper() == 'RESOLVED'
                         if _sc_col else pd.Series([False] * len(triage_dedup), index=triage_dedup.index))
        _res_flags = (_raw_resolved & ~_is_approaching_row).tolist()

    if len(_res_flags) == len(triage_dedup):
        _is_res = pd.Series(_res_flags, index=triage_dedup.index)
    else:
        # Length mismatch (groupby skipped some rows) — fall back to status column
        _is_approaching_row_fb = triage_dedup['Name'].isin(_approaching_set) if 'Name' in triage_dedup.columns else pd.Series([False]*len(triage_dedup), index=triage_dedup.index)
        _raw_resolved_fb = (triage_dedup[_sc_col].astype(str).str.strip().str.upper() == 'RESOLVED'
                            if _sc_col else pd.Series([False]*len(triage_dedup), index=triage_dedup.index))
        _is_res = _raw_resolved_fb & ~_is_approaching_row_fb
    _is_unr = ~_is_res

    total_rows     = len(triage_dedup)
    unique_cves    = int(triage_dedup['Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0
    unique_devices = int(triage_dedup['Name'].nunique())               if 'Name'               in triage_dedup.columns else 0
    score_col      = 'Vulnerability Score' if 'Vulnerability Score' in triage_dedup.columns else None
    crit_mask      = pd.to_numeric(triage_dedup[score_col], errors='coerce') >= 9.0 if score_col else pd.Series([True]*len(triage_dedup), index=triage_dedup.index)
    crit_rows      = int(crit_mask.sum())
    crit_cves      = int(triage_dedup.loc[crit_mask, 'Vulnerability Name'].nunique()) if score_col and 'Vulnerability Name' in triage_dedup.columns else unique_cves

    exploit_col     = 'Has Known Exploit' if 'Has Known Exploit' in triage_dedup.columns else None
    exploit_mask    = triage_dedup[exploit_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y']) if exploit_col else pd.Series([False]*len(triage_dedup), index=triage_dedup.index)
    exploit_count   = int(exploit_mask.sum())
    exploit_patched = int((exploit_mask & _is_res).sum())
    exploit_unpatch = int((exploit_mask & _is_unr).sum())

    # Server/workstation exploit breakdowns — detection row counts, active scope only
    if 'Device Type' in triage_dedup.columns:
        _srv_mask = triage_dedup['Device Type'].astype(str).str.lower().str.contains('server',      na=False)
        _wks_mask = triage_dedup['Device Type'].astype(str).str.lower().str.contains('workstation', na=False)
        srv_exp_total   = int((exploit_mask & _srv_mask).sum())
        srv_exp_patched = int((exploit_mask & _srv_mask & _is_res).sum())
        srv_exp_unpatch = int((exploit_mask & _srv_mask & _is_unr).sum())
        wks_exp_total   = int((exploit_mask & _wks_mask).sum())
        wks_exp_patched = int((exploit_mask & _wks_mask & _is_res).sum())
        wks_exp_unpatch = int((exploit_mask & _wks_mask & _is_unr).sum())
    else:
        srv_exp_total = srv_exp_patched = srv_exp_unpatch = 0
        wks_exp_total = wks_exp_patched = wks_exp_unpatch = 0

    # Device-type counts for Device Breakdown sub-table (unique devices)
    if 'Device Type' in triage_dedup.columns and 'Name' in triage_dedup.columns:
        srv_total   = int(triage_dedup.loc[_srv_mask,           'Name'].nunique())
        srv_unpatch = int(triage_dedup.loc[_srv_mask & _is_unr, 'Name'].nunique())
        wks_total   = int(triage_dedup.loc[_wks_mask,           'Name'].nunique())
        wks_unres   = int(triage_dedup.loc[_wks_mask & _is_unr, 'Name'].nunique())
    else:
        srv_total = srv_unpatch = wks_total = wks_unres = 0

    # Stale device-type counts (context — from stale_excluded_df)
    _stale_srv = _stale_wks = 0
    if stale_excluded_df is not None and not stale_excluded_df.empty and 'Device Type' in stale_excluded_df.columns and 'Name' in stale_excluded_df.columns:
        _stale_srv = int(stale_excluded_df[stale_excluded_df['Device Type'].astype(str).str.lower().str.contains('server',      na=False)]['Name'].nunique())
        _stale_wks = int(stale_excluded_df[stale_excluded_df['Device Type'].astype(str).str.lower().str.contains('workstation', na=False)]['Name'].nunique())

    # Stale + NIRM counts for Key Metrics math rows — computed once, reused in waterfall
    _stale_rows = int(len(stale_excluded_df)) if stale_excluded_df is not None and not stale_excluded_df.empty else 0
    _stale_devs = int(stale_excluded_df['Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Name' in stale_excluded_df.columns else 0
    _stale_crit = 0   # detection rows at CVSS 9+ on stale devices
    _stale_crit_cves = 0  # unique CVE types at CVSS 9+ on stale devices
    if stale_excluded_df is not None and not stale_excluded_df.empty and score_col and score_col in stale_excluded_df.columns:
        _stale_sc = pd.to_numeric(stale_excluded_df[score_col], errors='coerce')
        _stale_crit      = int((_stale_sc >= 9.0).sum())
        _stale_crit_cves = int(stale_excluded_df.loc[_stale_sc >= 9.0, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in stale_excluded_df.columns else 0
    # NIRM (not_in_rmm_cve_count already passed in as detection rows; compute CVSS 9+ subset from filtered_df)
    _nirm_devs  = not_in_rmm_count
    _nirm_crit  = 0
    _nirm_crit_cves = 0
    if 'Last Response' in filtered_df.columns and score_col and score_col in filtered_df.columns:
        _nirm_mask     = filtered_df['Last Response'] == 'Not Found in RMM'
        _nirm_sc       = pd.to_numeric(filtered_df.loc[_nirm_mask, score_col], errors='coerce')
        _nirm_crit     = int((_nirm_sc >= 9.0).sum())
        _nirm_crit_cves = int(filtered_df.loc[_nirm_mask & (pd.to_numeric(filtered_df[score_col], errors='coerce') >= 9.0), 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in filtered_df.columns else 0
    _excl_devs       = _stale_devs + _nirm_devs
    _excl_crit       = _stale_crit + _nirm_crit
    _excl_crit_cves  = _stale_crit_cves + _nirm_crit_cves   # note: may overlap; shown as informational

    # ── Key Metrics — 4-column grid: Metric | All | Active | Excluded ───────────
    # "All" = entire dataset (no stale/NIRM filter), "Active" = triage_df scope,
    # "Excluded" = stale + not-in-RMM. Active + Excluded = All (exact for rows).
    _zero_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center',
                                     'bg_color': '#E2EFDA', 'border': 1})
    _excl_hdr = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2',
                                     'font_color': '#595959', 'border': 1,
                                     'align': 'center'})
    _excl_val = workbook.add_format({'num_format': '#,##0', 'align': 'center',
                                     'bg_color': '#F2F2F2', 'font_color': '#595959',
                                     'border': 1})
    _all_hdr  = workbook.add_format({'bold': True, 'bg_color': '#D6E4F0',
                                     'font_color': '#1F3864', 'border': 1,
                                     'align': 'center'})
    _all_val  = workbook.add_format({'num_format': '#,##0', 'align': 'center',
                                     'bg_color': '#D6E4F0', 'font_color': '#1F3864',
                                     'border': 1})
    def _unr_fmt(n): return red_fmt if n > 0 else _zero_fmt

    # Compute "All" totals — deduped the same way as triage_dedup so All = Active + Excluded
    # Raw _all_df has cross-product duplicates (same CVE on same device under multiple
    # product versions) that inflate row counts vs what product sheets actually contain.
    # Apply the same per-Base-Product drop_duplicates to stale rows, then sum.
    if stale_excluded_df is not None and not stale_excluded_df.empty:
        if 'Base Product' in stale_excluded_df.columns and _p2s_keys:
            _stale_dedup_frames = [
                grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
                for bp, grp in stale_excluded_df.groupby('Base Product')
                if bp in _p2s_keys
            ]
            _stale_dedup = pd.concat(_stale_dedup_frames, ignore_index=True) if _stale_dedup_frames else stale_excluded_df.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        elif 'Base Product' in stale_excluded_df.columns:
            _stale_dedup_frames = [
                grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
                for _, grp in stale_excluded_df.groupby('Base Product')
            ]
            _stale_dedup = pd.concat(_stale_dedup_frames, ignore_index=True) if _stale_dedup_frames else stale_excluded_df.copy()
        else:
            _stale_dedup = stale_excluded_df.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
    else:
        _stale_dedup = pd.DataFrame()

    _all_df = pd.concat([triage_dedup, _stale_dedup], ignore_index=True) if not _stale_dedup.empty else triage_dedup.copy()

    _all_rows      = len(_all_df)          # = total_rows + stale_dedup_rows — math correct
    _all_cves      = int(_all_df['Vulnerability Name'].nunique()) if 'Vulnerability Name' in _all_df.columns else 0
    _all_devs      = int(_all_df['Name'].nunique())               if 'Name'               in _all_df.columns else 0
    _all_sc        = pd.to_numeric(_all_df.get(score_col, pd.Series(dtype=float)), errors='coerce') if score_col else pd.Series(dtype=float)

    # For CVSS 9+ counts use the FULL stale_excluded_df (deduplicated Name+CVE),
    # NOT _stale_dedup which filters to _p2s_keys (active products only).
    # A stale device running a product absent from active triage (e.g. curl only
    # on stale hosts) would otherwise be silently dropped, showing All=0.
    if stale_excluded_df is not None and not stale_excluded_df.empty:
        _stale_full_dedup = stale_excluded_df.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
    else:
        _stale_full_dedup = pd.DataFrame()
    _stale_full_sc = (
        pd.to_numeric(_stale_full_dedup[score_col], errors='coerce')
        if score_col and not _stale_full_dedup.empty and score_col in _stale_full_dedup.columns
        else pd.Series(dtype=float)
    )
    _stale_full_crit      = int((_stale_full_sc >= 9.0).sum())
    _stale_full_crit_cves = (
        int(_stale_full_dedup.loc[_stale_full_sc >= 9.0, 'Vulnerability Name'].nunique())
        if 'Vulnerability Name' in _stale_full_dedup.columns else 0
    )
    _all_crit      = crit_rows + _stale_full_crit
    _all_crit_cves = crit_cves + _stale_full_crit_cves

    # stale crit rows from _stale_dedup directly (consistent with _all_rows)
    if not _stale_dedup.empty and score_col and score_col in _stale_dedup.columns:
        _stale_crit_dedup = int((pd.to_numeric(_stale_dedup[score_col], errors='coerce') >= 9.0).sum())
    else:
        _stale_crit_dedup = _stale_crit

    _excl_rows          = len(_stale_dedup) + not_in_rmm_cve_count
    _excl_devs_tot      = _stale_devs + _nirm_devs
    _excl_crit_tot      = _stale_full_crit + _nirm_crit        # full stale, not _p2s_keys-filtered
    _excl_crit_cves_tot = _stale_full_crit_cves + _nirm_crit_cves

    # ── Patching Health Score (beta) ──────────────────────────────────────────
    # Only rendered when include_health_score=True (opt-in checkbox in the GUI).
    if include_health_score:
        # Build a separate dedup from health_triage_df (CVSS ≥ 7.0 by default) so the
        # critical-coverage component (CVSS ≥ 9) is genuinely distinct from resolution rate.
        # Falls back to triage_dedup when no broader scope was supplied.
        _health_raw = health_triage_df if (health_triage_df is not None and not health_triage_df.empty) else None
        if _health_raw is not None:
            _p2s_hs = set((product_to_sheet or {}).keys())
            if 'Base Product' in _health_raw.columns and _p2s_hs:
                _hs_frames = [grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
                              for bp, grp in _health_raw.groupby('Base Product') if bp in _p2s_hs]
            else:
                _hs_frames = [grp.drop_duplicates(subset=['Name', 'Vulnerability Name'])
                              for _, grp in (_health_raw.groupby('Base Product')
                                             if 'Base Product' in _health_raw.columns
                                             else [(None, _health_raw)])]
            _score_scope = (pd.concat(_hs_frames, ignore_index=True) if _hs_frames
                            else _health_raw.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy())

            # Resolved flags: patch evidence (2-tuple) | status column | suppress approaching-stale
            from data_pipeline import normalize_device_name as _ndv_hs, extract_cve_id as _eci_hs
            _hs_sc_col = ('Threat Status' if 'Threat Status' in _score_scope.columns
                          else 'Status'   if 'Status'        in _score_scope.columns else None)
            _hs_status_res = (_score_scope[_hs_sc_col].astype(str).str.strip().str.upper() == 'RESOLVED'
                              if _hs_sc_col
                              else pd.Series([False] * len(_score_scope), index=_score_scope.index))
            if patch_resolved_pairs:
                _pair_2d = {(p[0], p[1]) for p in patch_resolved_pairs}
                _hs_nk = _score_scope['Name'].astype(str).apply(_ndv_hs)
                _hs_ck = _score_scope['Vulnerability Name'].astype(str).apply(_eci_hs)
                _hs_patch = pd.Series([(nk, ck) in _pair_2d for nk, ck in zip(_hs_nk, _hs_ck)],
                                      index=_score_scope.index)
                _hs_combined = _hs_patch | _hs_status_res
            else:
                _hs_combined = _hs_status_res
            if _approaching_set and 'Name' in _score_scope.columns:
                _hs_combined = _hs_combined & ~_score_scope['Name'].isin(_approaching_set)
            _hs_is_res = _hs_combined
            _hs_is_unr = ~_hs_is_res
            _score_scope_threshold = health_score_threshold
        else:
            _score_scope = triage_dedup
            _hs_is_res   = _is_res
            _hs_is_unr   = _is_unr
            _score_scope_threshold = threshold

        _phs = compute_patching_health_score(
            _score_scope, _hs_is_res, _hs_is_unr,
            trend_data=trend_data,
            has_patch_report=has_patch_report,
            score_scope_threshold=_score_scope_threshold,
        )
        _phs_score  = _phs['score']
        _phs_grade  = _phs['grade']
        _phs_colour = _phs['grade_colour']
        _phs_comps  = _phs['components']
        _phs_pens   = _phs['penalties']
        _phs_conf   = _phs.get('confidence', {})

        # ── Live-formula helpers ───────────────────────────────────────────────
        # Product sheets have a fixed column layout (set in cols_order):
        #   A(0)=Resolved  B(1)=Score Lift  C(2)=Vulnerability Name  D(3)=Name
        #   E(4)=Device Type  F(5)=Vulnerability Severity  G(6)=Vulnerability Score
        #   H(7)=Risk Severity Index  I(8)=Has Known Exploit  J(9)=CISA KEV
        #
        # We build cross-sheet COUNTIFS formulas so the three score components
        # update automatically when ☐/☑ are toggled in any product sheet.
        # Penalties (persisting CVEs, unresolved KEVs) depend on external data
        # (trend comparison, KEV database) and cannot be recalculated in-sheet;
        # they remain static and are labelled "(fixed at generation)".
        _p2s_hs_vals = list((product_to_sheet or {}).values())   # list of sheet name strings
        _hs_live = bool(_p2s_hs_vals)

        if _hs_live:
            # ── Component 1: Resolution rate across ALL product sheets ──────────
            # Totals = ☑ + ☐ (avoids counting header/legend rows that COUNTA would hit)
            _f_hs_res   = ' + '.join([f"COUNTIF('{s}'!A:A,\"☑\")" for s in _p2s_hs_vals])
            _f_hs_unres = ' + '.join([f"COUNTIF('{s}'!A:A,\"☐\")" for s in _p2s_hs_vals])
            _f_hs_total = f'({_f_hs_res}) + ({_f_hs_unres})'

            # ── Component 2: Critical CVE coverage (CVSS ≥ 9) ─────────────────
            # Count rows where Resolved=☑ AND Vulnerability Score ≥ 9  (col G)
            # Count rows where (☑ OR ☐)   AND Vulnerability Score ≥ 9  → total critical rows
            _f_crit_res = ' + '.join(
                [f"COUNTIFS('{s}'!A:A,\"☑\",'{s}'!G:G,\">=\"&9)" for s in _p2s_hs_vals]
            )
            _f_crit_total = ' + '.join(
                [f"COUNTIFS('{s}'!G:G,\">=\"&9,'{s}'!A:A,\"<>\")" for s in _p2s_hs_vals]
            )
            # Subtract 1 per sheet for the header row that COUNTIFS would include via "<>"
            # Header cell in col G is text ("Vulnerability Score") so ">="&9 already
            # excludes it — no correction needed for crit_total.
            # For the "<>" total we do need to subtract the header rows:
            _f_crit_total = (
                ' + '.join([f"COUNTIFS('{s}'!G:G,\">=\"&9,'{s}'!A:A,\"☑\")"
                            for s in _p2s_hs_vals])
                + ' + '
                + ' + '.join([f"COUNTIFS('{s}'!G:G,\">=\"&9,'{s}'!A:A,\"☐\")"
                              for s in _p2s_hs_vals])
            )

            # ── Component 3: Known-exploit coverage (col I = Has Known Exploit) ─
            _f_exp_res = ' + '.join(
                [f"COUNTIFS('{s}'!A:A,\"☑\",'{s}'!I:I,\"Yes\")" for s in _p2s_hs_vals]
            )
            _f_exp_total = (
                ' + '.join([f"COUNTIFS('{s}'!I:I,\"Yes\",'{s}'!A:A,\"☑\")"
                            for s in _p2s_hs_vals])
                + ' + '
                + ' + '.join([f"COUNTIFS('{s}'!I:I,\"Yes\",'{s}'!A:A,\"☐\")"
                              for s in _p2s_hs_vals])
            )

            # Static penalty values (cannot be recalculated from the sheet columns)
            _pen_persist_pts = _phs_pens.get('persisting_cves', {}).get('pts', 0.0)
            _pen_kev_pts     = _phs_pens.get('kev_unresolved',  {}).get('pts', 0.0)
            _total_pen       = _pen_persist_pts + _pen_kev_pts

            # -- Live score formula -------------------------------------------------
            # Mirrors compute_patching_health_score:
            #   pts_res  = IF(total>0, res/total, 0) * 60
            #   pts_crit = IF(crit_total>0, crit_res/crit_total, 1) * 20
            #   pts_exp  = IF(exp_total>0,  exp_res/exp_total,   1) * 20
            #   score    = MAX(0, INT(ROUND(pts_res + pts_crit + pts_exp - penalties, 0)))
            #
            # IMPORTANT: embedding _f_score into _f_grade would repeat the full
            # cross-sheet COUNTIF expression 4+ times, easily exceeding Excel's
            # 8,192-char formula limit and corrupting the XLSX XML.
            # Fix: write the score into a hidden helper cell (col E, row 4 --
            # outside the visible 4-col A-D layout) and reference $E$4 everywhere.
            _f_pts_res  = f'IF(({_f_hs_total})>0,({_f_hs_res})/({_f_hs_total}),0)*60'
            _f_pts_crit = f'IF(({_f_crit_total})>0,({_f_crit_res})/({_f_crit_total}),1)*20'
            _f_pts_exp  = f'IF(({_f_exp_total})>0,({_f_exp_res})/({_f_exp_total}),1)*20'
            _f_raw      = f'({_f_pts_res})+({_f_pts_crit})+({_f_pts_exp})-{_total_pen}'
            _f_score    = f'MAX(0,INT(ROUND({_f_raw},0)))'

            # Helper cell E4 (row=3, col=4): holds the live numeric score.
            # Grade IFS and final score row both reference $E$4 -- short formulas.
            _helper_row = 3
            _helper_col = 4
            _helper_ref = '$E$4'
            _f_grade = (
                f'IFS({_helper_ref}>=90,"A",{_helper_ref}>=75,"B",'
                f'{_helper_ref}>=60,"C",{_helper_ref}>=40,"D",TRUE,"F")'
            )
            _f_score_ref = f'={_helper_ref}'

            # Excel's stored formula limit is 8,192 characters. With many product
            # sheets the live cross-sheet COUNTIFS for the health score can exceed
            # that limit, which causes Excel to show "We found a problem with some
            # content" and repair/remove the formulas when opening the workbook.
            # When any health-score formula is too long, fall back to the static
            # generation-time values rather than writing an invalid workbook.
            _candidate_live_formulas = [
                _f_score,
                _f_grade,
                _f_score_ref,
                f'IF(({_f_hs_total})>0,({_f_hs_res})/({_f_hs_total}),1)',
                _f_pts_res,
                f'IF(({_f_crit_total})>0,({_f_crit_res})/({_f_crit_total}),1)',
                _f_pts_crit,
                f'IF(({_f_exp_total})>0,({_f_exp_res})/({_f_exp_total}),1)',
                _f_pts_exp,
            ]
            _max_formula_len = max(len(str(f or '')) + 1 for f in _candidate_live_formulas)
            if _max_formula_len > 8192:
                log.warning(
                    "Patching Health Score live formulas disabled: longest formula is %d "
                    "characters, above Excel's 8,192 character limit. Static score values "
                    "will be written instead.",
                    _max_formula_len,
                )
                _hs_live = False

        # ── Format objects ──────────────────────────────────────────────────────
        # Score box and grade box: colour is driven by static _phs_colour at
        # generation time (correct for the initial state).  After toggling ☑/☐
        # the numeric score and grade letter update live; the background colour
        # is updated via conditional formatting rules added below.
        _live_score_bg   = '#1F4E79'   # neutral dark blue for the live score container
        _score_box_fmt = workbook.add_format({
            'bold': True, 'font_size': 36, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': _live_score_bg, 'border': 2,
        })
        _grade_box_fmt = workbook.add_format({
            'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': _live_score_bg, 'border': 2,
        })
        _score_lbl_fmt = workbook.add_format({
            'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
            'font_color': '#595959', 'bg_color': '#F9F9F9', 'border': 1,
        })
        _comp_hdr_fmt = workbook.add_format({
            'bold': True, 'font_size': 9, 'bg_color': '#D6E4F0', 'border': 1, 'align': 'center',
        })
        _comp_lbl_fmt = workbook.add_format({
            'font_size': 9, 'bg_color': '#F2F2F2', 'border': 1,
        })
        _comp_lbl_italic_fmt = workbook.add_format({
            'font_size': 9, 'bg_color': '#F2F2F2', 'border': 1, 'italic': True,
            'font_color': '#595959',
        })
        _comp_pct_fmt = workbook.add_format({
            'font_size': 9, 'num_format': '0%', 'align': 'right', 'border': 1,
            'bg_color': '#EBF3FB', 'font_color': '#1F3864',  # blue tint = live cell
        })
        _comp_pts_fmt = workbook.add_format({
            'font_size': 9, 'num_format': '0.0', 'align': 'right', 'border': 1,
            'bg_color': '#EBF3FB', 'font_color': '#1F3864',  # blue tint = live cell
        })
        _comp_pts_static_fmt = workbook.add_format({
            'font_size': 9, 'num_format': '0.0', 'align': 'right', 'border': 1,
            'font_color': '#1F3864',
        })
        _pen_neg_fmt = workbook.add_format({
            'font_size': 9, 'num_format': '0.0', 'align': 'right', 'border': 1, 'font_color': '#C00000',
        })
        _pen_zero_fmt = workbook.add_format({
            'font_size': 9, 'align': 'right', 'border': 1, 'font_color': '#595959',
        })
        _phs_final_lbl_fmt = workbook.add_format({
            'bold': True, 'font_size': 9, 'bg_color': _phs_colour,
            'font_color': 'white', 'border': 1,
        })
        _phs_final_val_fmt = workbook.add_format({
            'bold': True, 'font_size': 9, 'num_format': '0', 'align': 'right',
            'bg_color': '#EBF3FB', 'font_color': '#1F3864', 'border': 2,
        })
        _score_note_fmt = workbook.add_format({
            'italic': True, 'font_size': 8, 'font_color': '#595959', 'text_wrap': True,
        })
        _beta_fmt = workbook.add_format({
            'italic': True, 'font_size': 8, 'font_color': '#7F6000',
            'bg_color': '#FFF2CC', 'border': 1,
        })

        row = 3
        ws.set_row(row,     44)
        ws.set_row(row + 1, 14)

        # Score box (cols A-B) -- live formula or static fallback.
        # xlsxwriter supports formulas in merge_range by passing the formula string
        # as the data argument directly. Calling write_formula on the same cell
        # afterwards corrupts the XLSX XML, so we use a single merge_range call.
        _score_box_row = row   # remember for conditional formatting below
        if _hs_live:
            # Write the full score formula into hidden helper cell E4 first.
            # All visible cells reference $E$4 to stay under the 8192-char limit.
            ws.write_formula(_helper_row, _helper_col, f'={_f_score}',
                             workbook.add_format({'num_format': '0', 'font_color': 'white',
                                                  'bg_color': 'white'}), _phs_score)
            ws.merge_range(row, 0, row + 1, 1, _f_score_ref, _score_box_fmt)
        else:
            ws.merge_range(row, 0, row + 1, 1, _phs_score, _score_box_fmt)

        # Grade box (col C) -- same pattern
        if _hs_live:
            ws.merge_range(row, 2, row + 1, 2, f'={_f_grade}', _grade_box_fmt)
        else:
            ws.merge_range(row, 2, row + 1, 2, _phs_grade, _grade_box_fmt)

        ws.merge_range(row, 3, row + 1, 3, 'Patching Health Score  (0\u2013100)', _score_lbl_fmt)
        row += 2

        # ── Conditional formatting: colour score/grade boxes by live score value ─
        # Applied to the top-left cell of each merged region (xlsxwriter targets the
        # top-left cell; Excel applies the format to the entire merged range).
        if _hs_live:
            _cf_base = {'type': 'cell', 'multi_range': f'A{_score_box_row+1} C{_score_box_row+1}'}
            for _cf_min, _cf_max, _cf_bg in [
                (90, 100, '#375623'),   # A — dark green
                (75,  89, '#70AD47'),   # B — green
                (60,  74, '#ED7D31'),   # C — orange
                (40,  59, '#C00000'),   # D — red
                ( 0,  39, '#7B0000'),   # F — dark red
            ]:
                _cf_fmt = workbook.add_format({
                    'bold': True, 'font_color': 'white', 'bg_color': _cf_bg, 'border': 2,
                })
                ws.conditional_format(_score_box_row, 0, _score_box_row + 1, 2, {
                    'type': 'cell', 'criteria': 'between',
                    'minimum': _cf_min, 'maximum': _cf_max,
                    'format': _cf_fmt,
                })

        ws.write(row, 0, 'Score Component',  _comp_hdr_fmt)
        ws.write(row, 1, 'Coverage Rate',    _comp_hdr_fmt)
        ws.write(row, 2, 'Points Earned',    _comp_hdr_fmt)
        ws.write(row, 3, '/ Max',            _comp_hdr_fmt)
        row += 1

        # ── Component rows ────────────────────────────────────────────────────
        # (rate, pts) are live formulas when _hs_live, static values otherwise.
        _comp_rows = [
            ('Resolution rate',                     'resolution',       60,
             _f_hs_res   if _hs_live else None,
             _f_hs_total if _hs_live else None),
            ('Critical CVE coverage (CVSS \u22659)', 'critical_coverage', 20,
             _f_crit_res   if _hs_live else None,
             _f_crit_total if _hs_live else None),
            ('Known-exploit coverage',               'exploit_coverage', 20,
             _f_exp_res   if _hs_live else None,
             _f_exp_total if _hs_live else None),
        ]
        _comp_pts_fmls = [
            _f_pts_res  if _hs_live else None,
            _f_pts_crit if _hs_live else None,
            _f_pts_exp  if _hs_live else None,
        ]

        for (_lbl, _key, _max, _f_res_n, _f_tot_n), _f_pts_n in zip(_comp_rows, _comp_pts_fmls):
            _c = _phs_comps.get(_key, {})
            _static_rate = _c.get('rate', 1.0)
            _static_pts  = _c.get('pts',  float(_max))
            ws.write(row, 0, _lbl, _comp_lbl_fmt)
            if _hs_live and _f_res_n and _f_tot_n:
                ws.write_formula(row, 1,
                    f'=IF(({_f_tot_n})>0,({_f_res_n})/({_f_tot_n}),1)',
                    _comp_pct_fmt, _static_rate)
                ws.write_formula(row, 2, f'={_f_pts_n}', _comp_pts_fmt, _static_pts)
            else:
                ws.write(row, 1, _static_rate, _comp_pct_fmt)
                ws.write(row, 2, _static_pts,  _comp_pts_static_fmt)
            ws.write(row, 3, f'/ {_max}', _comp_lbl_fmt)
            row += 1

        # ── Penalty rows (always static — depend on external data) ─────────────
        for _plbl, _pkey, _punit in [
            ('Persisting CVE types (\u22120.5 each, max \u22125)  \u2013 fixed at generation',
             'persisting_cves', 'CVE types'),
            ('Unresolved CISA KEV CVEs (\u22121 each, max \u22125)  \u2013 fixed at generation',
             'kev_unresolved',  'KEV CVEs'),
        ]:
            _pen  = _phs_pens.get(_pkey, {})
            _ppts = _pen.get('pts', 0)
            _pcnt = _pen.get('count', 0)
            ws.write(row, 0, _plbl,               _comp_lbl_italic_fmt)
            ws.write(row, 1, f'{_pcnt} {_punit}', _comp_lbl_italic_fmt)
            ws.write(row, 2, -_ppts if _ppts > 0 else None,
                     _pen_neg_fmt if _ppts > 0 else _pen_zero_fmt)
            ws.write(row, 3, 'penalty',            _comp_lbl_italic_fmt)
            row += 1

        # ── Final score row ────────────────────────────────────────────────────
        ws.merge_range(row, 0, row, 1,
                       f'Patching Health Score  (Grade {_phs_grade})',
                       _phs_final_lbl_fmt)
        if _hs_live:
            ws.write_formula(row, 2, _f_score_ref, _phs_final_val_fmt, _phs_score)
        else:
            ws.write(row, 2, _phs_score, _phs_final_val_fmt)
        ws.write(row, 3, '/ 100', _phs_final_lbl_fmt)
        row += 1

        _scope_note = (
            f'Score calculated from CVSS ≥ {_score_scope_threshold:.1f} active scope'
            if _score_scope_threshold < threshold
            else 'Score calculated from active scope'
        )
        _kev_cap_note = ''
        _kev_cnt = _phs_pens.get('kev_unresolved', {}).get('count', 0)
        if not _phs_conf.get('has_patch_report') and _kev_cnt > 0:
            _kev_cap_note = '  ·  No patch report: KEV grade cap active.'
        elif _kev_cnt > 0:
            _kev_cap_note = '  ·  Unresolved KEV CVEs: grade cap active.'
        _live_note_hs = (
            '\u26a1 Blue cells update automatically when \u2610/\u2611 are toggled in product sheets.  '
            'Penalties are fixed at report generation (depend on trend/KEV data).  '
        ) if _hs_live else ''
        ws.merge_range(row, 0, row, 3,
            f'{_live_note_hs}'
            'Grade bands:  A \u2265 90  |  B \u2265 75  |  C \u2265 60  |  D \u2265 40  |  F < 40  ·  '
            'Score = Resolution rate (60 pts) + Critical CVE coverage (20 pts) + '
            'Known-exploit coverage (20 pts) \u2212 penalties.  '
            f'{_scope_note} (stale / not-in-RMM excluded).{_kev_cap_note}',
            _score_note_fmt)
        ws.set_row(row, 28)
        row += 1

        ws.merge_range(row, 0, row, 3,
            '\u26a0  Beta feature \u2014 scoring methodology may change. '
            'Do not use for formal reporting without validation.',
            _beta_fmt)
        row += 2

    else:
        # Health score disabled — Key Metrics starts at row 3
        row = 3

    ws.merge_range(row, 0, row, 3, '  Key Metrics', sect_fmt)
    row += 1

    # Header row
    ws.write(row, 0, 'Metric',       hdr_fmt)
    ws.write(row, 1, 'All',          _all_hdr)
    ws.write(row, 2, 'Active only',  hdr_fmt)
    ws.write(row, 3, 'Excl. (stale / not in RMM)', _excl_hdr)
    row += 1

    for metric, all_val, active_val, excl_val, active_fmt in [
        ('Total detection rows',        _all_rows,          total_rows,     _excl_rows,             val_fmt),
        ('Unique CVE types',            _all_cves,          unique_cves,    0,                      val_fmt),   # CVEs overlap — excl shown as n/a
        ('Unique devices',              _all_devs,          unique_devices, _excl_devs_tot,         val_fmt),
        ('Detections at CVSS 9.0+',    _all_crit,          crit_rows,      _excl_crit_tot,         red_fmt if crit_rows else val_fmt),
        ('Unique CVEs at CVSS 9.0+',   _all_crit_cves,     crit_cves,      0,                      red_fmt if crit_cves else val_fmt),   # overlap
    ]:
        ws.write(row, 0, metric,      lbl_fmt)
        ws.write(row, 1, all_val,     _all_val)
        ws.write(row, 2, active_val,  active_fmt)
        ws.write(row, 3, excl_val if excl_val else '—',  _excl_val)
        row += 1

    ws.merge_range(row, 0, row, 3,
                   '\u2139  All = full dataset at CVSS \u2265 threshold.  '
                   'Active = excludes stale (\u226530 days without response) and devices not found in RMM.  '
                   'Active + Excluded = All for row counts and device counts.  '
                   'Unique CVE type counts may overlap between active and excluded devices (shown as \u2014).  '
                   '\u26a0  Unique device counts may appear lower than detection totals: a device running '
                   'Chrome, Edge and Firefox appears once per product sheet but counts as one unique device.',
                   note_fmt)
    ws.set_row(row, 54)
    row += 2
    # Build COUNTIF formula strings — these make both tables live when ☐/☑ are toggled
    _p2s = product_to_sheet or {}
    if _p2s:
        # Row A = Resolved checkbox, col C = Name, col D = Device Type (0-indexed → A,C,D)
        _f_res   = ' + '.join([f"COUNTIF('{s}'!A:A,\"☑\")" for s in _p2s.values()])
        _f_unres = ' + '.join([f"COUNTIF('{s}'!A:A,\"☐\")" for s in _p2s.values()])
        # Total = ☑ + ☐ only — COUNTA overcounts due to legend/header rows in each sheet
        _f_total = f'({_f_res}) + ({_f_unres})'
        # Device type unresolved: count rows where Resolved=☐ AND DevType matches
        _f_srv_unr = None   # device unique-count across sheets needs SUMPRODUCT — kept static
        _f_wks_unr = None
        _f_srv_all = None
        _f_wks_all = None
        _live = True
    else:
        _live = False
        _f_res   = None   # guard: prevents UnboundLocalError when no product sheets exist
        _f_unres = None
        _f_total = None

    _live_fmt  = workbook.add_format({'num_format': '#,##0', 'align': 'center',
                                      'bg_color': '#EBF3FB', 'border': 1,
                                      'font_color': '#1F3864'})  # blue tint = live formula cell
    _live_pct  = workbook.add_format({'num_format': '0%',   'align': 'center',
                                      'bg_color': '#EBF3FB', 'border': 1,
                                      'font_color': '#1F3864'})
    _live_note = '\u26a1 Blue cells update automatically when ☑/☐ are toggled in product sheets.'

    def _write_val(r, c, formula, static, fmt):
        """Write a live formula if available, otherwise static value."""
        if _live and formula:
            ws.write_formula(r, c, f'={formula}', fmt, static)
        else:
            ws.write(r, c, static, fmt)

    # ── Resolution Status (active devices only) — LIVE ──────────────────────────
    row += 1
    ws.merge_range(row, 0, row, 3, '  Resolution Status  (active devices only)', sect_fmt); row += 1
    ws.write(row, 0, 'Status',          hdr_fmt)
    ws.write(row, 1, 'Detection Rows',  hdr_fmt)
    ws.write(row, 2, '% of Total',      hdr_fmt)
    ws.write(row, 3, 'Unique CVE Types (at generation)', hdr_fmt)
    row += 1

    _rr  = int(_is_res.sum()); _ur = int(_is_unr.sum()); _tot = _rr + _ur
    _rc  = int(triage_dedup.loc[_is_res, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0
    _uc  = int(triage_dedup.loc[_is_unr, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0

    res_row = row
    ws.write(row, 0, 'Resolved',   lbl_fmt)
    _write_val(row, 1, _f_res,   _rr,  grn_fmt)
    # % = resolved / total — live
    if _live:
        ws.write_formula(row, 2, f'=IF({_f_total}>0,({_f_res})/({_f_total}),0)', _live_pct, _rr/_tot if _tot else 0)
    else:
        ws.write(row, 2, _rr/_tot if _tot else 0, val_pct)
    ws.write(row, 3, _rc, grn_fmt)
    row += 1

    ws.write(row, 0, 'Unresolved', lbl_fmt)
    _write_val(row, 1, _f_unres, _ur,  red_fmt)
    if _live:
        ws.write_formula(row, 2, f'=IF({_f_total}>0,({_f_unres})/({_f_total}),0)', _live_pct, _ur/_tot if _tot else 0)
    else:
        ws.write(row, 2, _ur/_tot if _tot else 0, val_pct)
    ws.write(row, 3, _uc, red_fmt)
    row += 1

    ws.write(row, 0, 'Total', lbl_fmt)
    _write_val(row, 1, _f_total, _tot, val_fmt)
    ws.write(row, 2, 1.0, val_pct)
    ws.write(row, 3, triage_dedup['Vulnerability Name'].nunique() if 'Vulnerability Name' in triage_dedup.columns else 0, val_fmt)
    row += 1

    ws.merge_range(row, 0, row, 3,
                   f'\u2139  {_live_note}  '
                   f'Unique CVE Type counts are fixed at report generation — they do not update live.',
                   note_fmt)
    ws.set_row(row, 36); row += 2

    # ── Device Breakdown (active scope) — unique device counts ───────────────────
    ws.merge_range(row, 0, row, 3,
                   '  Device Breakdown  (active scope \u2014 unique devices, CVE \u2265 threshold)',
                   sect_fmt)
    row += 1
    ws.write(row, 0, 'Type',                             hdr_fmt)
    ws.write(row, 1, 'Unique devices',                   hdr_fmt)
    ws.write(row, 2, 'Devices with unresolved CVEs',     hdr_fmt)
    ws.write(row, 3, 'Stale devices (excluded)',         hdr_fmt)
    row += 1

    ws.write(row, 0, 'Servers',      lbl_fmt)
    ws.write(row, 1, srv_total,      val_fmt)
    ws.write(row, 2, srv_unpatch,    red_fmt if srv_unpatch else _zero_fmt)
    ws.write(row, 3, _stale_srv,     wf_mval if _stale_srv else val_fmt)
    row += 1

    ws.write(row, 0, 'Workstations', lbl_fmt)
    ws.write(row, 1, wks_total,      val_fmt)
    ws.write(row, 2, wks_unres,      red_fmt if wks_unres else _zero_fmt)
    ws.write(row, 3, _stale_wks,     wf_mval if _stale_wks else val_fmt)
    row += 1

    ws.merge_range(row, 0, row, 3,
                   '\u2139  Counts are unique devices — a device with Chrome, Edge and Firefox counts as one device.  '
                   '"Devices with unresolved CVEs" = at least one detection still marked \u2610 across any product.  '
                   '"Stale devices (excluded)" = inactive \u226530 days, moved to Stale Excluded Devices sheet.  '
                   'All counts fixed at report generation.',
                   note_fmt)
    ws.set_row(row, 42); row += 2

    # Data Filtering Reconciliation waterfall
    # Uses deduped counts (_all_df, _stale_dedup) so [+] - [-] = [=] exactly.
    _stale_cves_dedup = int(_stale_dedup['Vulnerability Name'].nunique()) if not _stale_dedup.empty and 'Vulnerability Name' in _stale_dedup.columns else 0
    _cutoff_lbl = cutoff_date if cutoff_date else 'N/A (all dates included)'

    row += 1
    ws.merge_range(row, 0, row, 3, f'  Data Filtering Reconciliation  (CVSS \u2265 {threshold})', sect_fmt); row += 1
    ws.write(row, 0, 'Filter Step',      hdr_fmt); ws.write(row, 1, 'Unique Devices', hdr_fmt)
    ws.write(row, 2, 'Detection Rows',   hdr_fmt); ws.write(row, 3, 'Unique CVE Types', hdr_fmt); row += 1
    ws.write(row, 0, '[+]  Total detections (all devices, CVSS \u2265 threshold, deduplicated per product)', wf_plus)
    ws.write(row, 1, _all_devs, val_fmt); ws.write(row, 2, _all_rows, val_fmt); ws.write(row, 3, _all_cves, val_fmt); row += 1
    if not _stale_dedup.empty:
        _stale_dedup_devs = int(_stale_dedup['Name'].nunique()) if 'Name' in _stale_dedup.columns else _stale_devs
        ws.write(row, 0, f'[-]  Excluded: stale devices  (last seen before {_cutoff_lbl} OR \u226530 days without response)', wf_minus)
        ws.write(row, 1, _stale_dedup_devs, wf_mval); ws.write(row, 2, len(_stale_dedup), wf_mval); ws.write(row, 3, _stale_cves_dedup, wf_mval); row += 1
    if not_in_rmm_count > 0:
        ws.write(row, 0, '[-]  Excluded: device not found in RMM', wf_minus)
        ws.write(row, 1, not_in_rmm_count, wf_mval); ws.write(row, 2, not_in_rmm_cve_count, wf_mval); ws.write(row, 3, not_in_rmm_unique_cves, wf_mval); row += 1
    ws.write(row, 0, '[=]  Active tracked scope  (Key Metrics above)', wf_eq_lbl)
    ws.write(row, 1, unique_devices, wf_eq_val); ws.write(row, 2, total_rows, wf_eq_val); ws.write(row, 3, unique_cves, wf_eq_val); row += 1
    ws.merge_range(row, 0, row, 3,
                   '\u2139  All counts deduplicated per product sheet (same CVE on same device in multiple '
                   'product versions counts once per product). '
                   'Unique CVE types may not subtract exactly \u2014 a CVE on both stale and active devices '
                   'is counted in both groups. Stale devices are listed in the Stale Excluded Devices sheet.',
                   note_fmt)
    ws.set_row(row, 42); row += 1

    # ── Top At-Risk Devices ──────────────────────────────────────────────────────
    # Priority: 1) every server with ≥1 unresolved CVE  2) any device with a
    # known-exploit CVE  3) remainder by highest unresolved CVE count, up to 10.
    row += 1
    ws.merge_range(row, 0, row, 4,
                   '🚨  Top At-Risk Devices  (unresolved CVEs only)', sect_fmt)
    row += 1

    _has_uname   = 'Username'               in triage_df.columns
    _has_exploit = 'Has Known Exploit'       in triage_df.columns
    _has_dt      = 'Device Type'             in triage_df.columns
    _has_lr      = 'Last Response'           in triage_df.columns
    _has_days    = 'Days Since Last Response' in triage_df.columns
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

    ws.write(row, 0, '💻 Device Name',            _th)
    ws.write(row, 1, '👤 Username',               _th)
    ws.write(row, 2, '⚠️ Unresolved CVEs',         _th)
    ws.write(row, 3, '💣 Has Exploit',             _th)
    ws.write(row, 4, '🖥️ Device Type',             _th)
    ws.write(row, 5, '🕐 Last Response',           _th)
    ws.write(row, 6, '📅 Days Since Response',     _th)
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
                last_response=('Last Response', 'first') if _has_lr else ('Name', lambda s: ''),
                days_since  =('Days Since Last Response', lambda s:
                    pd.to_numeric(s, errors='coerce').max())
                    if _has_days else ('Name', lambda s: ''),
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

            _approaching = approaching_stale_names or set()
            _td_approach   = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00'})
            _td_approach_r = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00',
                                                   'align': 'right', 'num_format': '#,##0'})

            for _r in _top:
                _srv        = 'server' in str(_r.device_type).lower()
                _exp        = str(_r.has_exploit).strip().lower() == 'yes'
                _near_stale = _r.Name in _approaching
                if _exp:
                    _bf, _nf = _td_exp, _td_exp_r
                elif _near_stale:
                    _bf, _nf = _td_approach, _td_approach_r
                elif _srv:
                    _bf, _nf = _td_srv, _td_srv_r
                else:
                    _bf, _nf = _td, _td_r
                _name_label = f'⚠ {_r.Name}' if _near_stale else str(_r.Name)
                _days_val   = int(_r.days_since) if hasattr(_r, 'days_since') and not (isinstance(_r.days_since, float) and pd.isna(_r.days_since)) else ''
                ws.write(row, 0, _name_label,               _bf)
                ws.write(row, 1, str(_r.username),           _bf)
                ws.write(row, 2, int(_r.cve_count),          _nf)
                ws.write(row, 3, str(_r.has_exploit),        _bf)
                ws.write(row, 4, str(_r.device_type),        _bf)
                ws.write(row, 5, str(_r.last_response) if hasattr(_r, 'last_response') else '', _bf)
                ws.write(row, 6, _days_val,                  _nf)
                row += 1

            _approach_note = (
                f'  🟧 Orange = offline \u2265 {stale_warning_days}d (⚠ prefix on name).  '
                if _approaching else ''
            )
            ws.merge_range(row, 0, row, 6,
                f'ℹ  🟡 Amber = Server.  🟥 Red = known exploit.  '
                f'{_approach_note}'
                f'Up to 10 devices. Unresolved CVE counts only.',
                note_fmt)
            ws.set_row(row, 30); row += 1
        else:
            ws.merge_range(row, 0, row, 6, 'No unresolved CVE data.', note_fmt); row += 1
    else:
        ws.merge_range(row, 0, row, 6, 'No active device data.', note_fmt); row += 1

    ws.set_column('A:A', 32); ws.set_column('B:B', 22)
    ws.set_column('C:C', 18); ws.set_column('D:D', 14); ws.set_column('E:E', 16)
    ws.set_column('F:F', 24); ws.set_column('G:G', 20)

    # CVSS Score Split -- TODO: re-enable when layout is agreed
    # (block commented out; score_split_data/start/end stubs kept for
#  any downstream code that references these variables)
    score_split_start = row
    score_split_data  = []
    score_split_end   = row - 1

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
    row += 2

    # ── Top 10 Products (Score 9.0+) ────────────────────────────────────────────
    ws.merge_range(row, 0, row, 3, f'  Top 10 Products  (Score {threshold}+, active devices only)', sect_fmt)
    row += 1
    ws.write(row, 0, 'Product',          hdr_fmt)
    ws.write(row, 1, 'Devices affected', hdr_fmt)
    ws.write(row, 2, 'Unique CVEs',      hdr_fmt)
    ws.write(row, 3, 'Detections',       hdr_fmt)
    row += 1

    _p2s = product_to_sheet or {}
    if 'Base Product' in triage_df.columns:
        _prod_grp = (
            triage_df.groupby('Base Product')
            .agg(devices=('Name', 'nunique'),
                 cves=('Vulnerability Name', 'nunique'),
                 detections=('Vulnerability Name', 'count'))
            .sort_values('devices', ascending=False)
            .head(10)
        )
        _link_fmt_prod = workbook.add_format({'font_color': '#0563C1', 'underline': True,
                                              'border': 1, 'bg_color': '#FFFFFF'})
        for prod, pr in _prod_grp.iterrows():
            sheet = _p2s.get(prod)
            if sheet:
                ws.write_url(row, 0, f"internal:'{sheet}'!A1",
                             _link_fmt_prod, string=str(prod))
            else:
                ws.write(row, 0, str(prod), lbl_fmt)
            ws.write(row, 1, int(pr['devices']),    val_fmt)
            ws.write(row, 2, int(pr['cves']),        val_fmt)
            ws.write(row, 3, int(pr['detections']),  val_fmt)
            row += 1
    else:
        ws.merge_range(row, 0, row, 3, 'Base Product column not available.', note_fmt)
        row += 1
    ws.merge_range(row, 0, row, 3,
                   '\u2139  Product names are hyperlinked to their triage sheet. '
                   'Counts are for active devices only (stale/not-in-RMM excluded).',
                   note_fmt)
    ws.set_row(row, 30)

    log.debug("Summary sheet written")


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

def build_patch_sheets(writer, overview_df, full_df, patch_df):
    for df, name in ((overview_df, 'Patch Match Overview'),
                     (full_df,     'Patch Match Full Data'),
                     (patch_df,    'Patch Report (Full)')):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

# ==============================================================================