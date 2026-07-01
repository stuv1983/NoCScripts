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


# EXCEL SHEET BUILDERS
# ==============================================================================

# get_workbook_styles moved to formatting.py — re-exported here so existing
# `from excel_builder import get_workbook_styles` imports (e.g. orchestrator.py)
# keep working unchanged.
from formatting import get_workbook_styles, get_band_formats, build_legend_entries, COLORS  # noqa: E402,F401


# _write_cve_links / _write_nvd_links moved to sheet_helpers.py (shared by
# excel_builder.py and product_sheets.py without either importing the other).
from sheet_helpers import write_cve_links as _write_cve_links, write_nvd_links as _write_nvd_links  # noqa: E402,F401

# build_product_sheets, _build_patch_confirmed_sheet, and compute_score_lift
# moved to product_sheets.py — re-exported here so existing
# `from excel_builder import build_product_sheets` imports (orchestrator.py)
# keep working unchanged.
from product_sheets import build_product_sheets  # noqa: E402,F401

# build_trend_summary_sheet and build_trend_detail_sheets moved to
# trend_sheets.py — re-exported here so existing
# `from excel_builder import build_trend_summary_sheet, build_trend_detail_sheets`
# imports (orchestrator.py) keep working unchanged.
from trend_sheets import build_trend_summary_sheet, build_trend_detail_sheets  # noqa: E402,F401

# build_client_summary_sheet and compute_patching_health_score moved to
# summary_sheet.py — re-exported here so existing
# `from excel_builder import build_client_summary_sheet` imports
# (orchestrator.py) keep working unchanged.
from summary_sheet import build_client_summary_sheet, compute_patching_health_score  # noqa: E402,F401

# build_diagnostics_sheets, build_patch_resolved_sheet,
# build_products_not_tracked_sheet, build_patch_failure_sheet, and
# build_patch_sheets moved to patch_sheets.py — re-exported here so
# existing `from excel_builder import ...` imports (orchestrator.py)
# keep working unchanged.
from patch_sheets import (  # noqa: E402,F401
    build_diagnostics_sheets, build_patch_resolved_sheet,
    build_products_not_tracked_sheet, build_patch_failure_sheet,
    build_patch_sheets,
)

# build_all_detections_sheet, build_stale_excluded_sheet,
# build_stale_cves_sheet, build_device_report_sheet, and
# build_raw_data_sheet moved to device_sheets.py — re-exported here so
# existing `from excel_builder import ...` imports (orchestrator.py)
# keep working unchanged.
from device_sheets import (  # noqa: E402,F401
    build_all_detections_sheet, build_stale_excluded_sheet,
    build_stale_cves_sheet, build_device_report_sheet,
    build_raw_data_sheet,
)

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

        _band_fmt = get_band_formats(workbook)
        _nday_hdr, _nday_crit, _nday_high, _nday_amb, _nday_ok, _nday_lbl = (
            _band_fmt['header'], _band_fmt['critical'], _band_fmt['high'],
            _band_fmt['amber'], _band_fmt['ok'], _band_fmt['label'],
        )

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

# ==============================================================================