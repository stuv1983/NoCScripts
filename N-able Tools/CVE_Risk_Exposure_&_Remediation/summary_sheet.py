"""
summary_sheet.py — builds the workbook's Summary sheet: Key Metrics,
Resolution Status, Device Breakdown, Data Filtering Reconciliation,
Top At-Risk Devices, Month-over-Month Patching Progress, and the
Patching Health Score.

Resolution Status math here must stay in sync with resolution.py and
test_resolution.py — see resolution.py for the shared logic.

Author : Stu Villanti <s.villanti@kenstra.com>
"""
from datetime import datetime
from typing import Optional, Set
import logging

import pandas as pd

from sheet_helpers import hs_subtotal_ref as _hs_ref

log = logging.getLogger(__name__)

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

def build_client_summary_sheet(workbook, filtered_df, triage_df, threshold,
                               trend_data=None, customer_name='',
                               cutoff_date=None, stale_excluded_df=None,
                               not_in_rmm_count=0, not_in_rmm_cve_count=0,
                               not_in_rmm_unique_cves=0,
                               not_in_rmm_df: 'Optional[pd.DataFrame]' = None,
                               report_month='',
                               product_to_sheet: Optional[dict] = None,
                               include_health_score: bool = False,
                               patch_resolved_pairs: Optional[set] = None,
                               health_triage_df: 'Optional[pd.DataFrame]' = None,
                               health_score_threshold: float = 7.0,
                               has_patch_report: bool = False,
                               prev_report_name: str = '',
                               patch_check_active_df: 'Optional[pd.DataFrame]' = None,
                               patch_check_active_names: Optional[Set[str]] = None,
                               advanced_summary: bool = False,
                               snapshot_history: Optional[list] = None,
                               snapshot_current: Optional[dict] = None,
                               root_cause_counts: Optional[dict] = None):
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
    patch_check_active_df  — pre-built rows (Device, Device Type, Customer, Site, Check Description,
                             Last Failure, Days Since) for active devices whose RMM Patch Status Check
                             is failing (see data_pipeline.load_patch_check_report()). Already filtered
                             to active scope by the caller (orchestrator.py).
    patch_check_active_names — normalised device names (data_pipeline.normalize_device_name) of the
                             same devices, used to highlight matches inside Top At-Risk Devices below.
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
    def_fmt   = workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9,
                                      'border': 1, 'align': 'left'})
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

    ws.set_column('A:A', 44); ws.set_column('B:D', 18); ws.set_column('E:E', 48)
    title_text = (f'{customer_name}  \u2014  ' if customer_name else '') + 'CVE Risk Exposure Summary'
    ws.merge_range('A1:D1', title_text, title_fmt); ws.set_row(0, 28)
    ws.write('A2', f'Report Month: {report_month}  |  Generated: {datetime.now().strftime("%d %b %Y")}',
             workbook.add_format({'italic': True, 'font_color': '#595959', 'font_size': 9}))

    # Stale devices are excluded purely by this user-entered cutoff date —
    # there is no separate fixed day-count rule. Computed once here so every
    # section below that references the stale cutoff (Key Metrics footnote,
    # Device Breakdown footnote, Data Filtering Reconciliation) stays in sync.
    _cutoff_lbl = cutoff_date if cutoff_date else 'N/A (all dates included)'

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

    # ── Compute resolved/unresolved by replaying the exact same ☑/☐ logic
    # that build_product_sheets writes into column A of each product sheet.
    # This guarantees the cached values supplied to write_formula() match what
    # the live COUNTIF formulas will compute.
    #
    # Resolution sources — see resolution.py for the authoritative priority
    # rules. This MUST call the same function build_product_sheets uses for
    # its ☑/☐ column — that's the whole point of extracting it: the checkbox
    # column and this table can no longer silently drift apart.
    #
    # Factored into a local helper (_compute_resolved_series) so every
    # "unresolved" count on this sheet — Resolution Status, Device Breakdown,
    # AND Top At-Risk Devices — uses the identical definition. Top At-Risk
    # Devices used to recompute its own filter directly against the raw
    # Threat Status/Status column, silently ignoring patch evidence; a CVE
    # confirmed resolved by a --patch report but not yet reflected in
    # N-able's own scan would count as unresolved there while correctly
    # showing ☑ everywhere else in the workbook.
    from resolution import (split_patch_pairs as _split_patch_pairs_sum,
                            compute_resolved_series as _compute_resolved_series_shared)
    _p2s_sum   = product_to_sheet or {}
    _patch_2d_sum, _patch_3d_sum = _split_patch_pairs_sum(patch_resolved_pairs)

    def _compute_resolved_series(df: 'pd.DataFrame') -> 'pd.Series':
        """Return a bool Series aligned to df's index — True = ☑ resolved.
        Delegates to resolution.compute_resolved_series(), the single
        index-safe implementation shared with product_sheets.py — see that
        function's docstring for the row-misalignment bug this replaced."""
        return _compute_resolved_series_shared(df, _p2s_sum, patch_resolved_pairs)

    _is_res = _compute_resolved_series(triage_dedup)
    _is_unr = ~_is_res

    total_rows     = len(triage_dedup)
    unique_cves    = int(triage_dedup['Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0
    unique_devices = int(triage_dedup['Name'].nunique())               if 'Name'               in triage_dedup.columns else 0
    score_col      = 'Vulnerability Score' if 'Vulnerability Score' in triage_dedup.columns else None
    crit_mask      = pd.to_numeric(triage_dedup[score_col], errors='coerce') >= 9.0 if score_col else pd.Series([True]*len(triage_dedup), index=triage_dedup.index)
    crit_rows      = int(crit_mask.sum())
    crit_cves      = int(triage_dedup.loc[crit_mask, 'Vulnerability Name'].nunique()) if score_col and 'Vulnerability Name' in triage_dedup.columns else unique_cves


    # CISA KEV — Known Exploited Vulnerabilities catalog. Distinct from the
    # generic 'Has Known Exploit' column above: KEV is CISA's authoritative
    # "actively exploited in the wild" list, so it gets its own Key Metrics
    # rows and its own product/device tracking tables (see below).
    kev_col        = 'CISA KEV' if 'CISA KEV' in triage_dedup.columns else None
    kev_mask       = triage_dedup[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y']) if kev_col else pd.Series([False]*len(triage_dedup), index=triage_dedup.index)
    kev_rows       = int(kev_mask.sum())
    kev_cves       = int(triage_dedup.loc[kev_mask, 'Vulnerability Name'].nunique()) if kev_col and 'Vulnerability Name' in triage_dedup.columns else 0

    # Device-type counts for Device Breakdown sub-table (unique devices)
    if 'Device Type' in triage_dedup.columns and 'Name' in triage_dedup.columns:
        _srv_mask = triage_dedup['Device Type'].astype(str).str.lower().str.contains('server',      na=False)
        _wks_mask = triage_dedup['Device Type'].astype(str).str.lower().str.contains('workstation', na=False)
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
    _stale_devs = int(stale_excluded_df['Name'].nunique()) if stale_excluded_df is not None and not stale_excluded_df.empty and 'Name' in stale_excluded_df.columns else 0
    # NIRM (not_in_rmm_cve_count already passed in as detection rows; compute CVSS 9+ subset from filtered_df)
    _nirm_devs  = not_in_rmm_count
    _nirm_mask  = (filtered_df['Last Response'] == 'Not Found in RMM') \
                  if 'Last Response' in filtered_df.columns \
                  else pd.Series(False, index=filtered_df.index)
    _nirm_crit  = 0
    if score_col and score_col in filtered_df.columns:
        _nirm_sc   = pd.to_numeric(filtered_df.loc[_nirm_mask, score_col], errors='coerce')
        _nirm_crit = int((_nirm_sc >= 9.0).sum())

    # Stale + NIRM CISA KEV counts, mirroring the CVSS 9+ pattern above so
    # KEV gets the same All / Active / Excluded reconciliation.
    _nirm_kev  = 0
    if kev_col and kev_col in filtered_df.columns:
        _nirm_kev_mask = filtered_df[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y'])
        _nirm_kev      = int((_nirm_mask & _nirm_kev_mask).sum())

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

    # Compute "All" totals — deduped the same way as triage_dedup so All = Active + Excluded.
    # See resolution.build_all_scope_frame() / dedup_per_base_product() for
    # why this must not filter stale/not-in-RMM rows by product_to_sheet
    # membership — that filter existed here before and silently dropped
    # stale-only or not-in-RMM-only products from the "All" totals.
    from resolution import build_all_scope_frame as _build_all_scope_frame, dedup_per_base_product as _dedup_pbp
    _stale_dedup = _dedup_pbp(stale_excluded_df)
    _nirm_dedup  = _dedup_pbp(not_in_rmm_df)
    _all_df = _build_all_scope_frame(triage_dedup, stale_excluded_df, not_in_rmm_df)

    _all_rows      = len(_all_df)          # = total_rows + stale_dedup_rows — math correct
    _all_cves      = int(_all_df['Vulnerability Name'].nunique()) if 'Vulnerability Name' in _all_df.columns else 0
    _all_devs      = int(_all_df['Name'].nunique())               if 'Name'               in _all_df.columns else 0

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

    # Same full-stale-frame pattern for CISA KEV, so a stale device running a
    # KEV-flagged product absent from active triage isn't silently dropped
    # from the "All" KEV totals.
    _stale_full_kev_mask = (
        _stale_full_dedup[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y'])
        if kev_col and not _stale_full_dedup.empty and kev_col in _stale_full_dedup.columns
        else pd.Series(dtype=bool)
    )
    _stale_full_kev      = int(_stale_full_kev_mask.sum())
    _stale_full_kev_cves = (
        int(_stale_full_dedup.loc[_stale_full_kev_mask, 'Vulnerability Name'].nunique())
        if kev_col and 'Vulnerability Name' in _stale_full_dedup.columns and not _stale_full_dedup.empty else 0
    )
    _all_kev      = kev_rows + _stale_full_kev
    _all_kev_cves = kev_cves + _stale_full_kev_cves

    # Uses the deduped not-in-RMM row count (matches how _all_rows is now built)
    # when the actual dataframe was passed; falls back to the pre-computed
    # scalar for callers that haven't been updated to pass not_in_rmm_df yet.
    _nirm_excl_rows    = len(_nirm_dedup) if not _nirm_dedup.empty else not_in_rmm_cve_count
    _excl_rows          = len(_stale_dedup) + _nirm_excl_rows
    _excl_devs_tot      = _stale_devs + _nirm_devs
    _excl_crit_tot      = _stale_full_crit + _nirm_crit        # full stale, not _p2s_keys-filtered
    _excl_kev_tot       = _stale_full_kev + _nirm_kev          # full stale, not _p2s_keys-filtered

    # ── Patching Health Score (beta) ──────────────────────────────────────────
    # Only rendered when include_health_score=True (opt-in checkbox in the GUI).
    if include_health_score:
        # Build a separate dedup from health_triage_df (CVSS ≥ 7.0 by default) so the
        # critical-coverage component (CVSS ≥ 9) is genuinely distinct from resolution rate.
        # Falls back to triage_dedup when no broader scope was supplied.
        _health_raw = health_triage_df if (health_triage_df is not None and not health_triage_df.empty) else None
        if _health_raw is not None:
            # Was: only kept Base Products already present in product_to_sheet
            # — but product_to_sheet is built from triage_df, which uses the
            # report's OWN (narrower) threshold, while health_triage_df is
            # deliberately broader (CVSS ≥ 7.0 even when the report threshold
            # is 9.0). A product with only 7.0–8.9 rows and nothing at the
            # report's own threshold would never be a product_to_sheet key,
            # so its rows were silently dropped from the Health Score scope
            # entirely — even though the Health Score explicitly claims to
            # use the broader CVSS ≥ 7.0 scope. dedup_per_base_product()
            # includes every Base Product unconditionally; compute_resolved_series()
            # no longer needs product_to_sheet to resolve a group correctly
            # (see resolution.py), so there's no reason to pre-filter here.
            _score_scope = _dedup_pbp(_health_raw)

            # Resolved flags — reuses the same _compute_resolved_series helper as
            # the Resolution Status table and Top At-Risk Devices (see above).
            # This used to be reimplemented independently here, and that copy
            # flattened every patch-evidence pair to a bare (device, cve) 2-tuple
            # before checking membership — discarding the product-scoping that
            # a 3-tuple (device, cve, product) pair is supposed to have. A patch
            # confirmed resolved only for one product could incorrectly mark the
            # same device+CVE resolved under a different product too, inflating
            # the resolution-rate component above what the product sheets show.
            _hs_is_res = _compute_resolved_series(_score_scope)
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

        # health_triage_df is deliberately broader than the report's own
        # threshold (CVSS ≥ 7.0 even when the report itself is ≥ 9.0), so
        # _score_scope can contain rows that exist on NO product sheet:
        #   (a) whole products with nothing at the report threshold — no
        #       sheet, no entry in product_to_sheet at all; and
        #   (b) 7.0–8.9 rows of products that DO have a sheet, because
        #       product sheets are built from triage_df at the report's own
        #       threshold, not the health scope.
        # None of these rows has a ☑/☐ checkbox, so their resolved state can
        # never change inside Excel — it is frozen at generation time. So
        # rather than disabling live scoring (the old behaviour), their
        # generation-time counts are folded into the fleet sums as plain
        # numeric constants. The live score then equals the Python score
        # exactly at generation, and stays live for every row the reader can
        # actually toggle.
        #
        # The one case constants can NOT fix is the reverse direction: a
        # report threshold BELOW the health scope (e.g. 1.0) puts sub-7.0
        # rows — with toggleable checkboxes — onto the sheets, and the
        # per-sheet subtotal cells would count them into the health score as
        # they're toggled. Live scoring stays disabled there.
        _hs_residual = {'hs_res': 0, 'hs_unres': 0, 'crit_res': 0,
                        'crit_unres': 0, 'exp_res': 0, 'exp_unres': 0,
                        'kev_unres': 0}
        _hs_key_cols = ['Base Product', 'Name', 'Vulnerability Name']
        # A report threshold BELOW the health scope (e.g. 1.0) puts sub-7.0
        # rows with toggleable checkboxes on the sheets — but the fleet sums
        # only reference the scoped subtotal cells (R5-R9 carry the
        # 'Vulnerability Score >= 7' criterion inside their COUNTIFS), so
        # those rows can never feed the score and live scoring stays on at
        # ANY report threshold (v0.35).
        if _hs_live and not all(c in _score_scope.columns for c in _hs_key_cols):
            log.info(
                "Patching Health Score live formulas disabled: health scope is missing "
                "one of %s, so off-sheet rows can't be matched. Static score values "
                "will be written instead.", _hs_key_cols,
            )
            _hs_live = False
        elif _hs_live:
            # Rows present on a product sheet = triage_dedup rows for products
            # that have a sheet (product sheets dedup by Name + Vulnerability
            # Name per product, which is exactly how triage_dedup is built).
            _on_sheet = triage_dedup[triage_dedup['Base Product'].isin(set((product_to_sheet or {}).keys()))] \
                if 'Base Product' in triage_dedup.columns else triage_dedup.iloc[0:0]
            _sheet_idx = pd.MultiIndex.from_frame(_on_sheet[_hs_key_cols].astype(str)) \
                if not _on_sheet.empty else pd.MultiIndex.from_arrays([[], [], []], names=_hs_key_cols)
            _off_mask = ~pd.MultiIndex.from_frame(_score_scope[_hs_key_cols].astype(str)).isin(_sheet_idx)
            _off_mask = pd.Series(_off_mask, index=_score_scope.index)
            if _off_mask.any():
                if 'Vulnerability Score' in _score_scope.columns:
                    _hcm = pd.to_numeric(_score_scope['Vulnerability Score'], errors='coerce') >= 9.0
                else:
                    _hcm = pd.Series(False, index=_score_scope.index)
                if 'Has Known Exploit' in _score_scope.columns:
                    _hem = _score_scope['Has Known Exploit'].astype(str).str.strip().str.lower().isin(
                        ['yes', 'true', '1', 'y'])
                else:
                    _hem = pd.Series(False, index=_score_scope.index)
                if 'CISA KEV' in _score_scope.columns:
                    _hkm = _score_scope['CISA KEV'].astype(str).str.strip().str.lower().isin(
                        ['yes', 'true', '1', 'y'])
                else:
                    _hkm = pd.Series(False, index=_score_scope.index)
                _hs_residual = {
                    # Off-sheet rows are all within the health scope by
                    # construction, so they carry the hs_* keys.
                    'hs_res':     int((_off_mask & _hs_is_res).sum()),
                    'hs_unres':   int((_off_mask & _hs_is_unr).sum()),
                    'crit_res':   int((_off_mask & _hcm & _hs_is_res).sum()),
                    'crit_unres': int((_off_mask & _hcm & _hs_is_unr).sum()),
                    'exp_res':    int((_off_mask & _hem & _hs_is_res).sum()),
                    'exp_unres':  int((_off_mask & _hem & _hs_is_unr).sum()),
                    # Off-sheet unresolved KEV rows can never be ticked, so if
                    # any exist the live KEV cell $E$11 correctly never reaches
                    # 0 and the grade cap/penalty correctly never lift.
                    'kev_unres':  int((_off_mask & _hkm & _hs_is_unr).sum()),
                }
                log.info(
                    "Patching Health Score: %d health-scope row(s) exist on no product "
                    "sheet (broader CVSS ≥ %.1f scope vs report threshold %.1f, plus any "
                    "products with no sheet). They have no checkboxes and can never "
                    "change in Excel, so their counts are folded into the live fleet "
                    "sums as fixed constants. Live scoring remains enabled.",
                    int(_off_mask.sum()), health_score_threshold, threshold,
                )

        if _hs_live:
            # ── Live totals via per-sheet subtotal cells ─────────────────────────
            # Each product sheet carries a hidden subtotal block (col R rows 1-7,
            # written by product_sheets via sheet_helpers.write_hs_subtotals):
            # short LOCAL COUNTIF/COUNTIFS over that sheet only, so toggling ☑/☐
            # still flows through live.  The fleet totals here are therefore one
            # short cell reference per sheet ('Sheet'!$R$1 + ...) instead of a
            # full COUNTIFS per sheet — the old approach produced 20k+ character
            # formulas on big fleets, far above Excel's 8,192-char stored-formula
            # limit, which permanently forced the score back to static values.
            #
            # The six fleet sums are written ONCE into hidden helper cells
            # E5:E10 (next to the E4 score helper).  Every visible formula
            # references $E$5..$E$10, so visible formulas stay constant-size no
            # matter how many product sheets exist.  A second win: each sheet's
            # subtotals are built from its OWN column layout — Patch Confirmed
            # sheets have no Score Lift column, so the old Summary-side G:G/I:I
            # hard-coding silently tested the wrong columns there.
            # Each sum is one short cell ref per sheet, plus (when the health
            # scope contains products with no sheet — see _hs_residual above)
            # a single numeric constant carrying those frozen rows.
            def _fleet_sum(_key):
                _s = ' + '.join(_hs_ref(sn, _key) for sn in _p2s_hs_vals)
                if _hs_residual.get(_key):
                    _s = f"{_s} + {_hs_residual[_key]}"
                return _s

            # 'hs_res'/'hs_unres' (R8/R9) — NOT the whole-sheet R1/R2, which
            # include sub-scope rows when the report threshold is below the
            # health scope. R1/R2 remain the Resolution Status table's source.
            _f_sum_res        = _fleet_sum('hs_res')
            _f_sum_unres      = _fleet_sum('hs_unres')
            _f_sum_crit_res   = _fleet_sum('crit_res')
            _f_sum_crit_unres = _fleet_sum('crit_unres')
            _f_sum_exp_res    = _fleet_sum('exp_res')
            _f_sum_exp_unres  = _fleet_sum('exp_unres')
            # Live count of unresolved CISA KEV rows (per-sheet R7 cells plus
            # the off-sheet constant) — drives the live grade cap and the KEV
            # penalty lift below via hidden helper cell $E$11.
            _f_sum_kev_unres  = _fleet_sum('kev_unres')

            # Helper cell layout (col E = index 4, rows 4-11 in Excel terms):
            #   E4  live score          E5  Σ res        E6  Σ unres
            #   E7  Σ crit res          E8  Σ crit unres
            #   E9  Σ exp res           E10 Σ exp unres
            #   E11 Σ unresolved CISA KEV rows (drives cap/penalty lift)
            _helper_row = 3
            _helper_col = 4
            _helper_ref = '$E$4'
            _kev_live_ref = '$E$11'
            _hr = {'res': '$E$5', 'unres': '$E$6', 'crit_res': '$E$7',
                   'crit_unres': '$E$8', 'exp_res': '$E$9', 'exp_unres': '$E$10'}

            _f_hs_res     = _hr['res']
            _f_hs_total   = f"({_hr['res']}+{_hr['unres']})"
            _f_crit_res   = _hr['crit_res']
            _f_crit_total = f"({_hr['crit_res']}+{_hr['crit_unres']})"
            _f_exp_res    = _hr['exp_res']
            _f_exp_total  = f"({_hr['exp_res']}+{_hr['exp_unres']})"

            # Static penalty values (cannot be recalculated from the sheet columns)
            _pen_persist_pts = _phs_pens.get('persisting_cves', {}).get('pts', 0.0)
            _pen_kev_pts     = _phs_pens.get('kev_unresolved',  {}).get('pts', 0.0)

            # -- Live score formula -------------------------------------------------
            # Mirrors compute_patching_health_score:
            #   pts_res  = IF(total>0, res/total, 0) * 60
            #   pts_crit = IF(crit_total>0, crit_res/crit_total, 1) * 20
            #   pts_exp  = IF(exp_total>0,  exp_res/exp_total,   1) * 20
            #   score    = MIN(cap, MAX(0, INT(ROUND(pts - penalties, 0))))
            #
            # The persisting-CVE penalty is genuinely static (depends on trend
            # data that isn't in the workbook). The KEV penalty is conditionally
            # live: the Python penalty counts unique unresolved KEV CVE *types*
            # (−1 each, cap −5), which no in-sheet formula can deduplicate — but
            # the boundary IS exact: unresolved KEV rows = 0 ⇔ unresolved KEV
            # types = 0. So `IF($E$11>0, pts, 0)` keeps the generation-time
            # penalty for intermediate states and lifts it precisely when the
            # last KEV row is ticked ☑ (or keeps it forever if any unresolved
            # KEV row is off-sheet, where it can never be ticked — correct).
            _f_pts_res  = f"IF(({_f_hs_total})>0,({_f_hs_res})/({_f_hs_total}),0)*60"
            _f_pts_crit = f"IF(({_f_crit_total})>0,({_f_crit_res})/({_f_crit_total}),1)*20"
            _f_pts_exp  = f"IF(({_f_exp_total})>0,({_f_exp_res})/({_f_exp_total}),1)*20"
            _f_pen      = f"({_pen_persist_pts}+IF({_kev_live_ref}>0,{_pen_kev_pts},0))"
            _f_raw      = f"({_f_pts_res})+({_f_pts_crit})+({_f_pts_exp})-{_f_pen}"

            # KEV grade caps — compute_patching_health_score caps the numeric
            # score at 74 (KEV ≥ 3, or any KEV with no patch report) / 89 (any
            # unresolved KEV).  The cap *tier* (74 vs 89) stays a generation-time
            # constant, but *whether* it applies is live: it lifts the moment
            # the live unresolved-KEV count $E$11 reaches 0 — i.e. every KEV
            # row on every sheet is ticked ☑ and none exist off-sheet. Without
            # the live condition, a fully-remediated workbook stayed pinned at
            # the ceiling (74/89) no matter what the reader resolved.
            _kev_cnt_live = _phs_pens.get('kev_unresolved', {}).get('count', 0)
            if _kev_cnt_live >= 3 or (_kev_cnt_live > 0 and not _phs_conf.get('has_patch_report')):
                _score_cap = 74
            elif _kev_cnt_live > 0:
                _score_cap = 89
            else:
                _score_cap = None
            _f_score = f"MAX(0,INT(ROUND({_f_raw},0)))"
            if _score_cap is not None:
                _f_score = f"MIN(IF({_kev_live_ref}>0,{_score_cap},100),{_f_score})"

            # Nested IF rather than IFS(): IFS is an Excel 2019+ "future
            # function" that xlsxwriter writes verbatim — Excel stores such
            # functions internally as _xlfn.IFS, so a bare IFS renders as
            # #NAME?.  Nested IF is equivalent and works in every Excel and
            # LibreOffice version.
            _f_grade = (
                f'IF({_helper_ref}>=90,"A",IF({_helper_ref}>=75,"B",'
                f'IF({_helper_ref}>=60,"C",IF({_helper_ref}>=40,"D","F"))))'
            )
            _f_score_ref = f'={_helper_ref}'

            # Excel's stored formula limit is 8,192 characters.  With the
            # subtotal-cell design the only formulas that still grow with the
            # sheet count are the six helper-cell sums (one short cell ref per
            # sheet, ~25 chars each) — it would take ~250 max-length sheet
            # names to trip this, but keep the guard as a safety net: an
            # over-limit formula makes Excel "repair" (strip) formulas on open.
            _candidate_live_formulas = [
                _f_sum_res, _f_sum_unres, _f_sum_crit_res,
                _f_sum_crit_unres, _f_sum_exp_res, _f_sum_exp_unres,
                _f_sum_kev_unres, _f_score, _f_grade, _f_score_ref,
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
        # Score box and grade box background (v0.29):
        #   live   → neutral dark blue; the formula-based conditional-format
        #            rules below recolour both boxes from the live $E$4 score.
        #   static → the generation-time grade colour (_phs_colour) — no
        #            conditional formatting exists in static mode, so the box
        #            must carry the correct colour itself.
        _live_score_bg = '#1F4E79'   # neutral dark blue for the live score container
        _box_bg        = _live_score_bg if _hs_live else _phs_colour
        _score_box_fmt = workbook.add_format({
            'bold': True, 'font_size': 36, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': _box_bg, 'border': 2,
        })
        _grade_box_fmt = workbook.add_format({
            'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': _box_bg, 'border': 2,
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
        # merge_range cannot store a cached formula result, so readers that
        # don't recalculate on open (data_only tooling, some viewers) would
        # display 0 in the box. Merge a blank, then overwrite the top-left
        # cell via write_formula with the correct cached result — the
        # documented xlsxwriter pattern for formulas in merged ranges; this
        # write order produces a single cell entry in the sheet XML (v0.29).
        _score_box_row = row   # remember for conditional formatting below
        if _hs_live:
            # Hidden helper cells: E4 = live score; E5:E10 = the six fleet
            # sums over the per-sheet subtotal cells.  All visible cells
            # reference these, so every visible formula stays tiny.
            _helper_fmt = workbook.add_format({'num_format': '0', 'font_color': 'white',
                                               'bg_color': 'white'})
            # Cached generation-time values so data_only readers see numbers
            # (Excel recalculates the formulas on open regardless).
            _cache_res   = int(_hs_is_res.sum())
            _cache_unres = int(_hs_is_unr.sum())
            if 'Vulnerability Score' in _score_scope.columns:
                _cm = pd.to_numeric(_score_scope['Vulnerability Score'], errors='coerce') >= 9.0
            else:
                _cm = pd.Series(False, index=_score_scope.index)
            if 'Has Known Exploit' in _score_scope.columns:
                _em = _score_scope['Has Known Exploit'].astype(str).str.strip().str.lower().isin(
                    ['yes', 'true', '1', 'y'])
            else:
                _em = pd.Series(False, index=_score_scope.index)
            if 'CISA KEV' in _score_scope.columns:
                _km = _score_scope['CISA KEV'].astype(str).str.strip().str.lower().isin(
                    ['yes', 'true', '1', 'y'])
            else:
                _km = pd.Series(False, index=_score_scope.index)
            for _hrow, _hf, _hcache in [
                (4,  _f_sum_res,        _cache_res),
                (5,  _f_sum_unres,      _cache_unres),
                (6,  _f_sum_crit_res,   int((_cm & _hs_is_res).sum())),
                (7,  _f_sum_crit_unres, int((_cm & _hs_is_unr).sum())),
                (8,  _f_sum_exp_res,    int((_em & _hs_is_res).sum())),
                (9,  _f_sum_exp_unres,  int((_em & _hs_is_unr).sum())),
                (10, _f_sum_kev_unres,  int((_km & _hs_is_unr).sum())),
            ]:
                ws.write_formula(_hrow, _helper_col, f'={_hf}', _helper_fmt, _hcache)
            ws.write_formula(_helper_row, _helper_col, f'={_f_score}',
                             _helper_fmt, _phs_score)
            ws.merge_range(row, 0, row + 1, 1, '', _score_box_fmt)
            ws.write_formula(row, 0, _f_score_ref, _score_box_fmt, _phs_score)
        else:
            ws.merge_range(row, 0, row + 1, 1, _phs_score, _score_box_fmt)

        # Grade box (col C) -- same pattern
        if _hs_live:
            ws.merge_range(row, 2, row + 1, 2, '', _grade_box_fmt)
            ws.write_formula(row, 2, f'={_f_grade}', _grade_box_fmt, _phs_grade)
        else:
            ws.merge_range(row, 2, row + 1, 2, _phs_grade, _grade_box_fmt)

        ws.merge_range(row, 3, row + 1, 3, 'Patching Health Score  (0\u2013100)', _score_lbl_fmt)
        row += 2

        # ── Conditional formatting: colour score/grade boxes by live score value ─
        # Applied to the top-left cell of each merged region (xlsxwriter targets the
        # top-left cell; Excel applies the format to the entire merged range).
        if _hs_live:
            # Formula-based rules on the live $E$4 score — the previous
            # numeric 'cell between' rules could never match the grade cell's
            # TEXT value ("A".."F"), so the grade box stayed neutral dark blue
            # forever.  A formula rule evaluates $E$4 regardless of what the
            # formatted cell itself contains, so score box and grade box
            # recolour together.
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
                    'type': 'formula',
                    'criteria': f'=AND($E$4>={_cf_min},$E$4<={_cf_max})',
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

        # ── Penalty rows ────────────────────────────────────────────────────────
        # Persisting-CVE penalty is static (depends on trend data outside the
        # workbook). The KEV penalty display is live when the score is live:
        # =IF($E$11>0, −pts, 0) — it visibly clears alongside the score/cap
        # the moment the last KEV row is ticked ☑ (see the live-score formula
        # comment above for why the boundary is exact).
        for _plbl, _pkey, _punit, _plive in [
            ('Persisting CVE types (\u22120.5 each, max \u22125)  \u2013 fixed at generation',
             'persisting_cves', 'CVE types', False),
            ('Unresolved CISA KEV CVEs (\u22121 each, max \u22125)  \u2013 lifts when all KEV rows are \u2611'
             if _hs_live else
             'Unresolved CISA KEV CVEs (\u22121 each, max \u22125)  \u2013 fixed at generation',
             'kev_unresolved',  'KEV CVEs', True),
        ]:
            _pen  = _phs_pens.get(_pkey, {})
            _ppts = _pen.get('pts', 0)
            _pcnt = _pen.get('count', 0)
            ws.write(row, 0, _plbl,               _comp_lbl_italic_fmt)
            ws.write(row, 1, f'{_pcnt} {_punit}', _comp_lbl_italic_fmt)
            if _hs_live and _plive and _ppts > 0:
                ws.write_formula(row, 2,
                    f'=IF({_kev_live_ref}>0,{-_ppts},0)',
                    _pen_neg_fmt, -_ppts)
            else:
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
            'The KEV grade cap and KEV penalty lift when every KEV row is \u2611; '
            'the persisting-CVE penalty is fixed at report generation (depends on trend data).  '
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

    # ── Multi-Month Trend (advanced preview) ────────────────────────────────────
    # Built from snapshots/ history next to the output file — accumulates
    # across runs without needing previous workbooks. Advanced-only while
    # the layout settles.
    if advanced_summary and (snapshot_history or snapshot_current):
        _adv_hdr = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
        _adv_bold = workbook.add_format({'bold': True})
        ws.merge_range(row, 0, row, 6, '  Multi-Month Trend  \u26a0 advanced preview', sect_fmt)
        row += 1
        for _c, _h in enumerate(['Month', 'Health Score', 'Grade', 'Unique CVEs',
                                 'Unique Devices', 'Unresolved KEV CVEs', 'Unresolved pairs']):
            ws.write(row, _c, _h, _adv_hdr)
        row += 1

        def _snap_val(rec, key):
            v = rec.get(key)
            return v if v is not None else '\u2014'

        _SNAP_KEYS = ['health_score', 'health_grade', 'unique_cves',
                      'unique_devices', 'kev_unresolved_cves', 'unresolved_pairs']
        for _rec in (snapshot_history or []):
            ws.write(row, 0, str(_rec.get('report_month') or _rec.get('run_date', '')[:7]), lbl_fmt)
            for _c, _k in enumerate(_SNAP_KEYS, start=1):
                ws.write(row, _c, _snap_val(_rec, _k), val_fmt)
            row += 1
        if snapshot_current:
            ws.write(row, 0, f"{snapshot_current.get('report_month', '')} (this run)", _adv_bold)
            for _c, _k in enumerate(_SNAP_KEYS, start=1):
                ws.write(row, _c, _snap_val(snapshot_current, _k), _adv_bold)
            row += 1
        ws.merge_range(row, 0, row, 6,
                       '\u2139  Built from the snapshots/ folder next to the output file. '
                       'Months generated before this feature (or with it off) show \u2014 '
                       'for score fields.',
                       note_fmt)
        row += 2

    # ── Month-over-Month Remediation Summary ────────────────────────────────────
    # Only rendered when a previous report was supplied. The point of this
    # section is to make remediation WORK visible in one place — the Resolution Status table below only shows the current
    # snapshot, which can make a month of real patching look like "only
    # dropped by a few hundred rows" if new detections landed at the same
    # time.
    #
    # Previous / Current / Cleared / New are all computed from the same
    # full-scope device+CVE+product population in compute_trends(), so
    # Previous - Cleared + New = Current exactly.
    if trend_data:
        _mom = trend_data.get('metrics', {})
        _mom_up_fmt   = workbook.add_format({'bold': True, 'font_color': '#C00000'})
        _mom_down_fmt = workbook.add_format({'bold': True, 'font_color': '#375623'})
        _mom_same_fmt = workbook.add_format({'font_color': '#595959'})

        def _mom_change(prev_val, cur_val, lower_is_better=True):
            diff = cur_val - prev_val
            if diff == 0:
                return '\u2014  no change', _mom_same_fmt
            if (diff < 0) == lower_is_better:
                return f'\u25bc  {abs(diff):,}', _mom_down_fmt
            return f'\u25b2  {abs(diff):,}', _mom_up_fmt

        ws.merge_range(row, 0, row, 4, '  Month-over-Month Remediation Summary', sect_fmt)
        row += 1
        ws.write(row, 0, f'Compared against: {prev_report_name}' if prev_report_name
                         else 'Compared against previous report', note_fmt)
        row += 1
        ws.write(row, 0, 'Metric',      hdr_fmt)
        ws.write(row, 1, 'Previous',    hdr_fmt)
        ws.write(row, 2, 'Current',     hdr_fmt)
        ws.write(row, 3, 'Change',      hdr_fmt)
        ws.write(row, 4, 'Definition',  hdr_fmt)
        row += 1

        _prev_pairs = int(_mom.get('previous_unresolved_pair_count', 0))
        _cur_pairs  = int(_mom.get('current_unresolved_pair_count', 0))
        ws.write(row, 0, 'Unresolved device/CVE pairs', lbl_fmt)
        ws.write(row, 1, _prev_pairs, val_fmt)
        ws.write(row, 2, _cur_pairs,  val_fmt)
        _s, _f = _mom_change(_prev_pairs, _cur_pairs, lower_is_better=True)
        ws.write(row, 3, _s, _f)
        ws.write(row, 4, 'Distinct device + CVE pairs still unresolved.', def_fmt)
        row += 1

        _prev_devs = int(_mom.get('previous_unresolved_device_count', 0))
        _cur_devs  = int(_mom.get('current_unresolved_device_count', 0))
        ws.write(row, 0, 'Devices with unresolved CVEs', lbl_fmt)
        ws.write(row, 1, _prev_devs, val_fmt)
        ws.write(row, 2, _cur_devs,  val_fmt)
        _s, _f = _mom_change(_prev_devs, _cur_devs, lower_is_better=True)
        ws.write(row, 3, _s, _f)
        ws.write(row, 4, 'Devices that have at least one unresolved CVE.', def_fmt)
        row += 1

        _cleared_count = int(_mom.get('cleared_previous_unresolved_count', 0))
        _cleared_pct   = float(_mom.get('cleared_previous_unresolved_pct', 0.0))
        ws.write(row, 0, 'Previous-period pairs now cleared', lbl_fmt)
        ws.write(row, 1, '\u2014', val_fmt)
        ws.write(row, 2, _cleared_count, val_fmt)
        ws.write(row, 3, _cleared_pct, val_pct)
        ws.write(row, 4, 'Pairs unresolved last report that are resolved now.', def_fmt)
        row += 1

        _new_count = int(_mom.get('new_unresolved_pair_count', 0))
        ws.write(row, 0, 'New unresolved pairs introduced', lbl_fmt)
        ws.write(row, 1, '\u2014', val_fmt)
        ws.write(row, 2, _new_count, val_fmt)
        ws.write(row, 3, f'\u25b2  {_new_count:,}' if _new_count else '\u2014',
                 _mom_up_fmt if _new_count else _mom_same_fmt)
        ws.write(row, 4, 'Pairs that became unresolved for the first time this period.', def_fmt)
        row += 1

        # Advanced: CVE-type level movement (pairs above are device+CVE level).
        if advanced_summary:
            for _lbl, _key, _defn in [
                ('New CVE types introduced', 'new_cve_count',
                 'Distinct CVE IDs seen this period but not last (common-product scope).'),
                ('CVE types resolved', 'resolved_cve_count',
                 'Distinct CVE IDs present last period and gone or resolved now.'),
                ('Persisting CVE types', 'persisting_cve_count',
                 'Distinct CVE IDs unresolved in both periods.'),
            ]:
                _v = int(_mom.get(_key, 0))
                ws.write(row, 0, _lbl, lbl_fmt)
                ws.write(row, 1, '\u2014', val_fmt)
                ws.write(row, 2, _v, val_fmt)
                ws.write(row, 3, '\u2014', _mom_same_fmt)
                ws.write(row, 4, _defn, def_fmt)
                row += 1
        row += 1

    ws.merge_range(row, 0, row, 4, '  Key Metrics', sect_fmt)
    row += 1

    # Header row
    ws.write(row, 0, 'Metric',       hdr_fmt)
    ws.write(row, 1, 'All',          _all_hdr)
    ws.write(row, 2, 'Active only',  hdr_fmt)
    ws.write(row, 3, 'Excl. (stale / not in RMM)', _excl_hdr)
    ws.write(row, 4, 'Definition',   hdr_fmt)
    row += 1

    _KEY_METRIC_DEFINITIONS = {
        'Total detection rows':       'One row per device + CVE + product detection.',
        'Unique CVE types':           'Distinct CVE IDs, regardless of how many devices have them.',
        'Unique devices':             'Distinct devices with at least one qualifying detection.',
        'Detections at CVSS 9.0+':    'Detection rows scoring 9.0 or higher.',
        'Unique CVEs at CVSS 9.0+':   'Distinct CVE IDs scoring 9.0 or higher.',
        'Detections with CISA KEV':   'Detection rows flagged in the CISA Known Exploited Vulnerabilities (KEV) catalog.',
        'Unique CVEs with CISA KEV':  'Distinct CVE IDs flagged in the CISA KEV catalog.',
    }

    for metric, all_val, active_val, excl_val, active_fmt in [
        ('Total detection rows',        _all_rows,          total_rows,     _excl_rows,             val_fmt),
        ('Unique CVE types',            _all_cves,          unique_cves,    0,                      val_fmt),   # CVEs overlap — excl shown as n/a
        ('Unique devices',              _all_devs,          unique_devices, _excl_devs_tot,         val_fmt),
        ('Detections at CVSS 9.0+',    _all_crit,          crit_rows,      _excl_crit_tot,         red_fmt if crit_rows else val_fmt),
        ('Unique CVEs at CVSS 9.0+',   _all_crit_cves,     crit_cves,      0,                      red_fmt if crit_cves else val_fmt),   # overlap
        ('Detections with CISA KEV',   _all_kev,           kev_rows,       _excl_kev_tot,          red_fmt if kev_rows else val_fmt),
        ('Unique CVEs with CISA KEV',  _all_kev_cves,      kev_cves,       0,                      red_fmt if kev_cves else val_fmt),   # overlap
    ]:
        ws.write(row, 0, metric,      lbl_fmt)
        ws.write(row, 1, all_val,     _all_val)
        ws.write(row, 2, active_val,  active_fmt)
        ws.write(row, 3, excl_val if excl_val else '—',  _excl_val)
        ws.write(row, 4, _KEY_METRIC_DEFINITIONS.get(metric, ''), def_fmt)
        row += 1


    ws.merge_range(row, 0, row, 4,
                   '\u2139  All = full dataset at CVSS \u2265 threshold.  '
                   f'Active = excludes stale devices (last seen before {_cutoff_lbl}) and devices not found in RMM.  '
                   'Active + Excluded = All for row counts and device counts.  '
                   'Unique CVE type counts may overlap between active and excluded devices (shown as \u2014).  '
                   '\u26a0  Unique device counts may appear lower than detection totals: a device running '
                   'Chrome, Edge and Firefox appears once per product sheet but counts as one unique device.',
                   note_fmt)
    ws.set_row(row, 54)
    row += 2

    # ── Products with Known Exploited Vulnerabilities (CISA KEV) ────────────────
    # Two tables split by device type: workstations are summarised by product
    # (compact — fleets can have hundreds), servers are listed with the actual
    # device names since server-level detail matters more for a smaller,
    # higher-value population. Both link the product name to its triage sheet,
    # matching the Top 10 Products table below.
    _kev_link_fmt = workbook.add_format({'font_color': '#0563C1', 'underline': True,
                                         'border': 1, 'bg_color': '#FFFFFF'})
    _kev_td       = workbook.add_format({'border': 1})
    _kev_td_r     = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})
    _kev_td_red   = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6'})
    _kev_td_red_r = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6',
                                          'align': 'right', 'num_format': '#,##0'})
    _p2s_kev = product_to_sheet or {}

    ws.merge_range(row, 0, row, 3,
                   '  Products with Known Exploited Vulnerabilities  (CISA KEV, unresolved only)',
                   sect_fmt)
    row += 1

    if kev_col and 'Base Product' in triage_dedup.columns:
        _kev_unr_df = triage_dedup[kev_mask & _is_unr].copy()
        _has_dt_kev = 'Device Type' in _kev_unr_df.columns
        _kev_srv_mask = (_kev_unr_df['Device Type'].astype(str).str.lower().str.contains('server', na=False)
                          if _has_dt_kev else pd.Series([False] * len(_kev_unr_df), index=_kev_unr_df.index))

        # -- Workstations: summarised by product --
        ws.write(row, 0, 'Workstation Products', hdr_fmt)
        ws.merge_range(row, 1, row, 3, '', hdr_fmt)
        row += 1
        ws.write(row, 0, 'Product',              hdr_fmt)
        ws.write(row, 1, 'Devices affected',      hdr_fmt)
        ws.write(row, 2, 'Unresolved KEV CVEs',   hdr_fmt)
        row += 1
        _wks_kev_df = _kev_unr_df[~_kev_srv_mask]
        if not _wks_kev_df.empty:
            _wks_grp = (_wks_kev_df.groupby('Base Product')
                        .agg(devices=('Name', 'nunique'),
                             cves=('Vulnerability Name', 'nunique'))
                        .sort_values('devices', ascending=False))
            for prod, pr in _wks_grp.iterrows():
                sheet = _p2s_kev.get(prod)
                if sheet:
                    ws.write_url(row, 0, f"internal:'{sheet}'!A1", _kev_link_fmt, string=str(prod))
                else:
                    ws.write(row, 0, str(prod), _kev_td)
                ws.write(row, 1, int(pr['devices']), _kev_td_r)
                ws.write(row, 2, int(pr['cves']),    _kev_td_r)
                row += 1
        else:
            ws.merge_range(row, 0, row, 2, 'No Work Stations found with KEV', note_fmt)
            row += 1
        row += 1

        # -- Servers: every product with an UNRESOLVED CVE, regardless of KEV --
        # Title says "regardless of KEV" — this tracks any unresolved CVE on
        # a server (not just ones on the CISA KEV catalog), grouped by
        # product per device (not one row per CVE — a product with 23
        # unresolved CVEs gets one row, not 23). 💣 Has Exploit marks a
        # product/device pair where at least one of those unresolved CVEs
        # is on the CISA KEV catalog specifically.
        ws.write(row, 0, 'All Servers  (tracked regardless of KEV)', hdr_fmt)
        ws.merge_range(row, 1, row, 6, '', hdr_fmt)
        row += 1
        ws.write(row, 0, 'Device',                    hdr_fmt)
        ws.write(row, 1, 'Device Type',                hdr_fmt)
        ws.write(row, 2, 'Product',                    hdr_fmt)
        ws.write(row, 3, 'Unresolved CVEs',            hdr_fmt)
        ws.write(row, 4, '\U0001f4a3 Has Exploit',      hdr_fmt)
        ws.write(row, 5, 'Last Response',              hdr_fmt)
        ws.write(row, 6, 'Days Since Last Response',   hdr_fmt)
        row += 1

        _has_dt_all = 'Device Type' in triage_dedup.columns
        _all_srv_mask = (triage_dedup['Device Type'].astype(str).str.lower().str.contains('server', na=False)
                          if _has_dt_all else pd.Series([False] * len(triage_dedup), index=triage_dedup.index))
        _srv_unr_df = triage_dedup[_all_srv_mask & _is_unr]

        if (not _srv_unr_df.empty and 'Name' in _srv_unr_df.columns
                and 'Vulnerability Name' in _srv_unr_df.columns and 'Base Product' in _srv_unr_df.columns):
            _has_lr_all   = 'Last Response' in _srv_unr_df.columns
            _has_days_all = 'Days Since Last Response' in _srv_unr_df.columns
            _srv_kev_row_mask = kev_mask.reindex(_srv_unr_df.index).fillna(False)

            _srv_prod_grp = (_srv_unr_df.assign(_is_kev=_srv_kev_row_mask)
                             .groupby(['Name', 'Base Product'])
                             .agg(device_type=('Device Type', 'first') if _has_dt_all else ('Name', lambda s: 'Unknown'),
                                  cve_count=('Vulnerability Name', 'nunique'),
                                  has_kev=('_is_kev', 'any'),
                                  last_response=('Last Response', 'first') if _has_lr_all else ('Name', lambda s: ''),
                                  days_since=('Days Since Last Response', lambda s: pd.to_numeric(s, errors='coerce').max())
                                             if _has_days_all else ('Name', lambda s: ''))
                             .reset_index()
                             .sort_values(by=['has_kev', 'Name', 'Base Product'], ascending=[False, True, True]))

            for _, _r in _srv_prod_grp.iterrows():
                _is_kev  = bool(_r['has_kev'])
                _bf, _nf = (_kev_td_red, _kev_td_red_r) if _is_kev else (_kev_td, _kev_td_r)
                _prod    = _r['Base Product']
                _sheet   = _p2s_kev.get(_prod) if _prod else None

                ws.write(row, 0, str(_r['Name']), _bf)
                ws.write(row, 1, str(_r['device_type']) if _has_dt_all else '', _bf)
                if _sheet:
                    ws.write_url(row, 2, f"internal:'{_sheet}'!A1", _kev_link_fmt, string=str(_prod))
                else:
                    ws.write(row, 2, str(_prod), _bf)
                ws.write(row, 3, int(_r['cve_count']), _nf)
                ws.write(row, 4, 'Yes' if _is_kev else 'No', _bf)
                ws.write(row, 5, str(_r['last_response']) if _has_lr_all else '', _bf)
                _days_v = _r['days_since'] if _has_days_all else None
                ws.write(row, 6, int(_days_v) if _has_days_all and pd.notna(_days_v) else '', _nf)
                row += 1
        else:
            ws.merge_range(row, 0, row, 6,
                           'All Servers Patched \u2014 no unresolved CVEs of any kind on active servers',
                           note_fmt)
            row += 1
    else:
        _kev_unr_df = pd.DataFrame()
        ws.merge_range(row, 0, row, 3, 'CISA KEV column not available in source data.', note_fmt)
        row += 1

    ws.merge_range(row, 0, row, 6,
                   '\u2139  Workstation product links go to the triage sheet for that product. '
                   'Server rows group unresolved CVEs by product per device (one row per product, not '
                   'per CVE) and link to that product\u2019s triage sheet \u2014 \U0001f4a3 Has Exploit = Yes '
                   'means at least one of those unresolved CVEs is on the CISA KEV catalog. '
                   'All counts fixed at report generation, active devices only.',
                   note_fmt)
    ws.set_row(row, 28)
    row += 2

    # ── All Devices with Unpatched Known Exploited Vulnerabilities ───────────────
    # Unlike Top At-Risk Devices below (capped to 10, mixed priority), this is
    # the complete list — every active device with at least one unresolved CVE
    # that is on the CISA KEV catalog, regardless of count.
    #
    # Last Response / Days Since Last Response are included specifically so a
    # device that still shows "unresolved" here can be checked against how
    # recently it actually checked in — a device that's gone quiet may simply
    # not have reported a fresh scan showing the CVE patched, rather than
    # genuinely still being vulnerable.
    _kev_has_lr   = 'Last Response' in triage_dedup.columns
    _kev_has_days = 'Days Since Last Response' in triage_dedup.columns

    ws.merge_range(row, 0, row, 5,
                   '  Devices with Unpatched Known Exploited Vulnerabilities  (CISA KEV)',
                   sect_fmt)
    row += 1
    ws.write(row, 0, 'Device',                       hdr_fmt)
    ws.write(row, 1, 'Device Type',                  hdr_fmt)
    ws.write(row, 2, 'Product(s)',                   hdr_fmt)
    ws.write(row, 3, 'Unresolved KEV CVEs',           hdr_fmt)
    ws.write(row, 4, 'Last Response',                hdr_fmt)
    ws.write(row, 5, 'Days Since Last Response',     hdr_fmt)
    row += 1
    if kev_col and not _kev_unr_df.empty and 'Name' in _kev_unr_df.columns:
        _dt_agg = ('Device Type', 'first') if 'Device Type' in _kev_unr_df.columns else ('Name', lambda s: 'Unknown')
        _dev_grp = (_kev_unr_df.groupby('Name')
                    .agg(device_type=_dt_agg,
                         products=('Base Product', lambda s: ', '.join(sorted(s.astype(str).unique())))
                                  if 'Base Product' in _kev_unr_df.columns else ('Name', lambda s: ''),
                         cves=('Vulnerability Name', 'nunique'),
                         last_response=('Last Response', 'first') if _kev_has_lr else ('Name', lambda s: ''),
                         days_since=('Days Since Last Response', lambda s:
                             pd.to_numeric(s, errors='coerce').max())
                             if _kev_has_days else ('Name', lambda s: ''))
                    .sort_values('cves', ascending=False))
        for dev, dr in _dev_grp.iterrows():
            _bf, _nf = _kev_td, _kev_td_r
            _name_label = str(dev)
            _days_val = (
                int(dr['days_since']) if _kev_has_days
                and not (isinstance(dr['days_since'], float) and pd.isna(dr['days_since']))
                else ''
            )
            ws.write(row, 0, _name_label,                          _bf)
            ws.write(row, 1, str(dr['device_type']),               _bf)
            ws.write(row, 2, str(dr['products']),                  _bf)
            ws.write(row, 3, int(dr['cves']),                      _nf)
            ws.write(row, 4, str(dr['last_response']) if _kev_has_lr else '', _bf)
            ws.write(row, 5, _days_val,                            _nf)
            row += 1
    else:
        ws.merge_range(row, 0, row, 5, 'No devices with unpatched KEV CVEs.', note_fmt)
        row += 1
    ws.merge_range(row, 0, row, 5,
                   '\u2139  Full list \u2014 not capped, unlike Top At-Risk Devices further below. '
                   'Includes every active device with at least one unresolved CISA KEV CVE, '
                   'fixed at report generation.  '
                   'A device that has not checked in recently may still show as unresolved simply '
                   'because no newer scan has confirmed the patch \u2014 use Last Response / Days Since '
                   'Last Response to tell that apart from a genuinely still-vulnerable device.',
                   note_fmt)
    ws.set_row(row, 28)
    row += 2

    # Build live formula strings — these make both tables live when ☐/☑ are toggled.
    # Each product sheet carries hidden per-sheet subtotal cells (col R rows 1-2,
    # written by product_sheets via sheet_helpers.write_hs_subtotals), so the
    # fleet totals here are one short cell reference per sheet rather than a
    # COUNTIF per sheet.  The sums are written ONCE into two hidden helper
    # cells next to the table (col G) and every visible formula references
    # those — so visible formulas stay constant-size no matter how many
    # product sheets exist (Excel's stored-formula limit is 8,192 chars; the
    # old COUNTIF-chain approach had NO length guard here and could silently
    # ship a corrupt workbook on big fleets).
    _p2s = product_to_sheet or {}
    if _p2s:
        _f_sum_rs_res   = ' + '.join(_hs_ref(sn, 'res')   for sn in _p2s.values())
        _f_sum_rs_unres = ' + '.join(_hs_ref(sn, 'unres') for sn in _p2s.values())
        _live = max(len(_f_sum_rs_res), len(_f_sum_rs_unres)) + 1 <= 8192
        if not _live:
            log.warning(
                "Resolution Status live formulas disabled: helper sum is %d characters, "
                "above Excel's 8,192 character limit. Static values will be written.",
                max(len(_f_sum_rs_res), len(_f_sum_rs_unres)) + 1,
            )
    else:
        _live = False
    if _live:
        # Helper cells are written just before the table rows below (the row
        # position isn't known yet); the visible formulas use these refs.
        _rs_helper_col = 6                     # col G — table itself uses A-D
        _f_res   = None                        # set once helper row is known
        _f_unres = None
        _f_total = None
    else:
        _f_res   = None   # guard: prevents UnboundLocalError when no product sheets exist
        _f_unres = None
        _f_total = None

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
    # "Resolved" combines two sources so this table doesn't sit at 0% forever
    # when no --patch report is supplied:
    #   1. CONFIRMED this period — explicit patch evidence, or Status/Threat
    #      Status == 'RESOLVED' in the raw export (the ☑/☐ checkbox system).
    #   2. INFERRED since the previous report — a (device, CVE) pair that was
    #      UNRESOLVED last report and is simply no longer present this period.
    #      N-able's detection export never emits a RESOLVED row for a patched
    #      CVE — a fixed CVE just stops appearing — so #2 is often the ONLY
    #      way resolution ever shows up without a patch report. It's inferred,
    #      not proven: disappearance can also mean a device was decommissioned,
    #      renamed, or dropped out of RMM. (The 'Resolved Since Previous
    #      Report' detail sheet was removed in v0.33 — inferred resolution
    #      double-reported work already tracked by the Patch Confirmed
    #      sheets; this count remains as a metric only.)
    # Total is therefore Resolved + Unresolved, where Unresolved = still-active
    # rows this period and Resolved = confirmed + inferred — i.e. "of everything
    # tracked as of the last report, how much is now closed out."
    row += 1

    # ── Active Devices Failing Patch Status Check ────────────────────────────
    # Sourced from an imported RMM monitoring-check export (Advanced Options
    # → Patch Status Check Report). These are devices where RMM's own
    # automated check can't confirm patch status at all — distinct from a
    # CVE being genuinely unresolved. A device here may show unresolved CVEs
    # purely because no fresh scan has ever confirmed a patch, not because
    # it's still genuinely vulnerable.
    if patch_check_active_df is not None and not patch_check_active_df.empty:
        _pchk_td   = workbook.add_format({'border': 1})
        _pchk_td_r = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0'})
        _pchk_td_red = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6'})
        _pchk_td_red_r = workbook.add_format({'border': 1, 'bg_color': '#FCE4D6',
                                               'align': 'right', 'num_format': '#,##0'})
        _pchk_date_fmt = workbook.add_format({'border': 1, 'num_format': 'yyyy-mm-dd hh:mm'})
        _pchk_date_fmt_red = workbook.add_format({'border': 1, 'num_format': 'yyyy-mm-dd hh:mm',
                                                   'bg_color': '#FCE4D6'})

        ws.merge_range(row, 0, row, 5,
                       f'  \u26a0  Active Devices Failing Patch Status Check  '
                       f'({len(patch_check_active_df)} device(s))',
                       sect_fmt)
        row += 1
        ws.write(row, 0, 'Device',            hdr_fmt)
        ws.write(row, 1, 'Device Type',       hdr_fmt)
        ws.write(row, 2, 'Customer',          hdr_fmt)
        ws.write(row, 3, 'Site',              hdr_fmt)
        ws.write(row, 4, 'Last Failure',      hdr_fmt)
        ws.write(row, 5, 'Days Since',        hdr_fmt)
        row += 1
        for _, _r in patch_check_active_df.iterrows():
            _days = _r.get('Days Since')
            _long_silent = pd.notna(_days) and float(_days) >= 30
            _bf   = _pchk_td_red   if _long_silent else _pchk_td
            _nf   = _pchk_td_red_r if _long_silent else _pchk_td_r
            _df_  = _pchk_date_fmt_red if _long_silent else _pchk_date_fmt
            ws.write(row, 0, str(_r.get('Device', '')),      _bf)
            ws.write(row, 1, str(_r.get('Device Type', '')), _bf)
            ws.write(row, 2, str(_r.get('Customer', '')),    _bf)
            ws.write(row, 3, str(_r.get('Site', '')),        _bf)
            _last_fail = _r.get('Last Failure')
            if pd.notna(_last_fail):
                ws.write_datetime(row, 4, _last_fail, _df_)
            else:
                ws.write(row, 4, '', _bf)
            ws.write(row, 5, int(_days) if pd.notna(_days) else '', _nf)
            row += 1
        ws.merge_range(row, 0, row, 5,
                       '\u2139  \U0001f7e7 Highlighted = failing \u2265 30 days. Imported via Advanced Options → '
                       'Patch Status Check Report. These devices are also flagged with a \U0001f527 marker '
                       'in Top At-Risk Devices below.',
                       note_fmt)
        ws.set_row(row, 28)
        row += 2


    ws.write(row, 0, 'Status',          hdr_fmt)
    ws.write(row, 1, 'Detection Rows',  hdr_fmt)
    ws.write(row, 2, '% of Total',      hdr_fmt)
    ws.write(row, 3, 'Unique CVE Types (at generation)', hdr_fmt)
    row += 1

    _rr_confirmed = int(_is_res.sum()); _ur = int(_is_unr.sum())
    _rc_confirmed = int(triage_dedup.loc[_is_res, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0
    _uc  = int(triage_dedup.loc[_is_unr, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in triage_dedup.columns else 0

    _trend_resolved_pairs = 0
    _trend_resolved_cves  = 0
    if trend_data:
        _tm = trend_data.get('metrics', {})
        _trend_resolved_pairs = int(_tm.get('resolved_pair_count', 0))
        _trend_resolved_cves  = int(_tm.get('resolved_cve_count', 0))

    _rr  = _rr_confirmed + _trend_resolved_pairs
    _rc  = _rc_confirmed + _trend_resolved_cves   # may double-count a CVE id seen in both sources; acceptable for a summary count
    _tot = _rr + _ur

    res_row = row
    if _live:
        # Hidden helper cells (col G, alongside the Resolved/Unresolved rows):
        # G holds the two long-ish per-sheet sums once; visible cells reference
        # them.  White-on-white so they don't visually clutter the Summary.
        _rs_helper_fmt = workbook.add_format({'num_format': '#,##0',
                                              'font_color': 'white', 'bg_color': 'white'})
        ws.write_formula(res_row,     _rs_helper_col, f'={_f_sum_rs_res}',
                         _rs_helper_fmt, _rr_confirmed)
        ws.write_formula(res_row + 1, _rs_helper_col, f'={_f_sum_rs_unres}',
                         _rs_helper_fmt, _ur)
        _f_res   = f'$G${res_row + 1}'          # Excel 1-indexed rows
        _f_unres = f'$G${res_row + 2}'
        _f_total = f'({_f_res}+{_f_unres})'
    ws.write(row, 0, 'Resolved',   lbl_fmt)
    if _live:
        _f_res_combined = f'({_f_res}) + {_trend_resolved_pairs}' if _trend_resolved_pairs else _f_res
        ws.write_formula(row, 1, f'={_f_res_combined}', grn_fmt, _rr)
    else:
        ws.write(row, 1, _rr, grn_fmt)
    # % = resolved / total — live
    if _live:
        _f_total_combined = f'({_f_total}) + {_trend_resolved_pairs}' if _trend_resolved_pairs else _f_total
        ws.write_formula(row, 2, f'=IF(({_f_total_combined})>0,({_f_res_combined})/({_f_total_combined}),0)', _live_pct, _rr/_tot if _tot else 0)
    else:
        ws.write(row, 2, _rr/_tot if _tot else 0, val_pct)
    ws.write(row, 3, _rc, grn_fmt)
    row += 1

    ws.write(row, 0, 'Unresolved', lbl_fmt)
    _write_val(row, 1, _f_unres, _ur,  red_fmt)
    if _live:
        ws.write_formula(row, 2, f'=IF(({_f_total_combined})>0,({_f_unres})/({_f_total_combined}),0)', _live_pct, _ur/_tot if _tot else 0)
    else:
        ws.write(row, 2, _ur/_tot if _tot else 0, val_pct)
    ws.write(row, 3, _uc, red_fmt)
    row += 1

    ws.write(row, 0, 'Total', lbl_fmt)
    if _live:
        ws.write_formula(row, 1, f'={_f_total_combined}', val_fmt, _tot)
    else:
        ws.write(row, 1, _tot, val_fmt)
    ws.write(row, 2, 1.0, val_pct)
    ws.write(row, 3, _rc + _uc, val_fmt)
    row += 1

    _trend_note = (
        ' \u2018Resolved\u2019 = confirmed this period (patch evidence / Status=RESOLVED) '
        '+ INFERRED (unresolved last report, absent this period \u2014 verify against a patch '
        'report if uncertain).'
        if trend_data else
        ' No previous report was supplied, so \u2018Resolved\u2019 reflects confirmed evidence only '
        '(patch data / Status=RESOLVED). Pass --previous to also infer resolutions from CVEs that '
        'dropped out of the detection export since last time.'
    )
    ws.merge_range(row, 0, row, 3,
                   f'\u2139  {_live_note}  '
                   f'Unique CVE Type counts are fixed at report generation — they do not update live.'
                   f'{_trend_note}',
                   note_fmt)
    ws.set_row(row, 60); row += 2

    # ── N-Day Exposure Age (advanced preview) ────────────────────────────────────
    # How long unresolved detections have been sitting. Age source, in order:
    # 'N Days Exposed' (patch-match runs), else 'First detected' /
    # 'Date Published' parsed against the current date.
    if advanced_summary:
        from formatting import get_band_formats
        _age_days = None
        _age_src  = None
        if 'N Days Exposed' in triage_dedup.columns:
            _age_days = pd.to_numeric(triage_dedup['N Days Exposed'], errors='coerce')
            _age_src  = 'N Days Exposed'
        else:
            for _dc in ('First detected', 'Date Published'):
                if _dc in triage_dedup.columns:
                    _dts = pd.to_datetime(triage_dedup[_dc], errors='coerce', dayfirst=True)
                    if _dts.notna().any():
                        _age_days = (pd.Timestamp.now() - _dts).dt.days
                        _age_src  = _dc
                        break

        ws.merge_range(row, 0, row, 3, '  N-Day Exposure Age  \u26a0 advanced preview', sect_fmt)
        row += 1
        if _age_days is None:
            ws.merge_range(row, 0, row, 3,
                           '\u2139  No age source in this export — needs a '
                           "'First detected' column or a patch-match run "
                           "(which adds 'N Days Exposed').", note_fmt)
            row += 2
        else:
            _bands = get_band_formats(workbook)
            _age_unres = ~_compute_resolved_series(triage_dedup)
            _age_kev   = (triage_dedup['CISA KEV'].astype(str).str.strip().str.lower()
                          .isin(['yes', 'y', 'true', '1'])
                          if 'CISA KEV' in triage_dedup.columns
                          else pd.Series(False, index=triage_dedup.index))
            ws.write(row, 0, 'Age band',               _bands['header'])
            ws.write(row, 1, 'Unresolved detections',  _bands['header'])
            ws.write(row, 2, 'Unique CVEs',            _bands['header'])
            ws.write(row, 3, 'Unresolved KEV CVEs',    _bands['header'])
            row += 1
            for _lbl, _fmt, _lo, _hi in [
                ('90+ days',   _bands['critical'], 90,   None),
                ('60\u201389 days', _bands['high'],     60,   90),
                ('30\u201359 days', _bands['amber'],    30,   60),
                ('Under 30 days',    _bands['ok'],        None, 30),
            ]:
                _m = _age_unres & _age_days.notna()
                if _lo is not None: _m &= _age_days >= _lo
                if _hi is not None: _m &= _age_days < _hi
                _sub = triage_dedup[_m]
                ws.write(row, 0, _lbl, _fmt)
                ws.write(row, 1, int(_m.sum()), val_fmt)
                ws.write(row, 2, int(_sub['Vulnerability Name'].nunique()) if not _sub.empty else 0, val_fmt)
                ws.write(row, 3, int(_sub.loc[_age_kev[_m.index][_m], 'Vulnerability Name'].nunique())
                                 if not _sub.empty else 0, val_fmt)
                row += 1
            _unknown = int((_age_unres & _age_days.isna()).sum())
            if _unknown:
                ws.write(row, 0, 'Unknown age', _bands['label'])
                ws.write(row, 1, _unknown, val_fmt)
                row += 1
            ws.merge_range(row, 0, row, 3,
                           f'\u2139  Age source: {_age_src}. Unresolved rows only, '
                           'active scope, deduplicated per product. Fixed at generation.',
                           note_fmt)
            row += 2

    # ── Top Patch-Gap Root Causes (advanced preview) ─────────────────────────────
    # Why patches aren't landing — same classification the diagnostics
    # sheets use, summarised for the client.
    if advanced_summary and root_cause_counts:
        _RC_LABELS = {
            'coverage_gap':       'Coverage gap — device not in patch report',
            'unmanaged_app':      'Unmanaged app — product not tracked in patch report',
            'detection_mismatch': 'Detection mismatch — CVE detected but no matching patch found',
            'patch_installing':   'Patch installing — awaiting next RMM sync',
        }
        ws.merge_range(row, 0, row, 1, '  Top Patch-Gap Root Causes  \u26a0 advanced preview', sect_fmt)
        row += 1
        ws.write(row, 0, 'Root cause',       hdr_fmt)
        ws.write(row, 1, 'Device-CVE pairs', hdr_fmt)
        row += 1
        for _cause, _cnt in sorted(root_cause_counts.items(), key=lambda x: -x[1]):
            ws.write(row, 0, _RC_LABELS.get(_cause, str(_cause)), lbl_fmt)
            ws.write(row, 1, int(_cnt), val_fmt)
            row += 1
        ws.merge_range(row, 0, row, 3,
                       '\u2139  From patch-evidence matching — see the diagnostics '
                       'sheets for per-device detail.', note_fmt)
        row += 2

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
                   f'"Stale devices (excluded)" = last seen before {_cutoff_lbl} (set in the app), moved to Stale Excluded Devices sheet.  '
                   'All counts fixed at report generation.',
                   note_fmt)
    ws.set_row(row, 42); row += 2

    # Data Filtering Reconciliation waterfall
    # Uses deduped counts (_all_df, _stale_dedup) so [+] - [-] = [=] exactly.
    _stale_cves_dedup = int(_stale_dedup['Vulnerability Name'].nunique()) if not _stale_dedup.empty and 'Vulnerability Name' in _stale_dedup.columns else 0

    row += 1
    ws.merge_range(row, 0, row, 3, f'  Data Filtering Reconciliation  (CVSS \u2265 {threshold})', sect_fmt); row += 1
    ws.write(row, 0, 'Filter Step',      hdr_fmt); ws.write(row, 1, 'Unique Devices', hdr_fmt)
    ws.write(row, 2, 'Detection Rows',   hdr_fmt); ws.write(row, 3, 'Unique CVE Types', hdr_fmt); row += 1
    ws.write(row, 0, '[+]  Total detections (all devices, CVSS \u2265 threshold, deduplicated per product)', wf_plus)
    ws.write(row, 1, _all_devs, val_fmt); ws.write(row, 2, _all_rows, val_fmt); ws.write(row, 3, _all_cves, val_fmt); row += 1
    if not _stale_dedup.empty:
        _stale_dedup_devs = int(_stale_dedup['Name'].nunique()) if 'Name' in _stale_dedup.columns else _stale_devs
        ws.write(row, 0, f'[-]  Excluded: stale devices  (last seen before {_cutoff_lbl})', wf_minus)
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

    if not triage_dedup.empty and 'Name' in triage_dedup.columns:
        # Was: computed against raw triage_df (every detection row, including
        # duplicates). build_product_sheets deduplicates by (Name,
        # Vulnerability Name) PER PRODUCT *before* determining ☑/☐ — so if a
        # device has multiple raw rows for the same CVE with mixed resolved
        # status (e.g. detected under two Affected Products variants, or a
        # stale scan entry alongside a fresh one), the written sheet shows
        # one final verdict for that CVE, while counting against raw
        # triage_df would count it as unresolved if ANY of the duplicate
        # rows was. That let this table disagree with the actual product
        # sheets by a handful of CVEs per affected device. Using
        # triage_dedup — the same per-Base-Product-deduplicated frame the
        # Resolution Status table and Device Breakdown already use, built
        # with the identical drop_duplicates(['Name','Vulnerability Name'])
        # rule build_product_sheets applies — guarantees this table can
        # never again show a different "unresolved" verdict for a CVE than
        # the product sheet the reader would go check it against.
        _unr_df = triage_dedup[~_compute_resolved_series(triage_dedup)].copy()
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

            _check_active = patch_check_active_names or set()
            _td_chkfail    = workbook.add_format({'border': 1, 'bg_color': '#E4DFEC', 'font_color': '#4C3B6E'})
            _td_chkfail_r  = workbook.add_format({'border': 1, 'bg_color': '#E4DFEC', 'font_color': '#4C3B6E',
                                                   'align': 'right', 'num_format': '#,##0'})
            if _check_active:
                from data_pipeline import normalize_device_name as _norm_dev_name

            for _r in _top:
                _srv        = 'server' in str(_r.device_type).lower()
                _exp        = str(_r.has_exploit).strip().lower() == 'yes'
                _chk_fail   = bool(_check_active) and _norm_dev_name(_r.Name) in _check_active
                if _exp:
                    _bf, _nf = _td_exp, _td_exp_r
                elif _chk_fail:
                    _bf, _nf = _td_chkfail, _td_chkfail_r
                elif _srv:
                    _bf, _nf = _td_srv, _td_srv_r
                else:
                    _bf, _nf = _td, _td_r
                _prefix     = '\U0001f527 ' if _chk_fail else ''
                _name_label = f'{_prefix}{_r.Name}'
                _days_val   = int(_r.days_since) if hasattr(_r, 'days_since') and not (isinstance(_r.days_since, float) and pd.isna(_r.days_since)) else ''
                ws.write(row, 0, _name_label,               _bf)
                ws.write(row, 1, str(_r.username),           _bf)
                ws.write(row, 2, int(_r.cve_count),          _nf)
                ws.write(row, 3, str(_r.has_exploit),        _bf)
                ws.write(row, 4, str(_r.device_type),        _bf)
                ws.write(row, 5, str(_r.last_response) if hasattr(_r, 'last_response') else '', _bf)
                ws.write(row, 6, _days_val,                  _nf)
                row += 1

            _chkfail_note = (
                '  \U0001f7ea Purple = active device failing its RMM Patch Status Check '
                '(\U0001f527 prefix on name) — see Patch Check Failures sheet.  '
                if _check_active else ''
            )
            ws.merge_range(row, 0, row, 6,
                f'ℹ  🟡 Amber = Server.  🟥 Red = known exploit.  '
                f'{_chkfail_note}'
                f'Up to 10 devices. Unresolved CVE counts only.',
                note_fmt)
            ws.set_row(row, 30); row += 1
        else:
            ws.merge_range(row, 0, row, 6, 'No unresolved CVE data.', note_fmt); row += 1
    else:
        ws.merge_range(row, 0, row, 6, 'No active device data.', note_fmt); row += 1

    ws.set_column('A:A', 32); ws.set_column('B:B', 22)
    ws.set_column('C:C', 18); ws.set_column('D:D', 14); ws.set_column('E:E', 22)
    ws.set_column('F:F', 24); ws.set_column('G:G', 20)

    # Month-over-Month
    if trend_data:
        m=trend_data['metrics']
        row+=1
        ws.merge_range(row,0,row,3,'  Month-over-Month Patching Progress',sect_fmt); row+=1
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
            row+=1

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