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
                               approaching_stale_names: Optional[Set[str]] = None,
                               stale_warning_days: int = 14,
                               product_to_sheet: Optional[dict] = None,
                               include_health_score: bool = False,
                               patch_resolved_pairs: Optional[set] = None,
                               health_triage_df: 'Optional[pd.DataFrame]' = None,
                               health_score_threshold: float = 7.0,
                               has_patch_report: bool = False,
                               prev_report_name: str = '',
                               patch_check_active_df: 'Optional[pd.DataFrame]' = None,
                               patch_check_active_names: Optional[Set[str]] = None):
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

    _approaching_set = approaching_stale_names or set()

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
    from resolution import (get_sheet_product_key as _get_sheet_pk_sum,
                            split_patch_pairs as _split_patch_pairs_sum,
                            compute_resolved_flags as _compute_flags_sum,
                            compute_resolved_series as _compute_resolved_series_shared)
    _p2s_sum   = product_to_sheet or {}
    _patch_2d_sum, _patch_3d_sum = _split_patch_pairs_sum(patch_resolved_pairs)

    def _compute_resolved_series(df: 'pd.DataFrame') -> 'pd.Series':
        """Return a bool Series aligned to df's index — True = ☑ resolved.
        Delegates to resolution.compute_resolved_series(), the single
        index-safe implementation shared with product_sheets.py — see that
        function's docstring for the row-misalignment bug this replaced."""
        return _compute_resolved_series_shared(df, _p2s_sum, patch_resolved_pairs,
                                               approaching_stale_names=_approaching_set)

    _is_res = _compute_resolved_series(triage_dedup)
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

    # CISA KEV — Known Exploited Vulnerabilities catalog. Distinct from the
    # generic 'Has Known Exploit' column above: KEV is CISA's authoritative
    # "actively exploited in the wild" list, so it gets its own Key Metrics
    # rows and its own product/device tracking tables (see below).
    kev_col        = 'CISA KEV' if 'CISA KEV' in triage_dedup.columns else None
    kev_mask       = triage_dedup[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y']) if kev_col else pd.Series([False]*len(triage_dedup), index=triage_dedup.index)
    kev_rows       = int(kev_mask.sum())
    kev_cves       = int(triage_dedup.loc[kev_mask, 'Vulnerability Name'].nunique()) if kev_col and 'Vulnerability Name' in triage_dedup.columns else 0
    kev_patched    = int((kev_mask & _is_res).sum())
    kev_unpatch    = int((kev_mask & _is_unr).sum())

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

    # Stale + NIRM CISA KEV counts, mirroring the CVSS 9+ pattern above so
    # KEV gets the same All / Active / Excluded reconciliation.
    _stale_kev = 0
    _stale_kev_cves = 0
    if stale_excluded_df is not None and not stale_excluded_df.empty and kev_col and kev_col in stale_excluded_df.columns:
        _stale_kev_mask  = stale_excluded_df[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y'])
        _stale_kev       = int(_stale_kev_mask.sum())
        _stale_kev_cves  = int(stale_excluded_df.loc[_stale_kev_mask, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in stale_excluded_df.columns else 0
    _nirm_kev  = 0
    _nirm_kev_cves = 0
    if 'Last Response' in filtered_df.columns and kev_col and kev_col in filtered_df.columns:
        _nirm_kev_mask   = filtered_df[kev_col].astype(str).str.strip().str.lower().isin(['yes','true','1','y'])
        _nirm_kev        = int((_nirm_mask & _nirm_kev_mask).sum())
        _nirm_kev_cves   = int(filtered_df.loc[_nirm_mask & _nirm_kev_mask, 'Vulnerability Name'].nunique()) if 'Vulnerability Name' in filtered_df.columns else 0
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

    # stale crit rows from _stale_dedup directly (consistent with _all_rows)
    if not _stale_dedup.empty and score_col and score_col in _stale_dedup.columns:
        _stale_crit_dedup = int((pd.to_numeric(_stale_dedup[score_col], errors='coerce') >= 9.0).sum())
    else:
        _stale_crit_dedup = _stale_crit

    # Uses the deduped not-in-RMM row count (matches how _all_rows is now built)
    # when the actual dataframe was passed; falls back to the pre-computed
    # scalar for callers that haven't been updated to pass not_in_rmm_df yet.
    _nirm_excl_rows    = len(_nirm_dedup) if not _nirm_dedup.empty else not_in_rmm_cve_count
    _excl_rows          = len(_stale_dedup) + _nirm_excl_rows
    _excl_devs_tot      = _stale_devs + _nirm_devs
    _excl_crit_tot      = _stale_full_crit + _nirm_crit        # full stale, not _p2s_keys-filtered
    _excl_crit_cves_tot = _stale_full_crit_cves + _nirm_crit_cves
    _excl_kev_tot       = _stale_full_kev + _nirm_kev          # full stale, not _p2s_keys-filtered
    _excl_kev_cves_tot  = _stale_full_kev_cves + _nirm_kev_cves

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
        # threshold (e.g. CVSS ≥ 7.0 even when the report itself is ≥ 9.0),
        # so _score_scope can legitimately contain a Base Product with no
        # entry in product_to_sheet at all — no dedicated sheet, no ☑/☐
        # column, nothing a COUNTIF formula could reference. The Python-side
        # score (_phs, above) already accounts for these "hidden" products
        # correctly via compute_resolved_series(). But a LIVE formula only
        # ever sums COUNTIFs over _p2s_hs_vals (existing sheets) — so if we
        # built one anyway, it would silently drop those products' rows the
        # moment Excel recalculates (e.g. the reader toggles any unrelated
        # checkbox), producing a number that visibly disagrees with the
        # correct one this workbook was generated with. Rather than ship a
        # formula that can drift wrong on first interaction, fall back to
        # the same static-value path already used when a formula is too
        # long for Excel's limit.
        if _hs_live and 'Base Product' in _score_scope.columns:
            _hs_hidden_products = set(_score_scope['Base Product'].unique()) - set((product_to_sheet or {}).keys())
            if _hs_hidden_products:
                log.info(
                    "Patching Health Score live formulas disabled: %d product(s) in the "
                    "health scope have no dedicated sheet (%s) — a live COUNTIF formula "
                    "can't reference rows that don't exist on any sheet. Static score "
                    "values will be written instead.",
                    len(_hs_hidden_products), ', '.join(sorted(_hs_hidden_products)),
                )
                _hs_live = False

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

    # ── Month-over-Month Remediation Summary ────────────────────────────────────
    # Only rendered when a previous report was supplied. The point of this
    # section is to make remediation WORK visible without opening Trend
    # Summary — the Resolution Status table below only shows the current
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
        ws.write(row, 3, '\u25b2  offset' if _new_count else '\u2014', _mom_up_fmt if _new_count else _mom_same_fmt)
        ws.write(row, 4, 'Pairs that became unresolved for the first time this period.', def_fmt)
        row += 2

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

        # -- Servers: EVERY active server device, regardless of KEV status --
        # Unlike Workstation Products above (summarised by product, KEV-only),
        # servers get their own row each — servers are typically few and
        # high-value, so visibility matters even for a server with zero KEV
        # CVEs right now. Rows with an unresolved KEV CVE are still called out
        # with a highlight so they don't get lost in the full list.
        ws.write(row, 0, 'All Servers  (tracked regardless of KEV)', hdr_fmt)
        ws.merge_range(row, 1, row, 5, '', hdr_fmt)
        row += 1
        ws.write(row, 0, 'Device',                    hdr_fmt)
        ws.write(row, 1, 'Device Type',                hdr_fmt)
        ws.write(row, 2, 'Product(s)',                 hdr_fmt)
        ws.write(row, 3, 'Unresolved KEV CVEs',        hdr_fmt)
        ws.write(row, 4, 'Last Response',              hdr_fmt)
        ws.write(row, 5, 'Days Since Last Response',   hdr_fmt)
        row += 1

        _has_dt_all = 'Device Type' in triage_dedup.columns
        _all_srv_mask = (triage_dedup['Device Type'].astype(str).str.lower().str.contains('server', na=False)
                          if _has_dt_all else pd.Series([False] * len(triage_dedup), index=triage_dedup.index))
        _srv_all_df = triage_dedup[_all_srv_mask]

        if not _srv_all_df.empty and 'Name' in _srv_all_df.columns:
            _kev_cve_by_device = (
                _kev_unr_df.groupby('Name')['Vulnerability Name'].nunique().to_dict()
                if not _kev_unr_df.empty and 'Name' in _kev_unr_df.columns else {}
            )
            _has_lr_all   = 'Last Response' in _srv_all_df.columns
            _has_days_all = 'Days Since Last Response' in _srv_all_df.columns
            _srv_all_grp = (_srv_all_df.groupby('Name')
                            .agg(device_type=('Device Type', 'first') if _has_dt_all else ('Name', lambda s: 'Unknown'),
                                 products=('Base Product', lambda s: sorted(s.astype(str).unique()))
                                          if 'Base Product' in _srv_all_df.columns else ('Name', lambda s: []),
                                 last_response=('Last Response', 'first') if _has_lr_all else ('Name', lambda s: ''),
                                 days_since=('Days Since Last Response', lambda s: pd.to_numeric(s, errors='coerce').max())
                                            if _has_days_all else ('Name', lambda s: ''))
                            .sort_values('device_type'))

            # If literally no server has an unresolved KEV CVE, a full table of
            # every server showing zero is just noise — a clean "all patched"
            # message is more useful. The full per-server table (still
            # regardless of KEV) only appears once at least one server has
            # something outstanding.
            _any_srv_kev = any(
                int(_kev_cve_by_device.get(dev, 0)) > 0 for dev in _srv_all_grp.index
            )
            if not _any_srv_kev:
                ws.merge_range(row, 0, row, 5, 'All Servers Patched', note_fmt)
                row += 1
            else:
                for dev, dr in _srv_all_grp.iterrows():
                    _kev_ct   = int(_kev_cve_by_device.get(dev, 0))
                    _has_kev  = _kev_ct > 0
                    _bf, _nf  = (_kev_td_red, _kev_td_red_r) if _has_kev else (_kev_td, _kev_td_r)
                    _products = dr['products'] if isinstance(dr['products'], list) else []
                    _prod_txt = ', '.join(_products) if _products else '\u2014'
                    _primary  = _products[0] if _products else None
                    _sheet    = _p2s_kev.get(_primary) if _primary else None

                    ws.write(row, 0, str(dev), _bf)
                    ws.write(row, 1, str(dr['device_type']), _bf)
                    if _sheet:
                        ws.write_url(row, 2, f"internal:'{_sheet}'!A1", _kev_link_fmt, string=_prod_txt)
                    else:
                        ws.write(row, 2, _prod_txt, _bf)
                    ws.write(row, 3, _kev_ct, _nf)
                    ws.write(row, 4, str(dr['last_response']) if _has_lr_all else '', _bf)
                    _days_v = dr['days_since']
                    ws.write(row, 5, int(_days_v) if _has_days_all and pd.notna(_days_v) else '', _nf)
                    row += 1
        else:
            ws.merge_range(row, 0, row, 5, 'No active server devices found.', note_fmt)
            row += 1
    else:
        _kev_unr_df = pd.DataFrame()
        ws.merge_range(row, 0, row, 3, 'CISA KEV column not available in source data.', note_fmt)
        row += 1

    ws.merge_range(row, 0, row, 5,
                   '\u2139  Product name(s) are hyperlinked to their triage sheet (multi-product cells link to '
                   'the first product listed). Unresolved KEV CVE counts are unresolved (\u2610) only, active '
                   'devices, fixed at report generation. \U0001f7e5 Highlighted server rows have \u2265 1 '
                   'unresolved KEV CVE; all other server rows are shown for visibility regardless of KEV status.',
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
    # genuinely still being vulnerable. Devices within stale_warning_days of
    # going stale get the same ⚠ / orange treatment as Top At-Risk Devices.
    _kev_has_lr   = 'Last Response' in triage_dedup.columns
    _kev_has_days = 'Days Since Last Response' in triage_dedup.columns
    _kev_approach = approaching_stale_names or set()
    _kev_td_approach   = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00'})
    _kev_td_approach_r = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00',
                                               'align': 'right', 'num_format': '#,##0'})

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
            _near_stale = dev in _kev_approach
            _bf, _nf = (_kev_td_approach, _kev_td_approach_r) if _near_stale else (_kev_td, _kev_td_r)
            _name_label = f'⚠ {dev}' if _near_stale else str(dev)
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
    _kev_approach_note = (
        f'  \U0001f7e7 Orange = offline \u2265 {stale_warning_days}d (\u26a0 prefix on name) \u2014 '
        'check Last Response before assuming the CVE is still genuinely unpatched.  '
        if _kev_approach else ''
    )
    ws.merge_range(row, 0, row, 5,
                   '\u2139  Full list \u2014 not capped, unlike Top At-Risk Devices further below. '
                   'Includes every active device with at least one unresolved CISA KEV CVE, '
                   f'fixed at report generation.  {_kev_approach_note}'
                   'A device that has not checked in recently may still show as unresolved simply '
                   'because no newer scan has confirmed the patch \u2014 use Last Response / Days Since '
                   'Last Response to tell that apart from a genuinely still-vulnerable device.',
                   note_fmt)
    ws.set_row(row, 28)
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
    #      renamed, or dropped out of RMM. See the 'Resolved Since Previous Report' sheet
    #      for the underlying device/CVE rows behind this number.
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
        'report if uncertain; see \u2018Resolved Since Previous Report\u2019 for the underlying rows).'
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

            _approaching = approaching_stale_names or set()
            _check_active = patch_check_active_names or set()
            _td_approach   = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00'})
            _td_approach_r = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'font_color': '#7B3F00',
                                                   'align': 'right', 'num_format': '#,##0'})
            _td_chkfail    = workbook.add_format({'border': 1, 'bg_color': '#E4DFEC', 'font_color': '#4C3B6E'})
            _td_chkfail_r  = workbook.add_format({'border': 1, 'bg_color': '#E4DFEC', 'font_color': '#4C3B6E',
                                                   'align': 'right', 'num_format': '#,##0'})
            if _check_active:
                from data_pipeline import normalize_device_name as _norm_dev_name

            for _r in _top:
                _srv        = 'server' in str(_r.device_type).lower()
                _exp        = str(_r.has_exploit).strip().lower() == 'yes'
                _near_stale = _r.Name in _approaching
                _chk_fail   = bool(_check_active) and _norm_dev_name(_r.Name) in _check_active
                if _exp:
                    _bf, _nf = _td_exp, _td_exp_r
                elif _chk_fail:
                    _bf, _nf = _td_chkfail, _td_chkfail_r
                elif _near_stale:
                    _bf, _nf = _td_approach, _td_approach_r
                elif _srv:
                    _bf, _nf = _td_srv, _td_srv_r
                else:
                    _bf, _nf = _td, _td_r
                _prefix     = '\U0001f527 ' if _chk_fail else ('⚠ ' if _near_stale else '')
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

            _approach_note = (
                f'  🟧 Orange = offline \u2265 {stale_warning_days}d (⚠ prefix on name).  '
                if _approaching else ''
            )
            _chkfail_note = (
                '  \U0001f7ea Purple = active device failing its RMM Patch Status Check '
                '(\U0001f527 prefix on name) — see Patch Check Failures sheet.  '
                if _check_active else ''
            )
            ws.merge_range(row, 0, row, 6,
                f'ℹ  🟡 Amber = Server.  🟥 Red = known exploit.  '
                f'{_approach_note}'
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