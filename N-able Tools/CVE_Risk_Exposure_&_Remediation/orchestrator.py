"""
orchestrator.py — pipeline coordinator.

Receives a DashboardRequest, runs the pipeline, writes the workbook,
and returns a DashboardResult.

No tkinter.  No filedialog.  Fully testable headless.
Business logic lives in: data_pipeline, diagnostics, snapshot, excel_builder.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, Set, Tuple

import pandas as pd

from config import FIXED_VERSION_RULES
from data_pipeline import (
    load_vulnerability_data, load_rmm_data, merge_data,
    process_patch_match, load_previous_report, compute_trends,
    normalize_device_name, extract_cve_id, clean_sheet_name,
)
from diagnostics import compute_patch_diagnostics, classify_root_cause
import snapshot as snap_store
from excel_builder import (
    build_trend_summary_sheet, build_trend_detail_sheets,
    build_overview_sheet, build_all_detections_sheet,
    build_product_sheets, build_stale_excluded_sheet,
    build_raw_data_sheet, build_patch_sheets, build_diagnostics_sheets,
)

log = logging.getLogger(__name__)


# ==============================================================================
# REQUEST / RESULT TYPES
# ==============================================================================

@dataclass
class DashboardRequest:
    """All inputs needed to produce one dashboard workbook."""
    vuln_path:        str
    output_path:      str
    rmm_path:         Optional[str]  = None
    skip_rmm:         bool           = False
    patch_path:       Optional[str]  = None
    include_patch:    bool           = False
    prev_report_path: Optional[str]  = None
    include_trend:    bool           = False
    threshold:        float          = 9.0
    cutoff_date:      Optional[str]  = None   # ISO date string; None = show all
    show_all_dates:   bool           = False


@dataclass
class DashboardResult:
    """Outcome returned to the caller (GUI or CLI)."""
    success:          bool
    output_path:      str             = ''
    message:          str             = ''
    trend_summary:    Optional[dict]  = None  # subset of trend metrics for display
    warnings:         list            = field(default_factory=list)


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

def run(request: DashboardRequest) -> DashboardResult:
    """
    Execute the full dashboard pipeline for one request.

    Always returns a DashboardResult — never raises to the caller.
    Errors are captured into result.success=False and result.message.
    Warnings are non-fatal issues logged and collected in result.warnings.
    """
    warnings: list[str] = []

    try:
        log.info("Dashboard run started — output: %s", request.output_path)

        # ── Load & merge ──────────────────────────────────────────────────────
        log.info("Loading vulnerability data: %s", request.vuln_path)
        df_vuln = load_vulnerability_data(request.vuln_path)
        log.info("  %d rows loaded", len(df_vuln))

        df_rmm = None
        if not request.skip_rmm and request.rmm_path:
            log.info("Loading RMM data: %s", request.rmm_path)
            df_rmm = load_rmm_data(request.rmm_path)
            log.info("  %d devices loaded", len(df_rmm))

        merged_df = merge_data(df_vuln, df_rmm, request.skip_rmm)
        log.info("Merged dataset: %d rows", len(merged_df))

        raw_df         = merged_df.copy()
        stale_excluded = pd.DataFrame()

        if not request.show_all_dates and request.cutoff_date:
            cutoff = pd.to_datetime(request.cutoff_date)
            high   = merged_df[merged_df['Vulnerability Score'] >= request.threshold]
            stale_excluded = high[
                (high['_Sort_Time'] < cutoff) &
                (high['Last Response'] != 'Not Found in RMM')
            ].copy()
            merged_df = merged_df[
                (merged_df['_Sort_Time'] >= cutoff) |
                (merged_df['Last Response'] == 'Not Found in RMM')
            ]
            log.info(
                "Date filter applied (>= %s): %d rows kept, %d stale excluded",
                request.cutoff_date, len(merged_df), len(stale_excluded),
            )

        if merged_df.empty:
            msg = "No vulnerability records found after applying date filter."
            log.warning(msg)
            return DashboardResult(success=False, message=msg)

        filtered_df = merged_df[merged_df['Vulnerability Score'] >= request.threshold].copy()
        triage_df   = filtered_df[filtered_df['Last Response'] != 'Not Found in RMM'].copy()

        not_in_rmm = filtered_df[filtered_df['Last Response'] == 'Not Found in RMM']['Name'].nunique()
        if not_in_rmm:
            w = f"{not_in_rmm} device(s) with score ≥ {request.threshold} not found in RMM — excluded from triage sheets"
            log.warning(w)
            warnings.append(w)

        log.info(
            "Filtered (score >= %.1f): %d rows, %d triage, %d not-in-RMM",
            request.threshold, len(filtered_df), len(triage_df), not_in_rmm,
        )

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

        # ── Optional: patch match ─────────────────────────────────────────────
        patch_data = None
        if request.include_patch and request.patch_path:
            log.info("Running patch match: %s", request.patch_path)
            p_ov, p_full, p_raw, tot_r, filt_r = process_patch_match(
                request.patch_path, merged_df.copy(), min_score=request.threshold)
            patch_data = (p_ov, p_full, p_raw, tot_r, filt_r)
            log.info("  Patch match: %d total rows, %d above threshold", tot_r, filt_r)

        # ── Optional: trend comparison ────────────────────────────────────────
        trend_data       = None
        prev_report_name = ''
        if request.include_trend and request.prev_report_path:
            log.info("Loading previous report for trend: %s", request.prev_report_path)
            prev_df          = load_previous_report(request.prev_report_path)
            prev_report_name = Path(request.prev_report_path).name
            trend_data       = compute_trends(merged_df, prev_df, request.threshold)
            m = trend_data['metrics']
            log.info(
                "Trend: %d new CVEs, %d resolved, %d persisting (common-product scope)",
                m['new_cve_count'], m['resolved_cve_count'], m['persisting_cve_count'],
            )

        # ── Customer name ─────────────────────────────────────────────────────
        customer_name = ''
        for col in ('Customer', 'Customer Name', 'Client', 'Client Name'):
            if col in merged_df.columns:
                vals = merged_df[col].dropna().astype(str).str.strip()
                vals = vals[vals.str.len() > 0]
                if not vals.empty:
                    customer_name = vals.iloc[0]
                    break

        # ── Classify patch gaps (explicit yellow-state categories) ────────────
        patch_resolved_pairs: Set[Tuple[str, str]] = set()
        patch_gap_pairs:      dict[Tuple[str, str], str] = {}
        diagnostics: dict = {'patch_lag_df': pd.DataFrame(),
                             'version_drift_df': pd.DataFrame(),
                             'root_cause_df': pd.DataFrame()}

        if patch_data:
            p_full = patch_data[1].copy()
            p_full['_nk'] = p_full['Name'].astype(str).apply(normalize_device_name)
            p_full['_ck'] = p_full['Vulnerability Name'].astype(str).apply(extract_cve_id)

            if 'Resolved (from Patch Report)' in p_full.columns:
                patch_resolved_pairs = set(zip(
                    p_full.loc[p_full['Resolved (from Patch Report)'] == 'Resolved', '_nk'],
                    p_full.loc[p_full['Resolved (from Patch Report)'] == 'Resolved', '_ck'],
                ))
                log.info("Patch-confirmed resolved pairs: %d", len(patch_resolved_pairs))

            # Root cause per row — drives both highlight colour and diagnostics sheet
            p_full['_root_cause'] = p_full.apply(classify_root_cause, axis=1)
            for _, row in p_full[p_full['_root_cause'].notna()].iterrows():
                patch_gap_pairs[(row['_nk'], row['_ck'])] = row['_root_cause']

            # Summarise gaps as warnings
            cause_counts: dict[str, int] = {}
            for c in patch_gap_pairs.values():
                cause_counts[c] = cause_counts.get(c, 0) + 1
            for cause, count in cause_counts.items():
                w = f"Patch gap [{cause}]: {count} device-CVE pair(s)"
                log.warning(w)
                warnings.append(w)

            # Full diagnostics (lag, drift, root cause table)
            product_rules = FIXED_VERSION_RULES  # reuse section of config for remediation hints
            diagnostics = compute_patch_diagnostics(patch_data[1], product_rules)

            rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
            if not rc_df.empty:
                mis = rc_df[rc_df.get('_cause_internal', rc_df.get('Patch Evidence Notes', '')) == 'version_compliant']
                if not mis.empty:
                    warnings.append(
                        f"{len(mis)} device-CVE pair(s) show 'Installed but still detected' — "
                        f"see 'Patch Evidence Notes' sheet"
                    )

        patch_confirmed_count = 0
        if patch_resolved_pairs:
            triage_keys = set(zip(
                triage_df['Name'].apply(normalize_device_name),
                triage_df['Vulnerability Name'].apply(extract_cve_id),
            ))
            patch_confirmed_count = len(patch_resolved_pairs & triage_keys)

        redetected_count = 0
        if trend_data and request.prev_report_path:
            try:
                xl_prev = pd.ExcelFile(request.prev_report_path)
                if 'Resolved' in xl_prev.sheet_names:
                    prev_res = xl_prev.parse('Resolved')
                    if 'Vulnerability Name' in prev_res.columns:
                        prev_res_cves  = set(prev_res['Vulnerability Name'].apply(extract_cve_id))
                        cur_cves_scope = set(triage_df['Vulnerability Name'].apply(extract_cve_id))
                        redetected_count = len(prev_res_cves & cur_cves_scope)
                        if redetected_count:
                            w = f"{redetected_count} CVE(s) resolved last period but re-detected this period"
                            log.warning(w)
                            warnings.append(w)
            except Exception as exc:
                log.warning("Could not compute re-detected count: %s", exc)

        # ── Write workbook ────────────────────────────────────────────────────
        log.info("Writing workbook: %s", request.output_path)
        with pd.ExcelWriter(request.output_path, engine='xlsxwriter') as writer:
            wb = writer.book
            link_fmt   = wb.add_format({'font_color': 'blue', 'underline': True})
            header_fmt = wb.add_format({'bold': True, 'font_size': 12,
                                        'bg_color': '#D9D9D9', 'border': 1})
            miss_fmt   = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            if trend_data:
                build_trend_summary_sheet(wb, trend_data, request.threshold,
                                          prev_report_name, header_fmt,
                                          customer_name=customer_name)

            build_overview_sheet(
                wb, merged_df, filtered_df, triage_df, request.threshold,
                product_to_sheet, header_fmt, link_fmt,
                customer_name=customer_name,
                patch_confirmed_count=patch_confirmed_count,
                redetected_count=redetected_count,
                sheet_name=overview_sheet_name,
                trend_metrics=trend_data['metrics'] if trend_data else None,
                health_score=diagnostics.get('health_score'),
                evidence_summary=diagnostics.get('evidence_summary'),
                recommended_actions=diagnostics.get('recommended_actions'),
            )

            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'New This Month', 'Persisting CVEs'})

            build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                                  patch_resolved_pairs=patch_resolved_pairs,
                                  patch_gap_pairs=patch_gap_pairs)

            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'Resolved'})

            build_all_detections_sheet(writer, merged_df, link_fmt, miss_fmt)

            if not stale_excluded.empty:
                build_stale_excluded_sheet(writer, stale_excluded)

            build_raw_data_sheet(writer, raw_df)

            if patch_data:
                build_patch_sheets(writer, patch_data[0], patch_data[1], patch_data[2])
                if any(not diagnostics[k].empty for k in diagnostics):
                    build_diagnostics_sheets(writer, diagnostics)

        log.info("Workbook written successfully")

        # ── Save structured snapshot ──────────────────────────────────────────
        rc_summary: dict[str, int] = {}
        rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
        if not rc_df.empty and 'Patch Evidence Notes' in rc_df.columns:
            rc_summary = rc_df['Patch Evidence Notes'].value_counts().to_dict()

        snap_store.save(
            output_path       = request.output_path,
            customer          = customer_name,
            threshold         = request.threshold,
            unique_cves       = int(filtered_df['Vulnerability Name'].nunique()),
            unique_devices    = int(filtered_df['Name'].nunique()),
            trend_metrics     = trend_data['metrics'] if trend_data else None,
            root_cause_summary= rc_summary or None,
        )

        trend_summary = None
        if trend_data:
            m = trend_data['metrics']
            trend_summary = {
                'new_cve_count':       m['new_cve_count'],
                'resolved_cve_count':  m['resolved_cve_count'],
                'persisting_cve_count':m['persisting_cve_count'],
            }

        return DashboardResult(
            success=True,
            output_path=request.output_path,
            message=f"Dashboard saved to:\n{request.output_path}",
            trend_summary=trend_summary,
            warnings=warnings,
        )

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log.error("Dashboard run failed: %s\n%s", exc, tb)
        return DashboardResult(success=False, message=str(exc))
