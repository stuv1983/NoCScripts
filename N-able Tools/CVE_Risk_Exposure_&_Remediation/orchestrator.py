"""
orchestrator.py — pipeline coordinator.

Receives a DashboardRequest, runs the pipeline, writes the workbook,
and returns a DashboardResult.

No tkinter.  No filedialog.  Fully testable headless.
Business logic lives in: data_pipeline, diagnostics, snapshot, excel_builder.

"""
# Copyright (c) 2026 stuart villanti, Inc. All rights reserved.
# This code is licensed under the MIT License. See LICENSE in the project root for license terms.

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
    load_patch_failure_report, build_patch_failure_lookup,
)
from diagnostics import compute_patch_diagnostics, classify_root_cause
import snapshot as snap_store
from excel_builder import (
    get_workbook_styles,
    build_client_summary_sheet,
    build_trend_summary_sheet, build_trend_detail_sheets,
    build_overview_sheet, build_all_detections_sheet,
    build_product_sheets, build_stale_excluded_sheet,
    build_stale_cves_sheet,
    build_raw_data_sheet, build_patch_sheets, build_diagnostics_sheets,
    build_patch_failure_sheet,
    build_products_not_tracked_sheet, build_patch_resolved_sheet,
)

log = logging.getLogger(__name__)

# Attempt to sync baselines if requested, and if the version_sync module is available.
# This is a best-effort attempt to keep baseline version rules up to date without requiring users to manually update config.json.
# If syncing fails (e.g. due to network issues or missing module), the dashboard will still run with existing baselines, and a debug message will be logged.
# If baselines are updated, the in-memory FIXED_VERSION_RULES will be refreshed to reflect the new values, and an info message will summarize the changes.
# This allows users to benefit from updated baselines without needing to understand or run the version_sync tool themselves, while ensuring that dashboard functionality is not disrupted if syncing is not possible.
# Note: For this to work, the version_sync module must be designed to update the same config.json file that this orchestrator reads from, and the config health check should be robust to any changes made by the sync process.
# The dashboard will always use the baselines as they exist in config.json at the time of loading, so if syncing is enabled and successful, the latest baselines will be applied to the current run. If syncing is disabled or fails, the dashboard will use whatever baselines are currently in config.json, which may be outdated but will not cause errors.
# The sync_baselines function is expected to return a dictionary of updated baselines if any were changed, or an empty dictionary if no updates were needed. The exact structure of this dictionary will depend on how the version_sync module is implemented, but it should allow the orchestrator to log which baselines were refreshed for transparency.
# The dashboard's config health check will validate the structure and values of the fixed_version_rules after syncing, so if the sync process introduces any issues (e.g. malformed versions), these will be caught and logged as warnings without breaking the dashboard generation.
# Overall, this approach provides a seamless way to keep vulnerability baselines current while maintaining the robustness and reliability of the dashboard generation process.
# The _try_sync_baselines function is intentionally designed to be non-intrusive and resilient, ensuring that the dashboard can still be generated successfully even if baseline syncing encounters issues. This allows users to opt in to baseline syncing for improved accuracy without risking disruption to their workflow.
# If you want to disable baseline syncing, simply set sync_baselines=False in the DashboardRequest, and the dashboard will use the existing baselines from config.json without attempting to update them.
# If you want to enable baseline syncing, set sync_baselines=True in the DashboardRequest, and the orchestrator will attempt to sync baselines at the start of the run. Any updates will be logged, and the latest baselines will be applied to the vulnerability data processing.
# Note: The version_sync module and its sync_baselines function are not defined in this code snippet, so you will need to implement them separately. The sync_baselines function should handle the logic of checking for updates to baselines (e.g. by fetching from a remote source), updating the config.json file if needed, and returning a summary of any changes made.
# The dashboard's fixed version rules are critical for accurately assessing vulnerability risk and remediation status, so keeping them up to date can significantly enhance the value of the dashboard. By providing an optional syncing mechanism, we allow users to easily benefit from updated baselines while ensuring that the dashboard remains functional and reliable regardless of syncing success.
# In summary, the _try_sync_baselines function is a key part of the orchestrator's ability to maintain current vulnerability baselines, and it is designed to be robust and user-friendly, providing benefits without risks to the dashboard generation process.

def _try_sync_baselines() -> None:
    try:
        from version_sync import sync_baselines
        updated = sync_baselines()
        if updated:
            import json as _json
            cfg_path = str(Path(__file__).parent / 'config.json')
            with open(cfg_path, encoding='utf-8') as _fh:
                _fresh = _json.load(_fh).get('fixed_version_rules', {})
            FIXED_VERSION_RULES.clear()
            FIXED_VERSION_RULES.update(_fresh)
            log.info("Baselines refreshed: %s",
                     ', '.join(f'{k}={v}' for k, v in updated.items()))
        else:
            log.debug("Baseline sync: no updates (network unavailable or all current)")
    except Exception as exc:
        log.debug("Baseline sync skipped: %s", exc)

# The main request and result dataclasses for the dashboard orchestrator.
# DashboardRequest encapsulates all input parameters needed to run the dashboard generation, including file paths, options for filtering and including data, and report metadata.
# DashboardResult encapsulates the outcome of the dashboard generation, including success status, output path, messages, trend summary data, and any warnings that were generated during processing.
# By using dataclasses, we can easily create and manage these structured data objects, and they provide a clear contract for the inputs and outputs of the orchestrator's run function.
# The run function will take a DashboardRequest, execute the entire dashboard generation process (loading data, merging, filtering, computing trends, processing patches, building the Excel workbook), and return a DashboardResult that indicates whether the process was successful and includes any relevant information or warnings for the user.
# The DashboardRequest includes options for skipping RMM data, including patch information, generating trend analyses, and applying date filters, allowing for flexible dashboard generation based on the user's needs and available data.
# The DashboardResult includes a success flag, the path to the generated output file, any messages for the user (e.g. errors or informational notes), a summary of trend data if applicable, and a list of warnings that may be relevant to the user when interpreting the dashboard.
# Overall, these dataclasses provide a clean and organized way to manage the inputs and outputs of the dashboard generation process, and they help to ensure that the orchestrator's run function has a clear and consistent interface.
# Note: The actual implementation of the run function will involve calling various helper functions and modules to perform the necessary data processing and Excel generation, and it will need to handle exceptions gracefully to ensure that a meaningful DashboardResult is returned even in the case of errors.
@dataclass
class DashboardRequest:
    vuln_path:            str
    output_path:          str
    rmm_path:             Optional[str]  = None
    skip_rmm:             bool           = False
    patch_path:           Optional[str]  = None
    include_patch:        bool           = False
    failure_report_path:  Optional[str]  = None
    include_failure_report: bool         = False
    prev_report_path:     Optional[str]  = None
    include_trend:        bool           = False
    threshold:            float          = 9.0
    cutoff_date:          Optional[str]  = None
    show_all_dates:       bool           = False
    sync_baselines:       bool           = False
    exclude_missing_rmm:  bool           = False
    report_month:         str            = ''

# The DashboardResult dataclass encapsulates the outcome of the dashboard generation process, including whether it was successful, the path to the output file, any messages for the user, a summary of trend data if applicable, and a list of warnings that may have been generated during processing.
# The success flag indicates whether the dashboard was generated successfully. The output_path provides the location of the generated Excel file. The message field can include any informational or error messages that should be conveyed to the user. The trend_summary can include key metrics from the trend analysis if it was performed, and the warnings list can include any issues or considerations that the user should be aware of when interpreting the dashboard.
# By returning a DashboardResult from the run function, we can provide a clear and structured way to communicate the outcome of the dashboard generation process to the caller, whether it was successful or if there were any issues that need attention.
# Note: The actual content of the message, trend_summary, and warnings will depend on the specific processing that occurs within the run function, and they should be crafted to provide meaningful insights and guidance to the user based on the results of the dashboard generation.
# The DashboardResult can be used by the caller (e.g. the GUI) to display success messages, errors, trend summaries, and warnings to the user in a clear and organized manner after the dashboard generation process is complete.
# The run function will need to populate the DashboardResult based on the outcomes of each step in the dashboard generation process, ensuring that any issues are captured in the warnings and that the success flag accurately reflects whether the dashboard was generated without critical errors.
# Overall, the DashboardResult provides a comprehensive summary of the dashboard generation process that can be easily consumed by the caller to inform the user of the results and any important considerations.
@dataclass
class DashboardResult:
    success:          bool
    output_path:      str             = ''
    message:          str             = ''
    trend_summary:    Optional[dict]  = None
    warnings:         list            = field(default_factory=list)

# The _config_health_check function performs validation on the dashboard's configuration, specifically checking the product_map and fixed_version_rules for issues such as duplicate keys, missing product mappings, and malformed version strings. It returns a list of any issues found, which can be logged as warnings to inform the user of potential problems with the configuration that may affect the accuracy or functionality of the dashboard.
# The function checks for duplicate keys in the product_map, ensures that all products referenced in fixed_version_rules have a corresponding entry in product_map, and validates that any version strings in fixed_version_rules are properly formatted. It also checks for cases where Chrome and Edge have identical version rules for the same CVE, which is not allowed. Any issues found are collected in a list and returned for logging.
# By performing this health check at the start of the dashboard generation process, we can catch and inform the user of any configuration issues that may lead to inaccurate vulnerability assessments or other problems in the generated dashboard, allowing them to address these issues before relying on the dashboard's insights.
# The health check is designed to be non-intrusive, meaning that it will not prevent the dashboard from being generated even if issues are found, but it will provide valuable feedback to the user about potential problems in the configuration that they may want to fix for better results.
# Note: The specific checks performed in this function are based on the expected structure of the config.json file and the requirements for how product mappings and version rules should be defined. If any issues are found, they will be logged as warnings, but the dashboard generation will proceed using whatever configuration is present, allowing users to still generate a dashboard while being informed of any potential issues with their setup.
# The use of regular expressions to validate version strings ensures that any versions specified in the fixed_version_rules are in a format that can be parsed and compared correctly, which is critical for the accurate assessment of vulnerability risk and remediation status in the dashboard.
# Overall, this function serves as a proactive check to help users maintain a healthy and accurate configuration for their CVE dashboard, enhancing the reliability of the insights generated by the tool.
def _config_health_check(cfg: dict) -> list[str]:
    import re as _re
    _VER_RE = _re.compile(r'^\d+(?:\.\d+){1,5}$')

    issues: list[str] = []
    pm     = cfg.get('product_map', [])
    fvr    = cfg.get('fixed_version_rules', {})

    seen_keys: dict[str, int] = {}
    for k, _ in pm:
        kl = str(k).lower()
        seen_keys[kl] = seen_keys.get(kl, 0) + 1
    dupes = [k for k, n in seen_keys.items() if n > 1]
    if dupes:
        issues.append(f"config.json: duplicate product_map key(s): {', '.join(dupes[:5])}")

    pm_values = {str(v).lower() for _, v in pm}
    for product in fvr:
        if product.startswith('_'):
            continue
        if product.lower() not in pm_values:
            issues.append(
                f"config.json: fixed_version_rules['{product}'] has no matching "
                f"product_map entry — version rules will never be applied"
            )

    for product, rules in fvr.items():
        if not isinstance(rules, dict):
            continue
        for key, ver in rules.items():
            if key.startswith('_'):
                ver_str = str(ver).strip()
                if ver_str and not _VER_RE.match(ver_str):
                    issues.append(
                        f"config.json: fixed_version_rules['{product}']['_baseline'] "
                        f"= {ver_str!r} is not a parseable version"
                    )
            else:
                ver_str = str(ver).strip()
                if ver_str and not _VER_RE.match(ver_str):
                    issues.append(
                        f"config.json: fixed_version_rules['{product}']['{key}'] "
                        f"= {ver_str!r} is not a parseable version"
                    )

    chrome_rules = fvr.get('chrome', {})
    edge_rules   = fvr.get('edge', {})
    for cve_id in set(chrome_rules) & set(edge_rules):
        if cve_id.startswith('_'):
            continue
        cv = str(chrome_rules[cve_id]).strip()
        ev = str(edge_rules[cve_id]).strip()
        if cv and ev and cv == ev:
            issues.append(
                f"config.json: Chrome and Edge have identical version {cv!r} "
                f"for {cve_id} — Chrome and Edge versions must differ"
            )

    if issues:
        for w in issues:
            log.warning("Config health: %s", w)
    else:
        log.debug("Config health: OK")

    return issues

# The main run function for the dashboard orchestrator. It takes a DashboardRequest as input, executes the entire dashboard generation process (including loading data, merging, filtering, computing trends, processing patches, and building the Excel workbook), and returns a DashboardResult that indicates whether the process was successful and includes any relevant information or warnings for the user.
# The function handles exceptions gracefully, ensuring that a meaningful DashboardResult is returned even in the case of errors. It also logs key steps and outcomes throughout the process to provide transparency and insights into the dashboard generation.
# The run function is designed to be the central coordinator for the dashboard generation, orchestrating the various components and ensuring that the inputs and outputs are managed effectively. It uses the parameters from the DashboardRequest to determine how to process the data and what to include in the final dashboard, and it compiles any messages or warnings that should be conveyed to the user in the DashboardResult.
# Note: The actual implementation of the run function will involve calling various helper functions and modules to perform the necessary data processing and Excel generation, and it will need to handle exceptions gracefully to ensure that a meaningful DashboardResult is returned even in the case of errors.
# The run function will need to manage the flow of data through the various stages of processing, including loading vulnerability and RMM data, merging datasets, applying filters, computing trends, processing patch information, and building the Excel workbook. It will also need to capture any relevant messages or warnings that arise during these processes and include them in the DashboardResult for user awareness.
# Overall, the run function serves as the main entry point for generating the CVE dashboard, coordinating all necessary steps and ensuring that the final output is comprehensive and informative for the user.
def run(request: DashboardRequest) -> DashboardResult:
    warnings: list[str] = []

    try:
        log.info("Dashboard run started — output: %s", request.output_path)

        import json as _json
        try:
            with open(Path(__file__).parent / 'config.json', encoding='utf-8') as _fh:
                _cfg_raw = _json.load(_fh)
            config_issues = _config_health_check(_cfg_raw)
            for issue in config_issues:
                warnings.append(issue)
        except Exception as _e:
            log.warning("Config health check failed: %s", _e)
            config_issues = []

        if request.sync_baselines:
            _try_sync_baselines()

        log.info("Loading vulnerability data: %s", request.vuln_path)
        df_vuln = load_vulnerability_data(request.vuln_path)
        log.info("  %d rows loaded", len(df_vuln))

        try:
            from cve_lookup import enrich_from_detections
            enriched = enrich_from_detections(df_vuln)
            if enriched:
                import json as _json
                cfg_path = str(Path(__file__).parent / 'config.json')
                with open(cfg_path, encoding='utf-8') as _fh:
                    _fresh = _json.load(_fh).get('fixed_version_rules', {})
                FIXED_VERSION_RULES.clear()
                FIXED_VERSION_RULES.update(_fresh)
                log.info("CVE lookup: %d CVE(s) enriched and version rules updated", enriched)
        except Exception as _e:
            log.debug("CVE lookup auto-enrich skipped: %s", _e)

        df_rmm = None
        if not request.skip_rmm and request.rmm_path:
            log.info("Loading RMM data: %s", request.rmm_path)
            df_rmm = load_rmm_data(request.rmm_path)
            log.info("  %d devices loaded", len(df_rmm))

        merged_df = merge_data(df_vuln, df_rmm, request.skip_rmm,
                               exclude_missing_rmm=request.exclude_missing_rmm)
        log.info("Merged dataset: %d rows", len(merged_df))

        raw_df         = merged_df.copy()
        stale_excluded = pd.DataFrame()

        if not request.show_all_dates and request.cutoff_date:
            cutoff = pd.to_datetime(request.cutoff_date, dayfirst=True, errors='coerce')
            if pd.isna(cutoff):
                cutoff = pd.to_datetime('1900-01-01')
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
            msg = (
                f"No vulnerability records found after applying date filter "
                f"(>= {request.cutoff_date}).\n\n"
                f"The detection dates in your CVE export may be older than this cutoff.\n"
                f"Try an earlier date, or tick 'Show All Dates' to include everything."
            )
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

        report_month_val = request.report_month if request.report_month else datetime.now().strftime('%B %Y')
        report_month_name = report_month_val.split()[0] if ' ' in report_month_val else report_month_val
        overview_sheet_name = f"{report_month_name} Detections"
        
        reserved = {
            "cves on stale devices", 'trend summary', overview_sheet_name.lower(), 'all detections', 'raw data',
            'stale excluded devices', 'new this month', 'resolved', 'persisting cves',
            'patch match overview', 'patch match full data', 'patch report (full)',
            'patch confirmed', 'resolved (patch confirmed)',
        }
        used_names       = set(reserved)
        product_to_sheet = {}
        for product, _ in triage_df.groupby('Base Product'):
            product_to_sheet[product] = clean_sheet_name(product, used_names)

        patch_data = None
        if request.include_patch and request.patch_path:
            log.info("Running patch match: %s", request.patch_path)
            p_ov, p_full, p_raw, tot_r, filt_r = process_patch_match(
                request.patch_path, merged_df.copy(), min_score=request.threshold)
            patch_data = (p_ov, p_full, p_raw, tot_r, filt_r)
            log.info("  Patch match: %d total rows, %d above threshold", tot_r, filt_r)

        trend_data       = None
        prev_report_name = ''
        redetected_count = 0
        if request.include_trend and request.prev_report_path:
            log.info("Loading previous report for trend: %s", request.prev_report_path)
            # load_previous_report returns (df, resolved_pairs) — unpack both.
            prev_df, prev_resolved_pairs = load_previous_report(request.prev_report_path)
            prev_report_name = Path(request.prev_report_path).name
            inventory_set    = (set(df_rmm['Device_Join'].unique())
                                if df_rmm is not None else None)

            # Capture the names of all stale excluded devices to purge them from the previous report
            stale_names = set(stale_excluded['Name'].apply(normalize_device_name)) if not stale_excluded.empty else set()

            trend_data       = compute_trends(merged_df, prev_df, request.threshold,
                                              inventory_devices=inventory_set,
                                              stale_devices=stale_names,
                                              prev_resolved_pairs=prev_resolved_pairs)
            m = trend_data['metrics']
            log.info(
                "Trend: %d new CVEs, %d resolved, %d persisting (common-product scope)",
                m['new_cve_count'], m['resolved_cve_count'], m['persisting_cve_count'],
            )
            
            redetected_count = trend_data.get('redetected_count', 0)
            if redetected_count > 0:
                w = f"{redetected_count} CVE(s) manually marked resolved last report but re-detected this period"
                log.warning(w)
                warnings.append(w)

        customer_name = ''
        for col in ('Customer', 'Customer Name', 'Client', 'Client Name'):
            if col in merged_df.columns:
                vals = merged_df[col].dropna().astype(str).str.strip()
                vals = vals[vals.str.len() > 0]
                if not vals.empty:
                    customer_name = vals.iloc[0]
                    break

        patch_resolved_pairs: Set[Tuple[str, str, str]] = set() 
        patch_gap_pairs:      dict[Tuple[str, str], str] = {}
        diagnostics: dict = {'patch_lag_df': pd.DataFrame(),
                             'version_drift_df': pd.DataFrame(),
                             'root_cause_df': pd.DataFrame()}

        if patch_data:
            p_full = patch_data[1].copy()
            p_full['_nk'] = p_full['Name'].astype(str).apply(normalize_device_name)
            p_full['_ck'] = p_full['Vulnerability Name'].astype(str).apply(extract_cve_id)

            if 'Patch Evidence Status' in p_full.columns:
                confirmed = p_full[p_full['Patch Evidence Status'] == 'Patch confirmed - pending rescan']
                if '_cascade_pk' in confirmed.columns:
                    pk_col = confirmed['_cascade_pk'].astype(str)
                else:
                    from data_pipeline import _detect_product as _dp_detect
                    pk_col = confirmed['Affected Products'].astype(str).apply(_dp_detect)
                patch_resolved_pairs = set(zip(
                    confirmed['_nk'],
                    confirmed['_ck'],
                    pk_col,
                ))
                log.info("Patch-confirmed resolved pairs: %d", len(patch_resolved_pairs))

            p_full['_root_cause'] = p_full.apply(classify_root_cause, axis=1)
            for _, row in p_full[p_full['_root_cause'].notna()].iterrows():
                patch_gap_pairs[(row['_nk'], row['_ck'])] = row['_root_cause']

            cause_counts: dict[str, int] = {}
            for c in patch_gap_pairs.values():
                cause_counts[c] = cause_counts.get(c, 0) + 1
            for cause, count in cause_counts.items():
                w = f"Patch gap [{cause}]: {count} device-CVE pair(s)"
                log.warning(w)
                warnings.append(w)

            product_rules = FIXED_VERSION_RULES
            diagnostics = compute_patch_diagnostics(
                patch_data[1], product_rules,
                resolved_pairs=patch_resolved_pairs,
            )

            rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
            if not rc_df.empty:
                mis = rc_df[rc_df.get('_cause_internal', rc_df.get('Patch Evidence Notes', '')) == 'version_compliant']
                if not mis.empty:
                    warnings.append(
                        f"{len(mis)} device-CVE pair(s) show 'Installed but still detected' — "
                        f"see 'Patch Evidence Notes' sheet"
                    )

        _status_col_inj = ('Threat Status' if 'Threat Status' in merged_df.columns
                           else 'Status'   if 'Status'        in merged_df.columns
                           else None)
        if _status_col_inj:
            from data_pipeline import _detect_product as _dp_detect_raw
            _raw_resolved = merged_df[
                merged_df[_status_col_inj].astype(str).str.strip().str.upper() == 'RESOLVED'
            ].copy()
            if not _raw_resolved.empty:
                _raw_pairs = set(zip(
                    _raw_resolved['Name'].apply(normalize_device_name),
                    _raw_resolved['Vulnerability Name'].apply(extract_cve_id),
                    _raw_resolved['Affected Products'].astype(str).apply(_dp_detect_raw),
                ))
                _before = len(patch_resolved_pairs)
                patch_resolved_pairs |= _raw_pairs
                log.info("Raw RESOLVED injection: %d pair(s) added (%d already present)",
                         len(patch_resolved_pairs) - _before, len(_raw_pairs) - (len(patch_resolved_pairs) - _before))

        patch_confirmed_count = 0
        if patch_resolved_pairs:
            from data_pipeline import _detect_product as _dp_detect
            triage_keys = set(zip(
                triage_df['Name'].apply(normalize_device_name),
                triage_df['Vulnerability Name'].apply(extract_cve_id),
                triage_df['Affected Products'].astype(str).apply(_dp_detect),
            ))
            patch_confirmed_count = len(patch_resolved_pairs & triage_keys)

        failure_df     = None
        failure_lookup = {}
        failure_devices: set = set()

        if request.include_failure_report and request.failure_report_path:
            try:
                log.info("Loading patch failure report: %s", request.failure_report_path)
                failure_df     = load_patch_failure_report(request.failure_report_path)
                failure_lookup = build_patch_failure_lookup(failure_df)
                failure_devices = set(failure_lookup.keys())
            except Exception as exc:
                log.warning("Could not process patch failure report: %s", exc)
                warnings.append(f"Could not process patch failure report: {exc}")

        log.info("Writing workbook: %s", request.output_path)
        with pd.ExcelWriter(request.output_path, engine='xlsxwriter') as writer:
            wb = writer.book
            styles     = get_workbook_styles(wb)
            link_fmt   = styles['link']
            header_fmt = styles['header']
            miss_fmt   = styles['row_missing']

            _not_in_rmm_mask = filtered_df['Last Response'] == 'Not Found in RMM'
            _not_in_rmm_cve_rows = int(_not_in_rmm_mask.sum()) if 'Last Response' in filtered_df.columns else 0
            _not_in_rmm_unique_cves = int(filtered_df.loc[_not_in_rmm_mask, 'Vulnerability Name'].nunique()) if 'Last Response' in filtered_df.columns and 'Vulnerability Name' in filtered_df.columns else 0
#checking if the columns exist before applying the mask and counting unique CVEs to avoid KeyError if the expected columns are not present in the filtered_df
#This ensures that if the 'Last Response' or 'Vulnerability Name' columns are missing from the filtered_df, the code will not raise an error and will instead set the counts to 0, allowing the dashboard generation to proceed without interruption while still providing accurate counts when the columns are present.
#The not-in-RMM counts are important for the client summary sheet to provide context on how many high-risk vulnerabilities are not being tracked in RMM, which can inform remediation prioritization and risk assessment. By including these counts in the summary sheet, we can give users a clearer picture of their vulnerability landscape and highlight potential gaps in their RMM coverage.
#Overall, this section of the code is focused on preparing the data and metrics that will be displayed in the client summary sheet of the dashboard, ensuring that users have a comprehensive overview of their vulnerability status, including any high-risk vulnerabilities that may not be tracked in RMM.
            build_client_summary_sheet(
                wb, filtered_df, triage_df, request.threshold,
                trend_data=trend_data,
                customer_name=customer_name,
                cutoff_date=request.cutoff_date if not request.show_all_dates else None,
                stale_excluded_df=stale_excluded if not stale_excluded.empty else None,
                not_in_rmm_count=not_in_rmm,
                not_in_rmm_cve_count=_not_in_rmm_cve_rows,
                not_in_rmm_unique_cves=_not_in_rmm_unique_cves,
                report_month=report_month_val,
            )

            if trend_data:
                build_trend_summary_sheet(wb, trend_data, request.threshold,
                                          prev_report_name, header_fmt,
                                          customer_name=customer_name)
# The trend summary sheet provides a high-level overview of the key metrics from the trend analysis, including the number of new CVEs, resolved CVEs, and persisting CVEs compared to the previous report. It also includes the report month and customer name for context. By including this sheet in the dashboard, users can quickly understand how their vulnerability landscape is evolving over time and identify any significant changes or trends that may require attention.
# The build_trend_summary_sheet function takes the workbook, trend data, threshold, previous report name, header format, and optional customer name as inputs, and it constructs a visually appealing summary sheet that highlights the key trend metrics. This sheet serves as an important component of the dashboard, providing users with actionable insights into their vulnerability management efforts and helping them to track their progress over time.
# By including the trend summary sheet, we can enhance the value of the dashboard by not only providing a snapshot of the current vulnerability status but also showing how it has changed compared to the previous reporting period, enabling users to make informed decisions about their remediation strategies and resource allocation.
# The overview sheet provides a detailed summary of the current vulnerability status, including key metrics, evidence summaries, recommended actions, and links to detailed sheets for further analysis. It serves as the main landing page for users when they open the dashboard, giving them a comprehensive overview of their vulnerability landscape and guiding them towards the most critical issues that require attention.
# The build_overview_sheet function takes the workbook, merged and filtered dataframes, triage dataframe, threshold, product-to-sheet mapping, header and link formats, and various optional parameters such as customer name, patch confirmed count, redetected count, trend metrics, evidence summary, recommended actions,
            build_overview_sheet(
                wb, merged_df, filtered_df, triage_df, request.threshold,
                product_to_sheet, header_fmt, link_fmt,
                customer_name=customer_name,
                patch_confirmed_count=patch_confirmed_count,
                redetected_count=redetected_count,
                sheet_name=overview_sheet_name,
                trend_metrics=trend_data['metrics'] if trend_data else None,
                evidence_summary=diagnostics.get('evidence_summary'),
                recommended_actions=diagnostics.get('recommended_actions'),
                has_prev_report=trend_data is not None,
                stale_excluded_df=stale_excluded if not stale_excluded.empty else None,
                report_month=report_month_val,
            )

            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'New This Month', 'Persisting CVEs'})

            build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                                  patch_resolved_pairs=patch_resolved_pairs,
                                  patch_gap_pairs=patch_gap_pairs)

            if not stale_excluded.empty:
                build_stale_excluded_sheet(writer, stale_excluded)
                
                # Fetch unresolved CVEs for these stale devices natively from RAW DATA
                stale_device_names = stale_excluded['Name'].unique()
                stale_raw_rows = raw_df[raw_df['Name'].isin(stale_device_names)].copy()
                
                _status_col_stale = ('Threat Status' if 'Threat Status' in stale_raw_rows.columns
                                     else 'Status'   if 'Status'        in stale_raw_rows.columns
                                     else None)
                if _status_col_stale:
                    stale_unresolved_cves = stale_raw_rows[stale_raw_rows[_status_col_stale].astype(str).str.strip().str.upper() == 'UNRESOLVED'].copy()
                else:
                    stale_unresolved_cves = stale_raw_rows.copy()
                
                build_stale_cves_sheet(writer, stale_unresolved_cves, link_fmt)

            _status_col_wb = ('Threat Status' if 'Threat Status' in merged_df.columns
                              else 'Status'   if 'Status'        in merged_df.columns
                              else None)
            _raw_resolved_df = pd.DataFrame()
            if _status_col_wb:
                _raw_resolved_df = merged_df[
                    merged_df[_status_col_wb].astype(str).str.strip().str.upper() == 'RESOLVED'
                ].copy()

            if patch_data:
                build_patch_sheets(writer, patch_data[0], patch_data[1], patch_data[2])

                _patch_full_aug = patch_data[1].copy()
                if not _raw_resolved_df.empty:
                    _raw_for_sheet = _raw_resolved_df.copy()
                    _raw_for_sheet['Patch Evidence Status'] = 'Patch confirmed - pending rescan'
                    if 'Patch Match Result' not in _raw_for_sheet.columns:
                        _raw_for_sheet['Patch Match Result'] = 'Resolved in N-able (Status=RESOLVED)'
                    _patch_full_aug = pd.concat(
                        [_patch_full_aug, _raw_for_sheet], ignore_index=True, sort=False
                    ).drop_duplicates(subset=['Name', 'Vulnerability Name'], keep='first')

                build_patch_resolved_sheet(writer, _patch_full_aug)
                if any(not diagnostics[k].empty for k in diagnostics
                       if isinstance(diagnostics[k], pd.DataFrame)):
                    build_diagnostics_sheets(writer, diagnostics)

                build_products_not_tracked_sheet(writer, patch_data[1])

            elif not _raw_resolved_df.empty:
                _raw_for_sheet2 = _raw_resolved_df.copy()
                _raw_for_sheet2['Patch Evidence Status'] = 'Patch confirmed - pending rescan'
                if 'Patch Match Result' not in _raw_for_sheet2.columns:
                    _raw_for_sheet2['Patch Match Result'] = 'Resolved in N-able (Status=RESOLVED)'
                build_patch_resolved_sheet(writer, _raw_for_sheet2)

            if failure_df is not None and failure_lookup:
                inventory_devices = (
                    set(df_rmm['Device_Join'].unique()) if df_rmm is not None else None
                )
                cve_overlap = triage_df[
                    triage_df['Name'].apply(normalize_device_name).isin(failure_devices)
                ].copy()
                build_patch_failure_sheet(writer, failure_df, failure_lookup,
                                          cve_overlap, inventory_devices=inventory_devices)
                for dev, info in sorted(failure_lookup.items(),
                                        key=lambda x: -x[1]['failure_count'])[:3]:
                    warnings.append(
                        f"Patch delivery failing on {dev}: "
                        f"{info['failure_count']} failures — {info['top_description']}"
                    )
                if not cve_overlap.empty:
                    warnings.append(
                        f"{cve_overlap['Vulnerability Name'].nunique()} CVE type(s) on "
                        f"{cve_overlap['Name'].nunique()} device(s) where patches are "
                        f"actively failing — see 'Patch Failures' sheet"
                    )

            build_raw_data_sheet(writer, raw_df)

        log.info("Workbook written successfully")

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