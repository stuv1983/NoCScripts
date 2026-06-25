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
    load_browser_audit, merge_browser_audit_into_drift,
    _drop_internal,
)
from diagnostics import compute_patch_diagnostics, classify_root_cause
import snapshot as snap_store
from excel_builder import (
    get_workbook_styles,
    build_client_summary_sheet,
    build_trend_summary_sheet, build_trend_detail_sheets,
    build_all_detections_sheet,
    build_product_sheets, build_stale_excluded_sheet,
    build_stale_cves_sheet,
    build_patch_sheets, build_diagnostics_sheets,
    build_patch_failure_sheet,
    build_products_not_tracked_sheet, build_patch_resolved_sheet,
    build_device_report_sheet,
)

log = logging.getLogger(__name__)


def _find_cve_repo() -> 'Optional[Path]':
    """Locate the local cvelistV5 git clone in standard locations."""
    candidates = [
        Path(r'C:\NoCScripts\N-able Tools\CVE_Risk_Exposure_&_Remediation\cvelistV5'),
        Path(__file__).resolve().parent / 'cvelistV5',
        Path(__file__).resolve().parent.parent / 'cvelistV5',
    ]
    return next((p for p in candidates if p.exists()), None)


def _pull_cve_repo(repo: 'Optional[Path]') -> None:
    """Run git pull --ff-only on the cvelistV5 repo (non-fatal on any failure)."""
    if repo is None:
        log.debug('cvelistV5 repo not found — skipping git pull')
        return
    import subprocess
    try:
        r = subprocess.run(
            ['git', '-C', str(repo), 'pull', '--ff-only'],
            capture_output=True, text=True, timeout=30,
        )
        msg = (r.stdout.strip() or r.stderr.strip() or '(no output)').splitlines()[0]
        if r.returncode == 0:
            log.info('cvelistV5 pull: %s', msg)
        else:
            log.warning('cvelistV5 pull failed (rc=%d): %s', r.returncode, msg[:120])
    except subprocess.TimeoutExpired:
        log.warning('cvelistV5 pull timed out (30s) — continuing with local data')
    except FileNotFoundError:
        log.debug('git not on PATH — skipping cvelistV5 pull')
    except Exception as _e:
        log.debug('cvelistV5 pull error: %s', _e)


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
    browser_audit_path:   Optional[str]  = None
    include_browser_audit: bool          = False
    prev_report_path:     Optional[str]  = None
    include_trend:        bool           = False
    threshold:            float          = 9.0
    cutoff_date:          Optional[str]  = None
    show_all_dates:       bool           = False
    sync_baselines:       bool           = False
    exclude_missing_rmm:  bool           = False
    report_month:         str            = ''
    stale_warning_days:   int            = 14   # flag active devices within this many days of going stale

@dataclass
class DashboardResult:
    success:          bool
    output_path:      str             = ''
    message:          str             = ''
    trend_summary:    Optional[dict]  = None
    warnings:         list            = field(default_factory=list)

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

        _cve_repo = _find_cve_repo()
        _pull_cve_repo(_cve_repo)

        try:
            from cve_lookup import enrich_from_detections
            enriched = enrich_from_detections(df_vuln, cve_repo_path=_cve_repo)
            if enriched:
                import json as _json
                cfg_path = str(Path(__file__).parent / 'config.json')
                with open(cfg_path, encoding='utf-8') as _fh:
                    _fresh = _json.load(_fh)
                FIXED_VERSION_RULES.clear()
                FIXED_VERSION_RULES.update(_fresh.get('fixed_version_rules', {}))
                # Also refresh the in-memory CVSS cache
                from config import _CONFIG as _orch_cfg
                _orch_cfg['cvss_score_cache'] = _fresh.get('cvss_score_cache', {})
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
        approaching_stale_names: set = set()
        _STALE_DAYS = 30   # fixed staleness threshold (days without a response)

        if not request.show_all_dates and request.cutoff_date:
            cutoff = pd.to_datetime(request.cutoff_date, dayfirst=True, errors='coerce')
            if pd.isna(cutoff):
                cutoff = pd.to_datetime('1900-01-01')
            high = merged_df[merged_df['Vulnerability Score'] >= request.threshold]

            # ── Pass 1: date-filter stale (last seen before cutoff_date) ────────
            stale_by_date = high[
                (high['_Sort_Time'] < cutoff) &
                (high['Last Response'] != 'Not Found in RMM')
            ].copy()

            # ── Pass 2: days-stale (last seen after cutoff but >= 30 days ago) ──
            if 'Days Since Last Response' in merged_df.columns:
                _days_col = pd.to_numeric(merged_df['Days Since Last Response'], errors='coerce')
                stale_names_by_days = set(
                    merged_df.loc[
                        (merged_df['Last Response'] != 'Not Found in RMM') &
                        (_days_col >= _STALE_DAYS),
                        'Name'
                    ].unique()
                )
            else:
                stale_names_by_days = set()

            stale_by_days_df = high[
                high['Name'].isin(stale_names_by_days) &
                (high['Last Response'] != 'Not Found in RMM') &
                ~high['Name'].isin(set(stale_by_date['Name'].unique()))
            ].copy()

            stale_excluded = pd.concat([stale_by_date, stale_by_days_df], ignore_index=True)
            all_stale_names = set(stale_excluded['Name'].unique())

            # Remove ALL stale devices (both passes) from the working dataset
            merged_df = merged_df[
                (~merged_df['Name'].isin(all_stale_names)) |
                (merged_df['Last Response'] == 'Not Found in RMM')
            ]

            log.info(
                "Date filter applied (>= %s): %d rows kept, "
                "%d stale excluded (%d by date-filter, %d by %d-day rule)",
                request.cutoff_date, len(merged_df),
                len(all_stale_names), len(stale_by_date['Name'].unique()),
                len(stale_names_by_days - set(stale_by_date['Name'].unique())),
                _STALE_DAYS,
            )

        # ── Approaching stale: DISABLED FOR TESTING ─────────────────────────────
        warning_days = max(1, int(request.stale_warning_days))
        approaching_stale_names: set = set()
        # if 'Days Since Last Response' in merged_df.columns:
        #     _days_col_ap  = pd.to_numeric(merged_df['Days Since Last Response'], errors='coerce')
        #     _active_mask  = merged_df['Last Response'] != 'Not Found in RMM'
        #     approaching_stale_names = set(
        #         merged_df.loc[
        #             _active_mask & (_days_col_ap >= warning_days),
        #             'Name'
        #         ].unique()
        #     )
        # log.info(
        #     "%d device(s) flagged as approaching stale (offline >= %d days)",
        #     len(approaching_stale_names), warning_days,
        # )

        if merged_df.empty:
            msg = (
                f"No vulnerability records found after applying date filter "
                f"(>= {request.cutoff_date}).\n\n"
                f"The detection dates in your CVE export may be older than this cutoff.\n"
                f"Try an earlier date, or tick 'Show All Dates' to include everything."
            )
            log.warning(msg)
            return DashboardResult(success=False, message=msg)

        # These are views — confirmed by audit that excel_builder.py only reads
        # from them via .loc[mask, col] (nunique/count), never writes.
        # Adding a write to any of these later will raise SettingWithCopyWarning.
        filtered_df = merged_df[merged_df['Vulnerability Score'] >= request.threshold]
        triage_df   = filtered_df[filtered_df['Last Response'] != 'Not Found in RMM']

        # Build a dedicated DataFrame for not-in-RMM devices so we can pass rows
        # (not just a count) into the stale sheet builders for audit tracking.
        not_in_rmm_df   = filtered_df[filtered_df['Last Response'] == 'Not Found in RMM']
        not_in_rmm      = int(not_in_rmm_df['Name'].nunique())
        if not_in_rmm:
            w = f"{not_in_rmm} device(s) with score ≥ {request.threshold} not found in RMM — excluded from triage sheets"
            log.warning(w)
            warnings.append(w)

        log.info(
            "Filtered (score >= %.1f): %d rows, %d triage, %d not-in-RMM",
            request.threshold, len(filtered_df), len(triage_df), not_in_rmm,
        )

        report_month_val = request.report_month if request.report_month else datetime.now().strftime('%B %Y')

        reserved = {
            "cves on stale devices", 'trend summary', 'all detections', 'raw data',
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
            prev_df, prev_resolved_pairs, prev_source_type = load_previous_report(request.prev_report_path)
            prev_report_name = Path(request.prev_report_path).name
            inventory_set    = (set(df_rmm['Device_Join'].unique())
                                if df_rmm is not None else None)

            # Capture the names of all stale excluded devices to purge them from the previous report
            stale_names = set(stale_excluded['Name'].apply(normalize_device_name)) if not stale_excluded.empty else set()

            trend_data       = compute_trends(merged_df, prev_df, request.threshold,
                                              inventory_devices=inventory_set,
                                              stale_devices=stale_names,
                                              prev_resolved_pairs=prev_resolved_pairs,
                                              prev_source_type=prev_source_type)
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

            # Merge browser audit data into version drift if provided
            if request.include_browser_audit and request.browser_audit_path:
                try:
                    browser_audit_df = load_browser_audit(request.browser_audit_path)
                    if not browser_audit_df.empty:
                        diagnostics['version_drift_df'] = merge_browser_audit_into_drift(
                            diagnostics.get('version_drift_df', pd.DataFrame()),
                            browser_audit_df,
                        )
                        log.info("Browser audit merged: %d device records", len(browser_audit_df))
                except Exception as exc:
                    log.warning("Could not process browser audit: %s", exc)
                    warnings.append(f"Could not process browser audit: {exc}")

            rc_df = diagnostics.get('root_cause_df', pd.DataFrame())
            if not rc_df.empty:
                mis = rc_df[rc_df.get('_cause_internal', rc_df.get('Patch Evidence Notes', '')) == 'version_compliant']
                if not mis.empty:
                    warnings.append(
                        f"{len(mis)} device-CVE pair(s) show 'Installed but still detected' — "
                        f"see 'Patch Evidence Notes' sheet"
                    )

        # ── Bulletproof raw scanner override ────────────────────────────────────
        # Read directly from raw_df (pre-filter, pre-join source of truth) so no
        # date-filter or RMM-join step can silently drop an UNRESOLVED row and let
        # a false-positive blue ☑ survive.  2-tuple (device, cve) matching means
        # product-name formatting differences can never cause a miss.
        from data_pipeline import _detect_product as _dp_detect_raw
        _unresolved_pairs_2d: set = set()
        _raw_inject_pairs:    set = set()

        for _col in ('Threat Status', 'Status', 'threat status', 'status'):
            if _col not in raw_df.columns:
                continue
            _col_upper = raw_df[_col].astype(str).str.strip().str.upper()
            _raw_unr = raw_df[_col_upper == 'UNRESOLVED']
            if not _raw_unr.empty:
                _unresolved_pairs_2d |= set(zip(
                    _raw_unr['Name'].apply(normalize_device_name),
                    _raw_unr['Vulnerability Name'].apply(extract_cve_id),
                ))
            _raw_res = raw_df[_col_upper == 'RESOLVED']
            if not _raw_res.empty:
                _raw_inject_pairs |= set(zip(
                    _raw_res['Name'].apply(normalize_device_name),
                    _raw_res['Vulnerability Name'].apply(extract_cve_id),
                    _raw_res['Affected Products'].astype(str).apply(_dp_detect_raw),
                ))

        # Step 1: strip false positives from patch tool memory.
        # If the scanner says UNRESOLVED for (device, cve), remove every matching
        # 3-tuple from patch_resolved_pairs regardless of product string.
        if _unresolved_pairs_2d and patch_resolved_pairs:
            to_remove = {p for p in patch_resolved_pairs if (p[0], p[1]) in _unresolved_pairs_2d}
            if to_remove:
                patch_resolved_pairs -= to_remove
                log.info(
                    "Scanner override: removed %d pair(s) from patch_resolved_pairs "
                    "because raw_df still shows UNRESOLVED — will render as red ☐",
                    len(to_remove),
                )

        # Step 2: inject raw RESOLVED pairs, skipping any (device, cve) that is
        # still UNRESOLVED in the scanner (UNRESOLVED always wins).
        if _raw_inject_pairs:
            clean = {p for p in _raw_inject_pairs if (p[0], p[1]) not in _unresolved_pairs_2d}
            skipped = len(_raw_inject_pairs) - len(clean)
            if skipped:
                log.info("Raw injection: skipped %d pair(s) where scanner also shows UNRESOLVED", skipped)
            _before = len(patch_resolved_pairs)
            patch_resolved_pairs |= clean
            log.info("Raw RESOLVED injection: %d pair(s) added", len(patch_resolved_pairs) - _before)

        patch_confirmed_count = 0
        if patch_resolved_pairs:
            from data_pipeline import _detect_product as _dp_detect
            triage_keys = set(zip(
                triage_df['Name'].apply(normalize_device_name),
                triage_df['Vulnerability Name'].apply(extract_cve_id),
                triage_df['Affected Products'].astype(str).apply(_dp_detect),
            ))
            # Count UNIQUE (device, cve) pairs confirmed — not 3-tuples.
            # patch_resolved_pairs uses (device, cve, product) 3-tuples so the
            # same device+CVE pair can appear multiple times (once per product
            # that matches it, e.g. Chrome AND Edge both resolving CVE-2024-X).
            # Counting 3-tuples produces a number larger than n_total (which is
            # 2-tuple unique pairs), causing Unresolved = Total - Resolved < 0.
            _confirmed_3tuples = patch_resolved_pairs & triage_keys
            patch_confirmed_count = len({(d, v) for d, v, _ in _confirmed_3tuples})

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
                approaching_stale_names=approaching_stale_names,
                stale_warning_days=request.stale_warning_days,
                product_to_sheet=product_to_sheet,
            )
            if trend_data:
                build_trend_summary_sheet(wb, trend_data, request.threshold,
                                          prev_report_name, header_fmt,
                                          customer_name=customer_name)

            if trend_data:
                build_trend_detail_sheets(writer, wb, trend_data, link_fmt,
                                          sheets_subset={'New This Month', 'Persisting CVEs'})

            build_product_sheets(writer, triage_df, product_to_sheet, link_fmt,
                                  patch_resolved_pairs=patch_resolved_pairs,
                                  patch_gap_pairs=patch_gap_pairs,
                                  approaching_stale_names=approaching_stale_names,
                                  stale_warning_days=request.stale_warning_days)

            if not stale_excluded.empty or not not_in_rmm_df.empty:
                build_stale_excluded_sheet(writer, stale_excluded,
                                           not_in_rmm_df=not_in_rmm_df if not not_in_rmm_df.empty else None)

                # Fetch unresolved CVEs for stale-date devices from RAW DATA
                stale_device_names  = stale_excluded['Name'].unique() if not stale_excluded.empty else []
                stale_raw_rows      = raw_df[raw_df['Name'].isin(stale_device_names)].copy()

                _status_col_stale = ('Threat Status' if 'Threat Status' in stale_raw_rows.columns
                                     else 'Status'   if 'Status'        in stale_raw_rows.columns
                                     else None)
                if _status_col_stale and not stale_raw_rows.empty:
                    stale_unresolved_cves = stale_raw_rows[stale_raw_rows[_status_col_stale].astype(str).str.strip().str.upper() == 'UNRESOLVED'].copy()
                else:
                    stale_unresolved_cves = stale_raw_rows.copy()

                # Fetch unresolved CVEs for not-in-RMM devices from RAW DATA
                not_in_rmm_names   = not_in_rmm_df['Name'].unique() if not not_in_rmm_df.empty else []
                not_in_rmm_raw     = raw_df[raw_df['Name'].isin(not_in_rmm_names)].copy()
                _status_col_nirm   = ('Threat Status' if 'Threat Status' in not_in_rmm_raw.columns
                                      else 'Status'   if 'Status'        in not_in_rmm_raw.columns
                                      else None)
                if _status_col_nirm and not not_in_rmm_raw.empty:
                    not_in_rmm_cves = not_in_rmm_raw[not_in_rmm_raw[_status_col_nirm].astype(str).str.strip().str.upper() == 'UNRESOLVED'].copy()
                else:
                    not_in_rmm_cves = not_in_rmm_raw.copy()

                build_stale_cves_sheet(writer, stale_unresolved_cves, link_fmt,
                                       not_in_rmm_cves_df=not_in_rmm_cves if not not_in_rmm_cves.empty else None)

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

            # Raw Data is written as a CSV sidecar alongside the Excel file
            # instead of a sheet.  Writing 83k × 19 cols to xlsxwriter was the
            # single largest write-time cost (~35s).  The CSV is used identically
            # by load_previous_report next month — it accepts both .xlsx and .csv.
            _raw_csv_path = (
                Path(request.output_path).parent
                / (Path(request.output_path).stem + '_RawData.csv')
            )
            try:
                _drop_internal(raw_df).to_csv(_raw_csv_path, index=False)
                log.info('Raw data written to %s (%d rows)', _raw_csv_path.name, len(raw_df))
            except Exception as _csv_err:
                log.warning('Could not write raw data CSV: %s', _csv_err)
            if df_rmm is not None:
                build_device_report_sheet(writer, df_rmm)

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