"""
contracts.py — pipeline boundary contracts.

One named assertion per historical failure class, executed at the pipeline
boundary where the bug lived:

  • v0.3  triage scope bug            → check_scopes()
  • v0.4  Status column collision     → check_patch_match()  (both directions)
  • v0.6+ column drift across exports → check_cve_export() / check_merged()
  • v0.12 category dtype round-trip   → check_merged(expect_category_cols=...)
  • join-key / internal-column leaks  → forbidden= on every boundary

Design
------
Two severities:

  STRUCTURAL — the frame cannot be correct (missing required column,
  duplicate column labels, scope-containment broken, a forbidden internal
  column leaked). These raise ContractError immediately: continuing would
  produce a plausible-looking wrong workbook, which is the worst outcome
  for a customer-facing report.

  SEMANTIC — the frame is structurally fine but a value vocabulary looks
  wrong (unknown Threat Status value, patch Status containing CVE threat
  statuses, Device Type outside its known set). These are returned as
  warning strings so the orchestrator can append them to
  DashboardResult.warnings — a new N-able export vocabulary should be
  surfaced loudly, not crash the run. The one exception is the v0.4
  collision signature (patch Status ⊆ {RESOLVED, UNRESOLVED}), which is
  promoted to STRUCTURAL because it is never legitimate.

Every check function returns list[str] (semantic warnings) and raises
ContractError on structural violations. No pandas mutation, ever — all
checks are read-only.

Zero dependencies beyond pandas. No imports from other project modules, so
this file can be imported by anything without cycles.
"""

from __future__ import annotations

import logging
from typing import Iterable, Optional, Sequence

import pandas as pd

log = logging.getLogger(__name__)


class ContractError(ValueError):
    """A pipeline-boundary invariant was violated. Message names the
    boundary, the column(s), and the historical bug it guards, so the
    person seeing it in a GUI error box knows what file to fix."""


# ══════════════════════════════════════════════════════════════════════════════
# Generic building blocks
# ══════════════════════════════════════════════════════════════════════════════

def check_frame(df: pd.DataFrame,
                name: str,
                *,
                required: Sequence[str] = (),
                forbidden: Sequence[str] = (),
                no_duplicate_columns: bool = True,
                numeric: Sequence[str] = (),
                allow_empty: bool = True) -> None:
    """Structural checks only — raises ContractError, returns nothing."""
    if not isinstance(df, pd.DataFrame):
        raise ContractError(f"{name}: expected DataFrame, got {type(df).__name__}")

    if not allow_empty and df.empty:
        raise ContractError(f"{name}: frame is empty")

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ContractError(
            f"{name}: missing required column(s) {missing} — "
            f"present: {sorted(map(str, df.columns))[:25]}"
        )

    leaked = [c for c in forbidden if c in df.columns]
    if leaked:
        raise ContractError(
            f"{name}: internal/forbidden column(s) leaked downstream: {leaked}"
        )

    if no_duplicate_columns and df.columns.duplicated().any():
        dupes = df.columns[df.columns.duplicated()].tolist()
        raise ContractError(
            f"{name}: duplicate column label(s) {dupes} — usually two source "
            f"aliases both renamed to the same canonical name "
            f"(see _rename_cve_columns single-assignment rule)"
        )

    for col in numeric:
        if col in df.columns and not df.empty:
            coerced = pd.to_numeric(df[col], errors='coerce')
            n_bad = int(coerced.isna().sum() - df[col].isna().sum())
            if n_bad > 0:
                raise ContractError(
                    f"{name}: column '{col}' has {n_bad} non-numeric value(s) "
                    f"that were expected to be numeric by this stage"
                )


def check_vocabulary(df: pd.DataFrame,
                     name: str,
                     col: str,
                     *,
                     allowed: Optional[Iterable[str]] = None,
                     forbidden_values: Iterable[str] = (),
                     normalise: bool = True,
                     ignore_blank: bool = True,
                     hard: bool = False) -> list[str]:
    """
    Value-vocabulary check on one column.

    allowed          — values outside this set produce a warning (or raise
                       if hard=True).
    forbidden_values — values that must never appear; always structural
                       when hard=True, otherwise warned.
    normalise        — compare on str().strip().upper().
    Returns list of warning strings (empty if clean or column absent).
    """
    if col not in df.columns or df.empty:
        return []

    series = df[col].dropna().astype(str).str.strip()
    if ignore_blank:
        series = series[series.str.len() > 0]
    if normalise:
        observed = set(series.str.upper().unique())
        allowed_n = {str(v).strip().upper() for v in allowed} if allowed is not None else None
        forbidden_n = {str(v).strip().upper() for v in forbidden_values}
    else:
        observed = set(series.unique())
        allowed_n = set(allowed) if allowed is not None else None
        forbidden_n = set(forbidden_values)

    issues: list[str] = []

    hit = observed & forbidden_n
    if hit:
        msg = (f"{name}: column '{col}' contains forbidden value(s) "
               f"{sorted(hit)}")
        if hard:
            raise ContractError(msg)
        issues.append(msg)

    if allowed_n is not None:
        unknown = observed - allowed_n
        if unknown:
            msg = (f"{name}: column '{col}' has unrecognised value(s) "
                   f"{sorted(unknown)[:8]} — expected one of {sorted(allowed_n)}")
            if hard:
                raise ContractError(msg)
            issues.append(msg)

    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 1 — CVE export, after load_vulnerability_data()
# ══════════════════════════════════════════════════════════════════════════════

# Vocabulary observed across N-able export variants. Deliberately warn-only:
# a new N-able value must surface as a warning, not kill the run.
THREAT_STATUS_VOCAB = {'RESOLVED', 'UNRESOLVED'}
YES_NO_VOCAB = {'YES', 'NO', 'TRUE', 'FALSE', 'Y', 'N', '1', '0'}


def check_cve_export(df: pd.DataFrame, source_name: str = 'cve_export') -> list[str]:
    """
    After load_vulnerability_data(). Asserts the alias-normalisation landed:
    canonical names present, no duplicate labels (the historical
    'Client' + 'Customer Name' → two 'Customer' columns bug), numeric score.
    """
    check_frame(
        df, source_name,
        required=('Name', 'Vulnerability Name', 'Vulnerability Score'),
        numeric=('Vulnerability Score',),
        allow_empty=False,
    )
    issues: list[str] = []
    issues += check_vocabulary(df, source_name, 'Threat Status',
                               allowed=THREAT_STATUS_VOCAB)
    issues += check_vocabulary(df, source_name, 'Has Known Exploit',
                               allowed=YES_NO_VOCAB)
    issues += check_vocabulary(df, source_name, 'CISA KEV',
                               allowed=YES_NO_VOCAB)

    if not df.empty:
        scores = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        out_of_range = int(((scores < 0) | (scores > 10)).sum())
        if out_of_range:
            issues.append(
                f"{source_name}: {out_of_range} row(s) with Vulnerability Score "
                f"outside 0–10 — check the score column alias mapping"
            )
        blank_names = int((df['Name'].astype(str).str.strip() == '').sum())
        if blank_names:
            issues.append(f"{source_name}: {blank_names} row(s) with blank device Name")

    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 2 — RMM inventory, after load_rmm_data()
# ══════════════════════════════════════════════════════════════════════════════

DEVICE_TYPE_VOCAB = {'SERVER', 'WORKSTATION', 'UNKNOWN'}


def check_rmm_inventory(df: pd.DataFrame, source_name: str = 'rmm_inventory') -> list[str]:
    check_frame(
        df, source_name,
        required=('Device', 'Last Response', 'Device_Join', 'Device Type', 'Username'),
        allow_empty=False,
    )
    issues = check_vocabulary(df, source_name, 'Device Type',
                              allowed=DEVICE_TYPE_VOCAB)
    if not df.empty and df['Device_Join'].duplicated().any():
        # load_rmm_data drop_duplicates on Device_Join — if this fires the
        # dedup was removed or bypassed and the CVE merge will fan out rows.
        raise ContractError(
            f"{source_name}: duplicate Device_Join keys — merge_data would "
            f"duplicate CVE rows on join (drop_duplicates contract broken)"
        )
    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 3 — merged frame, after merge_data()
# ══════════════════════════════════════════════════════════════════════════════

def check_merged(df: pd.DataFrame,
                 name: str = 'merged',
                 *,
                 expect_category_cols: Sequence[str] = ()) -> list[str]:
    """
    merge_data() guarantees: Last Response / Device Type / Username always
    exist (synthesised when no RMM provided). expect_category_cols documents
    the v0.12 re-downcast invariant — pass e.g. ('Device Type',) to assert
    the decategorise/re-downcast round trip completed. Category state is a
    warning, not an error: correctness survives object dtype, only memory
    suffers.
    """
    check_frame(
        df, name,
        required=('Name', 'Vulnerability Name', 'Vulnerability Score',
                  'Last Response', 'Device Type', 'Username'),
        numeric=('Vulnerability Score',),
    )
    issues = check_vocabulary(df, name, 'Device Type', allowed=DEVICE_TYPE_VOCAB)

    for col in expect_category_cols:
        if col in df.columns and not isinstance(df[col].dtype, pd.CategoricalDtype):
            issues.append(
                f"{name}: '{col}' expected category dtype after merge_data "
                f"re-downcast (v0.12) but is {df[col].dtype} — a conditional "
                f"write was likely added after the re-downcast step"
            )
    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 4 — scope frames, after the orchestrator builds them  (v0.3 guard)
# ══════════════════════════════════════════════════════════════════════════════

def check_scopes(merged_df: pd.DataFrame,
                 filtered_df: pd.DataFrame,
                 triage_df: pd.DataFrame,
                 threshold: float,
                 *,
                 health_filtered: Optional[pd.DataFrame] = None,
                 health_triage_df: Optional[pd.DataFrame] = None,
                 health_score_threshold: float = 7.0) -> list[str]:
    """
    The two-scope system introduced in v0.3, asserted directly:

      triage_df ⊆ filtered_df ⊆ merged_df          (row-index containment)
      filtered_df:  every row's score ≥ threshold
      triage_df:    zero 'Not Found in RMM' rows
      health scope: same shape at CVSS ≥ health_score_threshold, and the
                    health floor itself never drops below 7.0

    All violations here are STRUCTURAL — a scope leak silently corrupts
    every downstream sheet.
    """
    if not triage_df.index.isin(filtered_df.index).all():
        raise ContractError(
            "scopes: triage_df contains rows not present in filtered_df — "
            "triage must be a strict subset of the score-filtered scope (v0.3)"
        )
    if not filtered_df.index.isin(merged_df.index).all():
        raise ContractError(
            "scopes: filtered_df contains rows not present in merged_df — "
            "scope frames must be views/subsets of the merged frame"
        )

    if not filtered_df.empty:
        scores = pd.to_numeric(filtered_df['Vulnerability Score'], errors='coerce').fillna(0)
        n_below = int((scores < float(threshold)).sum())
        if n_below:
            raise ContractError(
                f"scopes: filtered_df has {n_below} row(s) below the CVSS "
                f"threshold {threshold} — wrong column compared? "
                f"(v0.3 'Score Threshold Default' failure class)"
            )

    if 'Last Response' in triage_df.columns and not triage_df.empty:
        n_nirm = int((triage_df['Last Response'] == 'Not Found in RMM').sum())
        if n_nirm:
            raise ContractError(
                f"scopes: triage_df contains {n_nirm} 'Not Found in RMM' "
                f"row(s) — not-in-RMM devices must never reach triage sheets (v0.3)"
            )

    issues: list[str] = []

    if health_triage_df is not None and health_filtered is not None:
        if float(health_score_threshold) < 7.0:
            raise ContractError(
                f"scopes: health_score_threshold={health_score_threshold} < 7.0 — "
                f"health scope floor is fixed at 7.0 so the score is comparable "
                f"across runs regardless of display threshold"
            )
        if not health_triage_df.index.isin(health_filtered.index).all():
            raise ContractError(
                "scopes: health_triage_df not a subset of health_filtered"
            )
        if 'Last Response' in health_triage_df.columns and not health_triage_df.empty:
            n = int((health_triage_df['Last Response'] == 'Not Found in RMM').sum())
            if n:
                raise ContractError(
                    f"scopes: health_triage_df contains {n} 'Not Found in RMM' row(s)"
                )
        if not health_filtered.empty:
            hs = pd.to_numeric(health_filtered['Vulnerability Score'], errors='coerce').fillna(0)
            n_below_h = int((hs < float(health_score_threshold)).sum())
            if n_below_h:
                raise ContractError(
                    f"scopes: health_filtered has {n_below_h} row(s) below the "
                    f"health floor {health_score_threshold}"
                )

    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 5 — patch match output, after process_patch_match()   (v0.4 guard)
# ══════════════════════════════════════════════════════════════════════════════

# Internal working columns process_patch_match promises to drop before return.
PATCH_INTERNAL_COLS = (
    '_cve_status_orig', '_patch_status',
    '_ck', '_sk', '_dk', '_mck', '_pk', '_pd', '_sr', '_kbs', '_pv', '_cves',
)

# The v0.4 collision signature: these are CVE threat statuses. If they show
# up in the patch install Status column, the merge kept the wrong side.
# This one is HARD — it is never legitimate.
CVE_STATUS_VALUES = ('RESOLVED', 'UNRESOLVED')


def check_patch_match(full_df: pd.DataFrame, name: str = 'patch_match.full') -> list[str]:
    """
    After process_patch_match(). Guards the v0.4 'Status Column Collision'
    bug from BOTH directions:

      1. structural — required output columns exist, internal underscore
         working columns were dropped, no duplicate labels;
      2. semantic (promoted to hard) — 'Status' must be the PATCH report's
         install status, so it must never contain RESOLVED/UNRESOLVED.
    """
    if full_df is None or full_df.empty:
        return []

    check_frame(
        full_df, name,
        required=('Status', 'Patch Match Result',
                  'Version Check Result', 'Patch Evidence Status'),
        forbidden=PATCH_INTERNAL_COLS,
    )

    # Raises ContractError on hit — the collision, reintroduced.
    check_vocabulary(full_df, name, 'Status',
                     forbidden_values=CVE_STATUS_VALUES, hard=True)

    # Evidence label sanity: the confirmed label must require the compliant
    # version check (classifier AND-chain), warn-only because label text may
    # evolve.
    issues: list[str] = []
    if {'Patch Evidence Status', 'Version Check Result'} <= set(full_df.columns):
        conf = full_df['Patch Evidence Status'].astype(str).str.contains(
            'Patch confirmed', case=False, na=False)
        if conf.any():
            vcr = full_df.loc[conf, 'Version Check Result'].astype(str).str.lower()
            n_bad = int((~vcr.str.contains('version compliant', na=False)).sum())
            if n_bad:
                issues.append(
                    f"{name}: {n_bad} row(s) marked 'Patch confirmed' without a "
                    f"'Version compliant' check result — classifier AND-chain "
                    f"may have regressed"
                )
    return issues


# ══════════════════════════════════════════════════════════════════════════════
# Orchestrator convenience — run a check, route warnings, never double-raise
# ══════════════════════════════════════════════════════════════════════════════

def run_check(fn, *args, warnings: Optional[list] = None, **kwargs) -> list[str]:
    """
    Execute a contract check, log + collect its semantic warnings into the
    orchestrator's warnings list, and let ContractError propagate (structural
    failures should stop the run with a clear message).

        run_check(check_cve_export, df_vuln, warnings=warnings)
    """
    issues = fn(*args, **kwargs)
    for msg in issues:
        log.warning("Contract: %s", msg)
        if warnings is not None:
            warnings.append(f"Data contract: {msg}")
    return issues