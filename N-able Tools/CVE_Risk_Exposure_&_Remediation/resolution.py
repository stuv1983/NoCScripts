"""
resolution.py — single source of truth for "is this CVE/device row resolved?"

Used by build_product_sheets() (the ☑/☐ checkbox column) and
build_client_summary_sheet() (Resolution Status table) — both must call
this rather than reimplementing the logic locally.

Precedence (first match wins):
  1. Patch evidence    — (device, cve[, product]) in patch_resolved_pairs
  2. Raw export status — Threat Status / Status column == 'RESOLVED'
  3. Neither           — unresolved (☐)

Trend-inferred resolution (a device/CVE pair unresolved last report and
now absent) is handled separately in data_pipeline.compute_trends() —
those rows don't exist in the current dataframe, so there's nothing here
to flag.
"""
from __future__ import annotations

from typing import Dict, Optional, Set, Tuple, List

import pandas as pd

from data_pipeline import _detect_product, normalize_device_name, extract_cve_id


def get_sheet_product_key(raw_affected_products, base_product: str) -> str:
    """
    Resolve the product key a given product sheet groups on.

    Tries each distinct 'Affected Products' value seen in the group first
    (most specific), falling back to the Base Product label itself.
    """
    for rpn in (raw_affected_products or []):
        pk = _detect_product(str(rpn))
        if pk:
            return pk
    return _detect_product(str(base_product))


def split_patch_pairs(
    patch_resolved_pairs: Optional[Set[tuple]],
) -> Tuple[Set[Tuple[str, str]], Dict[str, Set[Tuple[str, str]]]]:
    """
    Split a set that may contain a mix of (device, cve) 2-tuples and
    (device, cve, product) 3-tuples into:
      - a flat 2-tuple set   (applies across all products)
      - a dict of product_key -> 2-tuple set (applies only within that product)

    Do this once per report run and pass the results into
    compute_resolved_flags() for every group/product — avoids re-scanning
    the full pair set for every row or every product sheet.
    """
    patch_2d: Set[Tuple[str, str]] = set()
    patch_3d: Dict[str, Set[Tuple[str, str]]] = {}
    for p in (patch_resolved_pairs or set()):
        if len(p) == 3:
            patch_3d.setdefault(p[2], set()).add((p[0], p[1]))
        else:
            patch_2d.add((p[0], p[1]))
    return patch_2d, patch_3d


def compute_resolved_flags(
    df: pd.DataFrame,
    sheet_product_key: str,
    patch_2d: Set[Tuple[str, str]],
    patch_3d: Dict[str, Set[Tuple[str, str]]],
) -> List[bool]:
    """
    Return a list[bool] aligned to df's row order — True means resolved (☑).

    df must have 'Name' and 'Vulnerability Name' columns. 'Threat Status' or
    'Status' is read if present; a missing status column simply means source
    2 contributes nothing (rows fall through to patch evidence only, or
    unresolved).

    This is THE function both build_product_sheets (checkbox column) and
    build_client_summary_sheet (Resolution Status table) must call — do not
    reimplement this logic at either call site.
    """
    nk_list = [normalize_device_name(str(n)) for n in df['Name']]
    ck_list = [extract_cve_id(str(v)) for v in df['Vulnerability Name']]

    # Source 1: patch evidence
    product_patch = patch_3d.get(sheet_product_key, set())
    if product_patch or patch_2d:
        res_bool = [
            (bool(product_patch) and (nk, ck) in product_patch)
            or (bool(patch_2d) and (nk, ck) in patch_2d)
            for nk, ck in zip(nk_list, ck_list)
        ]
    else:
        res_bool = [False] * len(df)

    # Source 2: raw export status
    status_col = ('Threat Status' if 'Threat Status' in df.columns
                  else 'Status' if 'Status' in df.columns else None)
    if status_col:
        status_resolved = (
            df[status_col].astype(str).str.strip().str.upper().eq('RESOLVED').tolist()
        )
        res_bool = [res_bool[i] or status_resolved[i] for i in range(len(res_bool))]

    return res_bool


def reconcile_patch_evidence(
    patch_resolved_pairs: Optional[Set[tuple]],
    unresolved_pairs_2d: Optional[Set[Tuple[str, str]]],
    trust_patch_evidence: bool = False,
    redundant_pairs_2d: Optional[Set[Tuple[str, str]]] = None,
) -> Tuple[Set[tuple], int, int]:
    """
    Decide what happens to patch evidence that the detections export
    contradicts, returning (surviving_pairs, override_count, redundant_dropped).

    redundant_pairs_2d — (device, cve) pairs the export ALREADY reports
    RESOLVED. Patch evidence for these changes nothing: compute_resolved_flags()
    source 2 reads the status column per-row and resolves them anyway. They are
    dropped first, because carrying them costs real time at fleet scale — a
    ~1M-row export can produce ~450k patch-confirmed pairs of which only ~12%
    are contested, and the set is re-split and re-scanned on every
    compute_resolved_series() call during workbook writing. This is the same
    reasoning the orchestrator already applies when it declines to inject raw
    RESOLVED rows on an export that has a status column.

    Only pass redundant_pairs_2d when the frames being rendered actually carry
    a status column; without one, source 2 contributes nothing and dropping
    these pairs would silently lose resolutions.

    A "contested" pair is one the patch report proved patched while the raw
    export still reports the same (device, cve) as UNRESOLVED. The comparison
    is deliberately 2-tuple: a product-string formatting difference must never
    cause a contested pair to be missed in either direction.

    Contested pairs are the ONLY pairs that matter. Patch evidence can only
    ever change an outcome for a row the scanner calls UNRESOLVED — where the
    export already says RESOLVED, compute_resolved_flags() source 2 resolves
    the row on its own. So the original unconditional strip (CHANGELOG v0.24,
    "UNRESOLVED always wins") removed precisely the pairs the patch report
    exists to contribute, leaving its net effect on resolution status at zero.

    trust_patch_evidence=False (default)
        Scanner wins; contested pairs are dropped. Unchanged v0.24 behaviour.
    trust_patch_evidence=True
        Patch evidence wins; contested pairs survive. N-able can lag a patch
        install by a full rescan cycle, so a stale UNRESOLVED is not proof the
        CVE is still open.

    Enabling the override is not "trust anything the patch tool said".
    patch_resolved_pairs only ever contains rows _vec_pes already qualified as
    installed AND version-compliant AND installed on/after the CVE's own
    publication/detection date, so a surviving pair carries affirmative
    evidence rather than a mere absence of contradiction — which is what keeps
    v0.24's false-positive concern (stale cache entries, mismatched patch
    records, product-name drift) answered.

    override_count counts distinct (device, cve) pairs, not 3-tuples, so a CVE
    spanning several product keys on one device is reported once.
    """
    pairs = set(patch_resolved_pairs or set())
    unresolved = unresolved_pairs_2d or set()

    # Drop pairs the export already resolves on its own — see above.
    redundant_dropped = 0
    if pairs and redundant_pairs_2d:
        _redundant = {p for p in pairs
                      if (p[0], p[1]) in redundant_pairs_2d
                      and (p[0], p[1]) not in unresolved}
        if _redundant:
            pairs -= _redundant
            redundant_dropped = len(_redundant)

    if not pairs or not unresolved:
        return pairs, 0, redundant_dropped

    contested = {p for p in pairs if (p[0], p[1]) in unresolved}
    if not contested:
        return pairs, 0, redundant_dropped

    if trust_patch_evidence:
        return pairs, len({(p[0], p[1]) for p in contested}), redundant_dropped
    return pairs - contested, 0, redundant_dropped


def dedup_per_base_product(df: pd.DataFrame) -> pd.DataFrame:
    """
    Deduplicate df by (Name, Vulnerability Name) within each 'Base Product'
    group, unconditionally — do not filter by whether a product also
    exists in some other scope (e.g. product_to_sheet). A product that
    exists only on stale or not-in-RMM devices must still be kept.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    if 'Base Product' in df.columns:
        frames = [g.drop_duplicates(subset=['Name', 'Vulnerability Name'])
                  for _, g in df.groupby('Base Product')]
        return pd.concat(frames, ignore_index=True) if frames else df.copy()
    return df.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()


def build_all_scope_frame(
    triage_dedup: pd.DataFrame,
    stale_excluded_df: Optional[pd.DataFrame] = None,
    not_in_rmm_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Build the frame the Summary sheet's "All" column (Key Metrics,
    Data Filtering Reconciliation) is computed from: active scope + every
    stale-excluded row + every not-in-RMM row, each deduplicated the same
    way triage_dedup itself is — see dedup_per_base_product() for why this
    must not filter by product_to_sheet membership.
    """
    stale_dedup = dedup_per_base_product(stale_excluded_df)
    nirm_dedup  = dedup_per_base_product(not_in_rmm_df)
    parts = [p for p in [triage_dedup, stale_dedup, nirm_dedup]
             if p is not None and not p.empty]
    if parts:
        return pd.concat(parts, ignore_index=True)
    return triage_dedup.copy() if triage_dedup is not None else pd.DataFrame()


def compute_resolved_series(
    df: pd.DataFrame,
    product_to_sheet: Optional[dict] = None,
    patch_resolved_pairs: Optional[set] = None,
) -> pd.Series:
    """
    Whole-dataframe version of compute_resolved_flags(): groups df by
    'Base Product', resolves each group, and returns a bool Series
    correctly aligned to df's original index (built from (label, flag)
    pairs, not by row position — groupby iterates rows in per-group order,
    not df's original order, so position-based reindexing silently
    misattaches flags to the wrong rows).

    This function resolves every row it's given — it does not filter or
    exclude rows based on product_to_sheet. A Base Product not present in
    product_to_sheet is still resolved via get_sheet_product_key() and
    whatever patch evidence / raw status it has; product_to_sheet is not
    used to gate inclusion. If you need a narrower scope (e.g. only
    products with an active product sheet), filter df before calling this
    — do not expect this function to do that filtering for you. (It used
    to: a product absent from product_to_sheet was silently forced
    unresolved, which broke the Health Score / Score Lift denominators
    whenever their broader CVSS scope included a product with no rows at
    the report's own threshold — see build_client_summary_sheet's Health
    Score section and product_sheets.py's Score Lift setup for how they now
    build their own scope frame with dedup_per_base_product() before
    calling this, instead of relying on this function to narrow it.)

    Do not reimplement this locally.
    """
    patch_2d, patch_3d = split_patch_pairs(patch_resolved_pairs)

    if 'Base Product' not in df.columns:
        # No product column to group on at all — degrade to status-column-only,
        # still correctly aligned since there's no groupby involved.
        status_col = ('Threat Status' if 'Threat Status' in df.columns
                      else 'Status' if 'Status' in df.columns else None)
        is_resolved = (df[status_col].astype(str).str.strip().str.upper().eq('RESOLVED')
                       if status_col else pd.Series(False, index=df.index))
        return is_resolved

    flags: List[bool] = []
    flag_index: list = []
    for base_product, group in df.groupby('Base Product', sort=False):
        raw_pnames = (group['Affected Products'].dropna().astype(str).unique().tolist()
                      if 'Affected Products' in group.columns else [])
        sheet_pk = get_sheet_product_key(raw_pnames, base_product)
        flags.extend(compute_resolved_flags(group, sheet_pk, patch_2d, patch_3d))
        flag_index.extend(group.index.tolist())

    if len(flags) != len(df):
        # Defensive fallback for an unexpected groupby row-count mismatch —
        # should not happen, but fail toward "status column only" rather
        # than raise mid-report-generation.
        status_col = ('Threat Status' if 'Threat Status' in df.columns
                      else 'Status' if 'Status' in df.columns else None)
        is_resolved = (df[status_col].astype(str).str.strip().str.upper().eq('RESOLVED')
                       if status_col else pd.Series(False, index=df.index))
        return is_resolved

    # Build from (label, flag) pairs, then reindex to df's original row
    # order — pandas aligns by index LABEL here, not position, so this is
    # correct regardless of what order groupby visited rows in.
    return pd.Series(flags, index=flag_index).reindex(df.index)