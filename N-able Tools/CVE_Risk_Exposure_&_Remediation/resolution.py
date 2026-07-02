"""
resolution.py — single source of truth for "is this CVE/device row resolved?"

Used by build_product_sheets() (the ☑/☐ checkbox column) and
build_client_summary_sheet() (Resolution Status table) — both must call
this rather than reimplementing the logic locally.

Precedence (first match wins):
  1. Patch evidence    — (device, cve[, product]) in patch_resolved_pairs
  2. Raw export status — Threat Status / Status column == 'RESOLVED'
  3. Neither           — unresolved (☐)
Override: an approaching-stale device is always forced to ☐.

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
    approaching_stale_names: Optional[Set[str]] = None,
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
    approaching_stale_names = approaching_stale_names or set()

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

    # Override: approaching-stale devices always ☐
    if approaching_stale_names:
        name_list = df['Name'].tolist()
        res_bool = [
            False if name_list[i] in approaching_stale_names else res_bool[i]
            for i in range(len(res_bool))
        ]

    return res_bool


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
    approaching_stale_names: Optional[Set[str]] = None,
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
    approaching_stale_names = approaching_stale_names or set()
    patch_2d, patch_3d = split_patch_pairs(patch_resolved_pairs)

    if 'Base Product' not in df.columns:
        # No product column to group on at all — degrade to status-column-only,
        # still correctly aligned since there's no groupby involved.
        status_col = ('Threat Status' if 'Threat Status' in df.columns
                      else 'Status' if 'Status' in df.columns else None)
        is_approaching = (df['Name'].isin(approaching_stale_names)
                          if 'Name' in df.columns else pd.Series(False, index=df.index))
        is_resolved = (df[status_col].astype(str).str.strip().str.upper().eq('RESOLVED')
                       if status_col else pd.Series(False, index=df.index))
        return is_resolved & ~is_approaching

    flags: List[bool] = []
    flag_index: list = []
    for base_product, group in df.groupby('Base Product', sort=False):
        raw_pnames = (group['Affected Products'].dropna().astype(str).unique().tolist()
                      if 'Affected Products' in group.columns else [])
        sheet_pk = get_sheet_product_key(raw_pnames, base_product)
        flags.extend(compute_resolved_flags(group, sheet_pk, patch_2d, patch_3d,
                                            approaching_stale_names=approaching_stale_names))
        flag_index.extend(group.index.tolist())

    if len(flags) != len(df):
        # Defensive fallback for an unexpected groupby row-count mismatch —
        # should not happen, but fail toward "status column only" rather
        # than raise mid-report-generation.
        status_col = ('Threat Status' if 'Threat Status' in df.columns
                      else 'Status' if 'Status' in df.columns else None)
        is_approaching = (df['Name'].isin(approaching_stale_names)
                          if 'Name' in df.columns else pd.Series(False, index=df.index))
        is_resolved = (df[status_col].astype(str).str.strip().str.upper().eq('RESOLVED')
                       if status_col else pd.Series(False, index=df.index))
        return is_resolved & ~is_approaching

    # Build from (label, flag) pairs, then reindex to df's original row
    # order — pandas aligns by index LABEL here, not position, so this is
    # correct regardless of what order groupby visited rows in.
    return pd.Series(flags, index=flag_index).reindex(df.index)