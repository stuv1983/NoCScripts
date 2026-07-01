"""
resolution.py — single source of truth for "is this CVE/device row resolved?"

Before this module existed, this determination was implemented twice,
independently, in excel_builder.py:
  - build_product_sheets()       — writes the per-row ☑/☐ checkbox column
  - build_client_summary_sheet() — re-derived the same flags to compute the
                                    Resolution Status table's numbers

Both copies applied the same three-source priority rule, hand-copied. That's
exactly the kind of duplication that lets a rule change land in one place and
silently miss the other — the checkbox column and the summary math could, in
principle, disagree even though they're describing the same thing.

Precedence (first match wins):
  1. Patch evidence     — (device, cve[, product]) present in patch_resolved_pairs
  2. Raw export status  — Threat Status / Status column == 'RESOLVED'
  3. Neither            — unresolved (☐)
Override: a device flagged "approaching stale" is always forced to ☐
regardless of the above — patch confirmation is treated as unreliable once
a device is close to going stale (last response near the cutoff).

What this module deliberately does NOT cover: trend-inferred resolution (a
device/CVE pair that was unresolved last report and is simply absent this
period). Those rows don't exist in the current dataframe at all — there's
nothing to flag — so that signal stays in data_pipeline.compute_trends()
and is combined separately, at the aggregate level, by the caller (see
build_client_summary_sheet's Resolution Status section).
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


def compute_resolved_series(
    df: pd.DataFrame,
    product_to_sheet: Optional[dict],
    patch_resolved_pairs: Optional[set],
    approaching_stale_names: Optional[Set[str]] = None,
) -> pd.Series:
    """
    Whole-dataframe version of compute_resolved_flags(): groups df by
    'Base Product', resolves each group against its own product scope, and
    returns a single bool Series correctly aligned to df's original index.

    This exists because of a real bug: df.groupby('Base Product') iterates
    rows in per-group order, not df's original row order. Naively building a
    flat list of flags across groups and then doing
    pd.Series(flat_list, index=df.index) assumes list position i belongs to
    df.index[i] — false whenever a product's rows aren't already contiguous
    in df. That silently attaches every flag to the WRONG row while the
    aggregate count (sum of Trues) still comes out right, which is exactly
    what let it hide in production: top-line totals matched, but anything
    that combined the resolved flag with another per-row filter (device
    type, exploit status, KEV-unresolved-by-CVE, health score components)
    was scrambled. It was silently harmless in some call sites purely by
    accident of how their input frame happened to be constructed (already
    grouped by product), and silently wrong in others (Top At-Risk Devices,
    Score Lift's KEV bonus) where the input frame wasn't.

    Do not reimplement this locally — call it, the same way
    compute_resolved_flags() itself must be called rather than
    reimplemented per call site.
    """
    approaching_stale_names = approaching_stale_names or set()
    p2s = product_to_sheet or {}
    patch_2d, patch_3d = split_patch_pairs(patch_resolved_pairs)

    if 'Base Product' not in df.columns or not p2s:
        # No product grouping available — degrade to status-column-only,
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
        if base_product not in p2s:
            flags.extend([False] * len(group))
            flag_index.extend(group.index.tolist())
            continue
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