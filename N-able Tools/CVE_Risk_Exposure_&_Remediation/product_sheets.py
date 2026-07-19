"""
product_sheets.py — per-product CVE detail sheets (the ☑/☐ triage tabs).

Split out of excel_builder.py, which had grown to ~3,000 lines mixing
formatting, business logic, and layout for every sheet type in the
workbook. This module owns exactly one thing: building the per-product
sheets (build_product_sheets) and the "fully patch-confirmed" variant of
them (_build_patch_confirmed_sheet), plus the Score Lift calculation that
only those sheets use.

Resolution logic (the ☑/☐ priority rules) is NOT here — it lives in
resolution.py and is called from here, not reimplemented here. See
resolution.py's docstring for why that separation matters.

Author : Stu Villanti <s.villanti@kenstra.com>
"""
from typing import Optional, Tuple, Dict

import pandas as pd

from data_pipeline import normalize_device_name, extract_cve_id
from formatting import get_workbook_styles, build_legend_entries
from resolution import (
    split_patch_pairs as _split_patch_pairs,
    get_sheet_product_key as _get_sheet_pk,
    compute_resolved_flags as _compute_flags,
    compute_resolved_series as _compute_resolved_series,
    dedup_per_base_product as _dedup_per_base_product,
)
from sheet_helpers import write_hs_subtotals as _write_hs_subtotals

import logging
log = logging.getLogger(__name__)


def compute_score_lift(
    row: dict,
    total_rows: int,
    crit_total: int,
    exp_total: int,
    kev_unresolved_rows_by_cve: 'dict[str, int]',
    persisting_cves: 'set[str]',
) -> float:
    """
    Estimate health-score points recoverable by resolving this single row.

    Returns 0.0 for already-resolved rows.  All denominators come from the
    fleet-wide health scope (pre-computed in build_product_sheets) so values
    are comparable across all product sheets.

    Score lift components
    ─────────────────────────────────────────────
    Base resolution      60 / total_rows       every unresolved row
    Critical coverage   +20 / crit_total       only CVSS ≥ 9 rows
    Known exploit       +20 / exp_total        only rows with Has Known Exploit = Yes
    KEV penalty         +1.0                   only if this clears the LAST unresolved
                                                instance of this KEV CVE type
    Persisting penalty  +0.5                   only for CVE IDs in the persisting set
    """
    if str(row.get('Resolved', '')).strip() == '☑':
        return 0.0

    cve_id = extract_cve_id(str(row.get('Vulnerability Name', '')))
    lift   = 0.0

    if total_rows:
        lift += 60.0 / total_rows

    score = pd.to_numeric(row.get('Vulnerability Score', 0), errors='coerce')
    if crit_total and pd.notna(score) and score >= 9.0:
        lift += 20.0 / crit_total

    if exp_total and str(row.get('Has Known Exploit', '')).strip().lower() in ('yes', 'y', 'true', '1'):
        lift += 20.0 / exp_total

    if str(row.get('CISA KEV', '')).strip().lower() in ('yes', 'y', 'true', '1'):
        if kev_unresolved_rows_by_cve.get(cve_id, 0) == 1:
            lift += 1.0

    if cve_id in persisting_cves:
        lift += 0.5

    return round(lift, 2)

def _hs_subtotal_counts(df: 'pd.DataFrame', all_resolved: bool = False,
                        health_threshold: float = 7.0) -> dict:
    """
    Generation-time values for the nine per-sheet health-score subtotals
    (cached results for the local formulas written by write_hs_subtotals,
    so openpyxl/pandas data_only readers see correct numbers).

    all_resolved=True is used by Patch Confirmed sheets, where the builder
    writes \u2611 into every row regardless of the DataFrame's own Resolved
    values.
    """
    n = len(df)
    _yes = {'yes', 'true', '1', 'y'}
    if all_resolved or 'Resolved' not in df.columns:
        res_mask = pd.Series([bool(all_resolved)] * n, index=df.index)
    else:
        res_mask = df['Resolved'].astype(str).str.strip() == '\u2611'
    unres_mask = ~res_mask

    if 'Vulnerability Score' in df.columns:
        _sc = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        crit_mask = _sc >= 9.0
        hs_mask   = _sc >= health_threshold
    else:
        crit_mask = pd.Series([False] * n, index=df.index)
        hs_mask   = pd.Series([False] * n, index=df.index)
    if 'Has Known Exploit' in df.columns:
        exp_mask = df['Has Known Exploit'].astype(str).str.strip().str.lower().isin(_yes)
    else:
        exp_mask = pd.Series([False] * n, index=df.index)
    if 'CISA KEV' in df.columns:
        kev_mask = df['CISA KEV'].astype(str).str.strip().str.lower().isin(_yes)
    else:
        kev_mask = pd.Series([False] * n, index=df.index)

    return {
        'res':        int(res_mask.sum()),
        'unres':      int(unres_mask.sum()),
        'crit_res':   int((crit_mask & res_mask).sum()),
        'crit_unres': int((crit_mask & unres_mask).sum()),
        # exploit / KEV / hs counts are scoped to the health threshold so
        # sub-scope rows (present when the report threshold is lower) can
        # never feed the health score — mirrors the COUNTIFS criteria in
        # sheet_helpers.write_hs_subtotals.
        'exp_res':    int((exp_mask & hs_mask & res_mask).sum()),
        'exp_unres':  int((exp_mask & hs_mask & unres_mask).sum()),
        'kev_unres':  int((kev_mask & hs_mask & unres_mask).sum()),
        'hs_res':     int((hs_mask & res_mask).sum()),
        'hs_unres':   int((hs_mask & unres_mask).sum()),
    }


def _build_patch_confirmed_sheet(writer, sheet_name: str, product: str,
                                  out_df: 'pd.DataFrame', col_names: list) -> None:
    """
    Write a lightweight "Patch Confirmed" sheet for a product where every
    detected CVE is already resolved (all rows are checkmark).

    Layout is intentionally minimal:
      - Same column order as a full triage sheet (col A = Resolved = checkmark)
        so Summary COUNTIF / COUNTIFS formulas count correctly.
      - Green header banner with a clear "all resolved" message.
      - Data rows written read-only (no dropdown validation — nothing to change).
      - No Score Lift column, no unresolved colouring, no legend.

    This keeps the workbook truthful and formula-safe without the bulk of a
    full triage sheet.
    """
    wb  = writer.book
    ws  = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # Banner
    banner_fmt = wb.add_format({
        'bold': True, 'font_size': 13,
        'bg_color': '#375623', 'font_color': 'white',
        'border': 1, 'align': 'left', 'valign': 'vcenter',
    })
    note_fmt = wb.add_format({
        'italic': True, 'font_size': 9,
        'font_color': '#375623', 'bg_color': '#E2EFDA', 'border': 1,
    })
    n_cols = min(len(col_names) - 1, 9)
    ws.merge_range(0, 0, 0, n_cols,
                   f'\u2705  {product}  \u2014  All CVEs Patch Confirmed', banner_fmt)
    # Navigation: internal link back to the Summary sheet, just past the banner.
    _back_fmt = wb.add_format({'bold': True, 'font_color': '#0563C1',
                               'underline': True, 'align': 'center', 'valign': 'vcenter'})
    ws.write_url(0, n_cols + 1, "internal:'Summary'!A1",
                 _back_fmt, string='\u2190 Summary')
    ws.set_column(n_cols + 1, n_cols + 1, 12)
    ws.set_row(0, 28)
    ws.merge_range(1, 0, 1, n_cols,
                   'Every detected CVE for this product has patch evidence or is marked RESOLVED '
                   'in N-able.  No action required.  Rows are read-only (\u2611 pre-filled).  '
                   'This sheet is included so Summary health-score formulas count these '
                   'resolutions correctly.',
                   note_fmt)
    ws.set_row(1, 28)

    # Column header row (row index 2)
    hdr_fmt = wb.add_format({
        'bold': True, 'bg_color': '#375623', 'font_color': 'white', 'border': 1,
    })
    for ci, col in enumerate(col_names):
        ws.write(2, ci, col, hdr_fmt)

    # Data rows (start at row index 3)
    grn_fmt = wb.add_format({'bg_color': '#E2EFDA', 'border': 1})
    grn_chk = wb.add_format({'bg_color': '#E2EFDA', 'border': 1,
                              'bold': True, 'align': 'center'})

    resolved_idx = col_names.index('Resolved') if 'Resolved' in col_names else None
    vuln_idx     = col_names.index('Vulnerability Name') if 'Vulnerability Name' in col_names else None

    for ri, row_tuple in enumerate(out_df.itertuples(index=False, name=None), start=3):
        for ci, val in enumerate(row_tuple):
            if ci == resolved_idx:
                ws.write(ri, ci, '\u2611', grn_chk)
            elif ci == vuln_idx:
                # Plain text — CVE hyperlinks removed for write speed and to
                # stay clear of xlsxwriter's 65,530-URL-per-sheet ceiling.
                val_str = str(val) if val is not None else ''
                display = val_str[:255] if len(val_str) <= 255 else val_str[:252] + '...'
                ws.write(ri, ci, display, grn_fmt)
            else:
                ws.write(ri, ci, val if val is not None else '', grn_fmt)

    n_data = len(out_df)

    # Hidden subtotal block (cols Q/R) — same fixed location as triage
    # sheets; built from THIS sheet's column order. all_resolved=True.
    _write_hs_subtotals(ws, wb, col_names,
                        _hs_subtotal_counts(out_df, all_resolved=True))

    # Column widths (mirrors full triage sheet)
    _widths = {
        'Resolved':                  10,
        'Score Lift':                10,
        'Vulnerability Name':        25,
        'Name':                      25,
        'Device Type':               15,
        'Vulnerability Severity':    20,
        'Vulnerability Score':        8,
        'Risk Severity Index':       18,
        'Has Known Exploit':         18,
        'CISA KEV':                  12,
        'Last Response':             22,
        'Days Since Last Response':  22,
        'Affected Products':         30,
        'Baseline Compliance':       22,
    }
    for ci, col in enumerate(col_names):
        ws.set_column(ci, ci, _widths.get(col, 16))

    # Footer note
    foot_fmt = wb.add_format({'italic': True, 'font_color': '#595959', 'font_size': 8})
    ws.write(n_data + 4, 0,
             f'\u2139  {n_data} row(s) \u2014 all patch confirmed.  '
             'Sheet is read-only; use the full triage sheets for '
             'products with remaining unresolved CVEs.',
             foot_fmt)

    log.debug("Patch Confirmed sheet written for '%s': %d row(s)", product, n_data)


def build_product_sheets(writer, triage_df, product_to_sheet,
                          patch_resolved_pairs=None,
                          patch_gap_pairs: Optional[Dict[Tuple[str, str], str]] = None,
                          health_triage_df: 'Optional[pd.DataFrame]' = None,
                          trend_data: Optional[dict] = None,
                          include_health_score: bool = False):
    if patch_resolved_pairs is None:
        patch_resolved_pairs = set()
    if patch_gap_pairs is None:
        patch_gap_pairs = {}

    # Split patch_resolved_pairs by product key once so each sheet checks
    # only its own subset. Shared with the Summary sheet via resolution.py.
    _patch_2d, _patch_3d = _split_patch_pairs(patch_resolved_pairs)

    # Score Lift is a Health Score companion — skip all of it (including
    # the per-row column below) when Health Score is off.
    if include_health_score:
        # Fleet-wide denominators, computed once so every sheet divides by the same totals.
        _sl_scope = health_triage_df if (health_triage_df is not None and not health_triage_df.empty) else triage_df

        # Include every Base Product — the health scope (CVSS >= 7.0) can
        # contain products with no rows at the report's own threshold.
        _sl_dedup = _dedup_per_base_product(_sl_scope)

        _sl_sc_col    = 'Vulnerability Score' if 'Vulnerability Score' in _sl_dedup.columns else None
        _sl_total     = len(_sl_dedup)
        _sl_crit_total = int((pd.to_numeric(_sl_dedup[_sl_sc_col], errors='coerce') >= 9.0).sum()) if _sl_sc_col else 0
        _sl_exp_col   = 'Has Known Exploit' if 'Has Known Exploit' in _sl_dedup.columns else None
        _sl_exp_total = int(_sl_dedup[_sl_exp_col].astype(str).str.strip().str.lower().isin(['yes','y','true','1']).sum()) if _sl_exp_col else 0

        # Unresolved KEV row counts per CVE ID for penalty-recovery lift.
        # Uses compute_resolved_series() — same three-source logic as the
        # ☑/☐ column and the Summary's Resolution Status table.
        _sl_is_res = _compute_resolved_series(_sl_dedup, product_to_sheet, patch_resolved_pairs)
        _sl_is_unr = ~_sl_is_res

        _kev_unres_by_cve: dict = {}
        if 'CISA KEV' in _sl_dedup.columns and 'Vulnerability Name' in _sl_dedup.columns:
            _kev_mask = _sl_dedup['CISA KEV'].astype(str).str.strip().str.lower().isin(['yes','y','true','1'])
            for _cve_raw in _sl_dedup.loc[_kev_mask & _sl_is_unr, 'Vulnerability Name']:
                _cid = extract_cve_id(str(_cve_raw))
                _kev_unres_by_cve[_cid] = _kev_unres_by_cve.get(_cid, 0) + 1

        _persisting_cves: set = set()
        if trend_data is not None:
            _persisting_cves = trend_data.get('persisting_cve_ids', set()) or set()

    cols_order = ['Resolved'] + (['Score Lift'] if include_health_score else []) + [
                  'Vulnerability Name', 'Name', 'Device Type',
                  'Vulnerability Severity', 'Vulnerability Score', 'Risk Severity Index',
                  'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response', 'Affected Products',
                  'Baseline Compliance']

    def _chromium_sort_key(p: str) -> str:
        """Sort Chrome first, Edge immediately after (both are Chromium-based).
        All other products sort by their own name alphabetically."""
        pl = str(p).lower()
        if 'chrome' in pl and 'edge' not in pl:
            return 'google chrome\x00' + pl   # Chrome: sorts at its natural 'G' position
        if 'edge' in pl:
            return 'google chrome\x01' + pl   # Edge: immediately after Chrome
        return pl

    # Build groups once via groupby, then iterate in Chromium-aware order.
    # Fully-confirmed products are deferred so their sheets appear after all
    # active (partially-unresolved) product sheets and before the stale sheets.
    _product_groups = {p: g for p, g in triage_df.groupby('Base Product')}
    _deferred_confirmed: list = []   # (sheet_name, product, out_df, final_cols)

    for product in sorted(_product_groups.keys(), key=_chromium_sort_key):
        group = _product_groups[product]
        sheet_name = product_to_sheet[product]
        group = group.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
        # Primary sort applied after Score Lift is computed (below)

        _raw_pnames = group['Affected Products'].dropna().astype(str).unique().tolist()
        _sheet_pk = _get_sheet_pk(_raw_pnames, product)

        # Normalised keys for the sparse set_row loop below.
        _nk_list = [normalize_device_name(str(n)) for n in group['Name']]
        _ck_list = [extract_cve_id(str(v))        for v in group['Vulnerability Name']]

        # Resolved column — shared three-source logic, see resolution.py.
        _res_bool = _compute_flags(group, _sheet_pk, _patch_2d, _patch_3d)

        _res_list = ['☑' if x else '☐' for x in _res_bool]

        # Fully-patched products deferred so tab order is:
        # active → confirmed → stale/NIRM.
        if all(v == '☑' for v in _res_list):
            group.insert(0, 'Resolved', _res_list)
            final_cols = [c for c in cols_order if c in group.columns]
            _out = group[final_cols]
            _deferred_confirmed.append((sheet_name, product, _out, final_cols))
            continue

        group.insert(0, 'Resolved', _res_list)

        # ── Score Lift ─────────────────────────────────────────────────────────────────
        if include_health_score:
            # Only the five columns compute_score_lift reads —
            # to_dict('records') boxed every cell and dominated at scale.
            _sl_cols = ['Resolved', 'Vulnerability Name', 'Vulnerability Score',
                        'Has Known Exploit', 'CISA KEV']
            _sl_data = {c: (group[c].tolist() if c in group.columns
                            else [''] * len(group)) for c in _sl_cols}
            _sl_list = [
                compute_score_lift(
                    dict(zip(_sl_cols, vals)),
                    _sl_total, _sl_crit_total, _sl_exp_total,
                    _kev_unres_by_cve, _persisting_cves)
                for vals in zip(*(_sl_data[c] for c in _sl_cols))
            ]
            group.insert(1, 'Score Lift', _sl_list)
            _sort_cols = ['Score Lift', 'Vulnerability Score', '_Sort_Time', 'Name']
            _sort_asc  = [False, False, False, True]
        else:
            _sort_cols = ['Vulnerability Score', '_Sort_Time', 'Name']
            _sort_asc  = [False, False, True]

        group = group.sort_values(by=_sort_cols, ascending=_sort_asc)

        final_cols = [c for c in cols_order if c in group.columns]
        _out = group[final_cols]

        # Direct write_row (faster than to_excel); register in writer.sheets
        # so later set_column / conditional_format / autofilter calls work.
        wb_ = writer.book
        ws  = wb_.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws
        ws.write_row(0, 0, final_cols)
        # Navigation: one internal link back to the Summary sheet, placed just
        # past the last header so it sits outside the autofilter range.
        _back_fmt = wb_.add_format({'bold': True, 'font_color': '#0563C1',
                                    'underline': True, 'align': 'center'})
        ws.write_url(0, len(final_cols), "internal:'Summary'!A1",
                     _back_fmt, string='\u2190 Summary')
        ws.set_column(len(final_cols), len(final_cols), 12)
        for _ri, _row in enumerate(_out.itertuples(index=False, name=None), start=1):
            ws.write_row(_ri, 0, _row)

        ws.autofilter(0, 0, len(group), len(final_cols) - 1)

        styles_           = get_workbook_styles(wb_)
        patch_res_fmt     = styles_['row_blue']
        exploit_fmt       = wb_.add_format({'bg_color': '#FFE0CC'})
        coverage_fmt      = styles_['row_amber']
        unmanaged_fmt     = styles_['row_red']
        mismatch_fmt      = styles_['row_pink']
        installing_fmt    = styles_['row_teal']

        _GAP_FMTS = {
            'coverage_gap':        coverage_fmt,
            'unmanaged_app':       unmanaged_fmt,
            'detection_mismatch':  mismatch_fmt,
            'patch_installing':    installing_fmt,
        }

        cl = final_cols
        if 'Resolved' in cl:
            ri = cl.index('Resolved')
            ws.data_validation(1, ri, len(group), ri, {'validate': 'list', 'source': ['☐', '☑']})
            ws.set_column(ri, ri, 10)

        _last = len(group)   # last data row (1-indexed)
        _TRUE_VALS = {'yes', 'true', '1', 'y'}

        unresolved_fmt = wb_.add_format({'bg_color': '#FFCCCC', 'font_color': '#8B0000'})  # coral red — unresolved, clearly distinct from blue

        # Row colouring via conditional_format. Rules added first win in
        # Excel. Ranges start at Excel row 2, so formulas reference $A2 —
        # $A1 shifts every row's colour off by one. The Score column is
        # excluded from row-level ranges: the always-true ☑/☐ rules would
        # otherwise override the score-band colours on every row.
        _vs_idx = cl.index('Vulnerability Score') if 'Vulnerability Score' in cl else None

        def _row_cf(cf_dict):
            """Apply a row-level conditional format across all columns except
            the Vulnerability Score column (if present)."""
            if _vs_idx is None:
                ws.conditional_format(1, 0, _last, len(cl) - 1, cf_dict)
                return
            if _vs_idx > 0:
                ws.conditional_format(1, 0, _last, _vs_idx - 1, cf_dict)
            if _vs_idx < len(cl) - 1:
                ws.conditional_format(1, _vs_idx + 1, _last, len(cl) - 1, cf_dict)

        # Priority 1: Resolved (☑ in col A) → blue.
        _row_cf({
            'type':     'formula',
            'criteria': '=$A2="☑"',
            'format':   patch_res_fmt,
        })
        # Priority 2: Unresolved (☐ in col A) → light red. Makes unresolved rows
        # immediately visible against the blue resolved rows.
        _row_cf({
            'type':     'formula',
            'criteria': '=$A2="☐"',
            'format':   unresolved_fmt,
        })
        # Priority 3: Known exploit → darker orange (overrides unresolved red).
        _exp_col = 'Has Known Exploit'
        if _exp_col in cl:
            _ec = chr(ord('A') + cl.index(_exp_col))
            _row_cf({
                'type':     'formula',
                'criteria': f'=OR(${_ec}2=TRUE,UPPER(TEXT(${_ec}2,"@"))="YES")',
                'format':   exploit_fmt,
            })

        # set_row only for patch-gap rows; skip when none exist (common case).
        if patch_gap_pairs:
            _exp_arr     = (group[_exp_col].astype(str).str.strip().str.lower().tolist()
                            if _exp_col in group.columns else [''] * len(group))
            for _ri, (_nk, _ck, _rv, _ev) in enumerate(
                zip(_nk_list, _ck_list, _res_list, _exp_arr), start=1
            ):
                # Priority 1: resolved rows → handled by conditional_format
                if _rv == '☑':
                    continue
                # Priority 2: known exploit → handled by conditional_format
                if _ev in _TRUE_VALS:
                    continue
                # Priority 4: patch gap types
                _gap = patch_gap_pairs.get((_nk, _ck))
                if _gap and _gap in _GAP_FMTS:
                    ws.set_row(_ri, None, _GAP_FMTS[_gap])

        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25)
        if 'Name'               in cl: ws.set_column(cl.index('Name'),               cl.index('Name'),               25)
        if 'Device Type'        in cl: ws.set_column(cl.index('Device Type'),        cl.index('Device Type'),        15)
        if 'Baseline Compliance' in cl: ws.set_column(cl.index('Baseline Compliance'), cl.index('Baseline Compliance'), 22)
        if _vs_idx is not None:
            ws.set_column(_vs_idx, _vs_idx, 8)
            # CVSS score bands (crit/high/med/low). Added after row-level CFs
            # so row colour still wins on resolved/exploit rows.
            _crit_fmt = wb_.add_format({'bg_color': '#FF0000', 'font_color': 'white',
                                        'num_format': '0.0', 'align': 'center'})
            _high_fmt = wb_.add_format({'bg_color': '#FFC000', 'font_color': 'black',
                                        'num_format': '0.0', 'align': 'center'})
            _med_fmt  = wb_.add_format({'bg_color': '#FFFF00', 'font_color': 'black',
                                        'num_format': '0.0', 'align': 'center'})
            _low_fmt  = wb_.add_format({'bg_color': '#92D050', 'font_color': 'black',
                                        'num_format': '0.0', 'align': 'center'})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': '>=', 'value': 9.0, 'format': _crit_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 7.0, 'maximum': 8.9,
                'format': _high_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 4.0, 'maximum': 6.9,
                'format': _med_fmt})
            ws.conditional_format(1, _vs_idx, _last, _vs_idx, {
                'type': 'cell', 'criteria': 'between', 'minimum': 0.1, 'maximum': 3.9,
                'format': _low_fmt})

        # Hidden subtotal block (cols Q/R) — one cell per sheet for the
        # Summary's live formulas (avoids Excel's 8,192-char formula limit).
        # Written unconditionally: Resolution Status table is always live.
        _write_hs_subtotals(ws, wb_, final_cols, _hs_subtotal_counts(group))

        legend_row = len(group) + 3
        l_title = wb_.add_format({'bold': True, 'font_size': 9, 'bg_color': '#F2F2F2', 'border': 1})
        l_cell  = wb_.add_format({'font_size': 9, 'border': 1})

        legend_entries = build_legend_entries()
        ws.write(legend_row, 0, 'Legend', l_title)
        for i, (colour, label, desc) in enumerate(legend_entries, start=1):
            fmt = wb_.add_format({'bg_color': colour, 'font_size': 9, 'border': 1})
            ws.write(legend_row + i, 0, f'  ({label})', fmt)
            ws.write(legend_row + i, 1, desc, l_cell)
            ws.set_row(legend_row + i, None, fmt)

    # ── Flush deferred confirmed sheets (after all active product sheets) ──────
    # These appear at the end of the product-sheet group, immediately before the
    # stale/NIRM sheets that the orchestrator writes next.
    for _sn, _prod, _out_df, _fcols in _deferred_confirmed:
        _build_patch_confirmed_sheet(writer, _sn, _prod, _out_df, _fcols)
    if _deferred_confirmed:
        log.debug(
            "Patch Confirmed sheets written: %d product(s) — %s",
            len(_deferred_confirmed),
            ', '.join(p for _, p, __, ___ in _deferred_confirmed),
        )