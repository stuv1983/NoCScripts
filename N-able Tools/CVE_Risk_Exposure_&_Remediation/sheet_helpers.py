"""
sheet_helpers.py — small xlsxwriter-writing helpers shared across
sheet-builder modules (excel_builder.py, product_sheets.py, and others).

Author : Stu Villanti <s.villanti@kenstra.com>
"""
from config import CVE_PATTERN


def write_cve_links(ws, vuln_name_series, col_idx, link_fmt):
    """Write clickable cve.org links for each CVE in a column."""
    for row_i, val in enumerate(vuln_name_series, start=1):
        val_str = str(val)
        m = CVE_PATTERN.search(val_str)
        if m:
            cve_id  = m.group(1).upper()
            display = val_str[:255] if len(val_str) <= 255 else val_str[:252] + '...'
            ws.write_url(row_i, col_idx,
                         f'https://www.cve.org/CVERecord?id={cve_id}',
                         link_fmt, string=display)


def write_nvd_links(ws, vuln_name_series, col_idx, link_fmt):
    """
    Write plain 'NVD ↗' text (not a real hyperlink) for each CVE row.

    xlsxwriter has a hard limit of 65,530 URLs per worksheet; large product
    sheets (Chrome/Edge can exceed 40k rows) blow past that with real
    hyperlinks, so this intentionally writes styled text instead.
    """
    for row_i, val in enumerate(vuln_name_series, start=1):
        if CVE_PATTERN.search(str(val)):
            ws.write(row_i, col_idx, 'NVD ↗', link_fmt)

# ── Patching Health Score per-sheet subtotals ────────────────────────────────
# Each product sheet totals its own ☑/☐ counts ONCE, in a hidden block at a
# fixed location (labels in col Q, values in col R, rows 1-7).  The Summary
# sheet's live health-score / Resolution Status formulas then reference one
# cell per sheet ('Sheet'!$R$1 + 'Sheet'!$R$1 + ...) instead of embedding one
# or two full COUNTIFS expressions per sheet.  With many product sheets the
# old approach produced formulas of 20k+ characters — far above Excel's
# 8,192-character stored-formula limit — which forced the Summary score back
# to static values.
#
# A second benefit: each sheet's subtotal formulas are built from that sheet's
# OWN column layout.  Patch Confirmed sheets have no Score Lift column, so
# their Vulnerability Score / Has Known Exploit columns sit one to the left of
# a full triage sheet's — the old Summary-side COUNTIFS hard-coded G:G / I:I
# for every sheet and silently tested the wrong columns on confirmed sheets.
HS_SUBTOTAL_LBL_COL = 16   # col Q (hidden) — human-readable labels
HS_SUBTOTAL_VAL_COL = 17   # col R (hidden) — the subtotal values/formulas
HS_SUBTOTAL_ROWS = {       # key → 0-indexed row  (R1..R7 in Excel terms)
    'res':        0,   # R1  ☑ rows
    'unres':      1,   # R2  ☐ rows
    'crit_res':   2,   # R3  ☑ rows with Vulnerability Score ≥ 9
    'crit_unres': 3,   # R4  ☐ rows with Vulnerability Score ≥ 9
    'exp_res':    4,   # R5  ☑ rows with Has Known Exploit = Yes
    'exp_unres':  5,   # R6  ☐ rows with Has Known Exploit = Yes
    'kev_unres':  6,   # R7  ☐ rows with CISA KEV = Yes (reserved for live
                       #     KEV penalty work; penalties are static today)
}
_HS_SUBTOTAL_LABELS = {
    'res':        'HS subtotal: resolved rows (☑)',
    'unres':      'HS subtotal: unresolved rows (☐)',
    'crit_res':   'HS subtotal: resolved CVSS ≥ 9 rows',
    'crit_unres': 'HS subtotal: unresolved CVSS ≥ 9 rows',
    'exp_res':    'HS subtotal: resolved known-exploit rows',
    'exp_unres':  'HS subtotal: unresolved known-exploit rows',
    'kev_unres':  'HS subtotal: unresolved CISA KEV rows',
}


def hs_subtotal_ref(sheet_name: str, key: str) -> str:
    """Cross-sheet reference to one subtotal cell, e.g. 'Google Chrome'!$R$3."""
    row = HS_SUBTOTAL_ROWS[key] + 1                       # Excel 1-indexed row
    col = chr(ord('A') + HS_SUBTOTAL_VAL_COL)             # 'R'
    safe = str(sheet_name).replace("'", "''")             # escape apostrophes
    return f"'{safe}'!${col}${row}"


def write_hs_subtotals(ws, workbook, col_names, counts: dict) -> None:
    """
    Write the seven health-score subtotal cells onto a product sheet.

    Each value cell holds a LOCAL formula over this sheet's own columns
    (so it stays live when ☑/☐ are toggled) with the generation-time count
    as the cached result (so data_only readers see correct values).  If a
    column needed by a formula doesn't exist on this sheet, the static
    count is written instead — a cell is ALWAYS written so cross-sheet
    references from the Summary sheet never point at an empty cell.

    Columns Q and R are hidden; labels are kept for anyone unhiding them.
    """
    def _col_letter(name):
        try:
            i = col_names.index(name)
        except ValueError:
            return None
        return chr(ord('A') + i) if i < 26 else None

    _c_res   = _col_letter('Resolved')
    _c_score = _col_letter('Vulnerability Score')
    _c_exp   = _col_letter('Has Known Exploit')
    _c_kev   = _col_letter('CISA KEV')

    _hidden_fmt = workbook.add_format({'font_color': '#BFBFBF', 'font_size': 8})

    def _formula(key):
        """Local formula string for one subtotal, or None if not computable."""
        if _c_res is None:
            return None
        mark = '☑' if key in ('res', 'crit_res', 'exp_res') else '☐'
        base = f'${_c_res}:${_c_res},"{mark}"'
        if key in ('res', 'unres'):
            return f'=COUNTIF({base})'
        if key in ('crit_res', 'crit_unres'):
            if _c_score is None:
                return None
            return f'=COUNTIFS({base},${_c_score}:${_c_score},">="&9)'
        if key == 'kev_unres':
            if _c_kev is None:
                return None
            return f'=COUNTIFS({base},${_c_kev}:${_c_kev},"Yes")'
        if _c_exp is None:
            return None
        return f'=COUNTIFS({base},${_c_exp}:${_c_exp},"Yes")'

    for key, r in HS_SUBTOTAL_ROWS.items():
        static = int(counts.get(key, 0))
        ws.write(r, HS_SUBTOTAL_LBL_COL, _HS_SUBTOTAL_LABELS[key], _hidden_fmt)
        f = _formula(key)
        if f is not None:
            ws.write_formula(r, HS_SUBTOTAL_VAL_COL, f, _hidden_fmt, static)
        else:
            ws.write_number(r, HS_SUBTOTAL_VAL_COL, static, _hidden_fmt)

    ws.set_column(HS_SUBTOTAL_LBL_COL, HS_SUBTOTAL_VAL_COL, None, None,
                  {'hidden': True})