"""
sheet_helpers.py — small xlsxwriter-writing helpers shared across
sheet-builder modules (excel_builder.py, product_sheets.py, and others).

Author : Stu Villanti <s.villanti@kenstra.com>
"""


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
# The health score is scoped to CVSS ≥ HEALTH_SCORE_THRESHOLD regardless of
# the report's own threshold, so a report built at (say) 1.0 puts sub-7.0
# rows — with toggleable checkboxes — onto the sheets. Rows R5-R9 therefore
# carry the score criterion INSIDE their COUNTIFS so those rows can never
# feed the health score, while R1/R2 stay whole-sheet (they drive the
# Summary's Resolution Status table, which is scoped to the report itself).
# R3/R4 need no extra criterion: ≥ 9 already implies ≥ 7.
HEALTH_SCORE_THRESHOLD = 7.0

HS_SUBTOTAL_ROWS = {       # key → 0-indexed row  (R1..R9 in Excel terms)
    'res':        0,   # R1  ☑ rows (whole sheet — Resolution Status table)
    'unres':      1,   # R2  ☐ rows (whole sheet — Resolution Status table)
    'crit_res':   2,   # R3  ☑ rows with Vulnerability Score ≥ 9
    'crit_unres': 3,   # R4  ☐ rows with Vulnerability Score ≥ 9
    'exp_res':    4,   # R5  ☑ rows, Has Known Exploit = Yes, score ≥ 7
    'exp_unres':  5,   # R6  ☐ rows, Has Known Exploit = Yes, score ≥ 7
    'kev_unres':  6,   # R7  ☐ rows, CISA KEV = Yes, score ≥ 7 (live KEV cap)
    'hs_res':     7,   # R8  ☑ rows with score ≥ 7 (health resolution rate)
    'hs_unres':   8,   # R9  ☐ rows with score ≥ 7 (health resolution rate)
}
_HS_SUBTOTAL_LABELS = {
    'res':        'HS subtotal: resolved rows (☑)',
    'unres':      'HS subtotal: unresolved rows (☐)',
    'hs_res':     'HS subtotal: ☑ rows at CVSS ≥ 7',
    'hs_unres':   'HS subtotal: ☐ rows at CVSS ≥ 7',
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


def write_hs_subtotals(ws, workbook, col_names, counts: dict,
                       health_threshold: float = HEALTH_SCORE_THRESHOLD) -> None:
    """
    Write the nine health-score subtotal cells onto a product sheet.

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

    _ht = health_threshold
    def _formula(key):
        """Local formula string for one subtotal, or None if not computable."""
        if _c_res is None:
            return None
        mark = '☑' if key in ('res', 'crit_res', 'exp_res', 'hs_res') else '☐'
        base = f'${_c_res}:${_c_res},"{mark}"'
        _scope = (f',${_c_score}:${_c_score},">="&{_ht:g}'
                  if _c_score is not None else None)
        if key in ('res', 'unres'):
            return f'=COUNTIF({base})'
        if key in ('crit_res', 'crit_unres'):
            if _c_score is None:
                return None
            return f'=COUNTIFS({base},${_c_score}:${_c_score},">="&9)'
        if key in ('hs_res', 'hs_unres'):
            if _scope is None:
                return None
            return f'=COUNTIFS({base}{_scope})'
        if key == 'kev_unres':
            if _c_kev is None or _scope is None:
                return None
            return f'=COUNTIFS({base},${_c_kev}:${_c_kev},"Yes"{_scope})'
        if _c_exp is None or _scope is None:
            return None
        return f'=COUNTIFS({base},${_c_exp}:${_c_exp},"Yes"{_scope})'

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