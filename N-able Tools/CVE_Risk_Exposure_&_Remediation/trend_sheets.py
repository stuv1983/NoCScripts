"""
trend_sheets.py — month-over-month trend detail sheets (New This Month,
Persisting CVEs).

Split out of excel_builder.py as part of breaking that file into one module
per sheet category.

v0.33 — the 'Trend Summary' sheet was removed: everything it showed is now
covered by the Summary sheet's two month-over-month sections (Remediation
Summary and Patching Progress), so it duplicated the Summary and pushed the
product sheets one tab further away. 'Resolved Since Previous Report' was
removed too: resolution is INFERRED there (a row merely absent from the
current scope), which double-reports work already tracked precisely by the
Patch Confirmed sheets and the Resolution Status table, and confused readers
by never being the full list. NVD columns and CVE hyperlinks were dropped for
write speed.

Author : Stu Villanti <s.villanti@kenstra.com>
"""
from formatting import COLORS


def build_trend_detail_sheets(writer, workbook, trend, sheets_subset=None):
    new_bg  = workbook.add_format({'bg_color': COLORS['PEACH_BG']})
    per_bg  = workbook.add_format({'bg_color': COLORS['AMBER_BG']})

    detail_cols = ['Name', 'Username', 'Device Type', 'Vulnerability Name', 'Vulnerability Score',
                   'Vulnerability Severity', 'Affected Products',
                   'Has Known Exploit', 'CISA KEV', 'Last Response', 'Days Since Last Response']

    all_sheets = [
        ('New This Month',  trend.get('new_cve_types_df', trend['new_df']), new_bg,
         'New CVE types not seen in the previous report — investigate and prioritise.'),
        ('Persisting CVEs', trend['persisting_df'], per_bg,
         'CVEs carried over from the previous report — still unresolved.'),
    ]

    for sheet_name, df, row_fmt, note in all_sheets:
        if sheets_subset and sheet_name not in sheets_subset:
            continue
        if df.empty:
            ws = workbook.add_worksheet(sheet_name)
            ws.write(0, 0, f'No records — {note}')
            continue

        df = df.copy()
        present = [c for c in detail_cols if c in df.columns]
        df = df[present]

        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

        cl = df.columns.tolist()
        if 'Name'               in cl: ws.set_column(cl.index('Name'),               cl.index('Name'),               25)
        if 'Username'           in cl: ws.set_column(cl.index('Username'),           cl.index('Username'),           22)
        if 'Device Type'        in cl: ws.set_column(cl.index('Device Type'),        cl.index('Device Type'),        15)
        if 'Affected Products'  in cl: ws.set_column(cl.index('Affected Products'),  cl.index('Affected Products'),  30)
        if 'Vulnerability Name' in cl:
            vn_idx = cl.index('Vulnerability Name')
            ws.set_column(vn_idx, vn_idx, 25)

        ws.conditional_format(1, 0, len(df), len(cl) - 1,
                               {'type': 'no_blanks', 'format': row_fmt})
        ws.write(len(df) + 2, 0, f'ℹ  {note}')