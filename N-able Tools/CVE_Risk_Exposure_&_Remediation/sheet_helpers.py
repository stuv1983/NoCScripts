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