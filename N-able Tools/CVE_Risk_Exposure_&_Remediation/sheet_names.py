"""
sheet_names.py — single source of truth for sheet names this tool manages
itself (as opposed to per-product sheets, which are named after whatever
product the data contains).

Previously this set was defined twice, independently:
  - orchestrator.py's `reserved`   — used to stop a product being named the
    same as one of the tool's own sheets when allocating product_to_sheet.
  - data_pipeline.py's `_RESERVED` — used by load_previous_report() to skip
    the tool's own sheets when scanning a previous dashboard for per-row
    ☑/☐ checkbox state.

The two lists had already drifted (one had 'resolved since previous report',
the other didn't; one had 'overview' / 'client summary', the other didn't).
That's a silent-failure risk: a product sheet could collide with a tool
sheet name, or a scan for checkbox state could wrongly include/exclude a
sheet. Both call sites must import RESERVED_SHEET_NAMES from here instead of
defining their own copy.

All names are lowercase — callers should lowercase whatever they're
comparing against before checking membership.
"""

RESERVED_SHEET_NAMES = {
    'summary', 'client summary', 'overview',
    'trend summary',
    'all detections', 'raw data',
    'stale excluded devices', 'cves on stale devices',
    'new this month', 'new device-cve pairs', 'new cve types',
    'resolved', 'resolved since previous report', 'persisting cves',
    'patch match overview', 'patch match full data', 'patch report (full)',
    'patch confirmed', 'resolved (patch confirmed)',
    'device inventory',
}