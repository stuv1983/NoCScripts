"""
sheet_names.py — reserved sheet names this tool uses for its own sheets,
as opposed to per-product sheets (named after whatever product the data
contains). Both orchestrator.py and data_pipeline.py must import
RESERVED_SHEET_NAMES from here rather than defining their own copy.

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