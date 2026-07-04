"""
excel_builder.py — xlsxwriter sheet-building functions.
No pandas data loading. No GUI. Receives DataFrames, writes sheets.

Author : Stu Villanti <s.villanti@kenstra.com>
"""

import logging
from datetime import datetime
from typing import Optional, Set, Tuple, Dict
import re
import pandas as pd
from config import CVE_PATTERN, INSTALLED_STATUSES
from data_pipeline import (
    normalize_device_name, extract_cve_id, get_base_product,
    clean_sheet_name, _drop_internal, parse_last_response, get_col_letter,
)

log = logging.getLogger(__name__)


# EXCEL SHEET BUILDERS
# ==============================================================================

# get_workbook_styles moved to formatting.py — re-exported here so existing
# `from excel_builder import get_workbook_styles` imports (e.g. orchestrator.py)
# keep working unchanged.
from formatting import get_workbook_styles, get_band_formats, build_legend_entries, COLORS  # noqa: E402,F401


# _write_cve_links / _write_nvd_links moved to sheet_helpers.py (shared by
# excel_builder.py and product_sheets.py without either importing the other).
from sheet_helpers import write_cve_links as _write_cve_links, write_nvd_links as _write_nvd_links  # noqa: E402,F401

# build_product_sheets, _build_patch_confirmed_sheet, and compute_score_lift
# moved to product_sheets.py — re-exported here so existing
# `from excel_builder import build_product_sheets` imports (orchestrator.py)
# keep working unchanged.
from product_sheets import build_product_sheets  # noqa: E402,F401

# build_trend_summary_sheet and build_trend_detail_sheets moved to
# trend_sheets.py — re-exported here so existing
# `from excel_builder import build_trend_summary_sheet, build_trend_detail_sheets`
# imports (orchestrator.py) keep working unchanged.
from trend_sheets import build_trend_summary_sheet, build_trend_detail_sheets  # noqa: E402,F401

# build_client_summary_sheet and compute_patching_health_score moved to
# summary_sheet.py — re-exported here so existing
# `from excel_builder import build_client_summary_sheet` imports
# (orchestrator.py) keep working unchanged.
from summary_sheet import build_client_summary_sheet, compute_patching_health_score  # noqa: E402,F401

# build_diagnostics_sheets, build_patch_resolved_sheet,
# build_products_not_tracked_sheet, build_patch_failure_sheet, and
# build_patch_sheets moved to patch_sheets.py — re-exported here so
# existing `from excel_builder import ...` imports (orchestrator.py)
# keep working unchanged.
from patch_sheets import (  # noqa: E402,F401
    build_diagnostics_sheets, build_patch_resolved_sheet,
    build_products_not_tracked_sheet, build_patch_failure_sheet,
    build_patch_sheets,
)

# build_all_detections_sheet, build_stale_excluded_sheet,
# build_stale_cves_sheet, build_device_report_sheet, and
# build_raw_data_sheet moved to device_sheets.py — re-exported here so
# existing `from excel_builder import ...` imports (orchestrator.py)
# keep working unchanged.
from device_sheets import (  # noqa: E402,F401
    build_all_detections_sheet, build_stale_excluded_sheet,
    build_stale_cves_sheet, build_device_report_sheet,
    build_raw_data_sheet,
)