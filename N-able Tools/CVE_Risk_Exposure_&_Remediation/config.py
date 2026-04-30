"""
config.py — loads config.json and exposes all shared constants.

config.json is the single source of truth for product mappings and version rules.
No product data lives in this file or any other Python file.
"""

import json
import logging
import re
from pathlib import Path
from typing import Optional, Set, Tuple

log = logging.getLogger(__name__)

# Pre-compiled once here; imported everywhere else.
CVE_PATTERN = re.compile(r'(CVE-\d{4}-\d{4,7})', re.IGNORECASE)

# Minimal built-in fallback used ONLY when config.json is absent in test/CI
# environments. Never used in production — ship config.json with the script.
_FALLBACK_PRODUCT_MAP = [
    ["mozilla firefox", "firefox"],
    ["google chrome",   "chrome"],
    ["microsoft edge",  "edge"],
]


def _load_config(strict: bool = True) -> dict:
    """
    Load config.json from the same directory as this file.

    strict=True  (default, production):
        Raises FileNotFoundError if config.json is missing.
        Raises RuntimeError if the file cannot be parsed.

    strict=False (test / CI / headless):
        Returns a minimal built-in fallback if config.json is absent.
        Logs a warning so the caller knows they're on fallback data.
        Still raises RuntimeError if the file EXISTS but is malformed
        (malformed config is always an error — missing config may be intentional).
    """
    config_path = Path(__file__).parent / 'config.json'

    if not config_path.exists():
        if strict:
            raise FileNotFoundError(
                "config.json not found.\n\n"
                f"Expected location: {config_path}\n\n"
                "config.json must be in the same folder as this script.\n"
                "Download or restore it from the project release package."
            )
        else:
            log.warning(
                "config.json not found at %s — using minimal built-in fallback. "
                "This is only acceptable in test/CI environments.",
                config_path,
            )
            return {'product_map': _FALLBACK_PRODUCT_MAP, 'fixed_version_rules': {}}

    try:
        with open(config_path, 'r', encoding='utf-8') as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            raise ValueError("config.json root must be a JSON object.")
        return data
    except (json.JSONDecodeError, ValueError, OSError) as e:
        raise RuntimeError(
            f"config.json could not be loaded: {e}\n\n"
            "Fix the file or restore it from the project release package."
        ) from e


# Use strict=False when running under pytest (PYTEST_CURRENT_TEST is set by pytest)
import os as _os
_strict = 'PYTEST_CURRENT_TEST' not in _os.environ
_CONFIG = _load_config(strict=_strict)

_raw_pm = _CONFIG.get('product_map', [])
if not _raw_pm:
    raise RuntimeError(
        "config.json 'product_map' is empty or missing.\n"
        "Restore config.json from the project release package."
    )

# (lowercase_substring, canonical_name) — matched top-to-bottom in _detect_product
PRODUCT_MAP: list = [(str(k).lower(), str(v).lower()) for k, v in _raw_pm]

# { canonical_name: { CVE-ID: minimum_fixed_version } }
FIXED_VERSION_RULES: dict = _CONFIG.get('fixed_version_rules', {})

# Patch-status constants — used by both data_pipeline and excel_builder
STATUS_RANK: dict = {
    'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
    'Pending': 3,   'Missing': 2,          'Failed': 1,
}
STATUS_LABEL: dict = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
INSTALLED_STATUSES: Set[str] = {'Installed', 'Reboot Required'}
