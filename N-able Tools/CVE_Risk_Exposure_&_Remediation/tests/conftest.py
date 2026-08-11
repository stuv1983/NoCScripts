"""
tests/conftest.py — pytest configuration and shared fixtures.

Stubs out the `config` module so tests can import data_pipeline without
needing config.json, tkinter, or any other production dependency present.

Two rules matter here, because a conftest is imported before every test
module in its scope — including, in a whole-suite run, the repo-root test
files that do their own config stubbing:

  1. Register with sys.modules.setdefault(), never a forced assignment.
     data_pipeline binds `from config import PRODUCT_MAP` ONCE at import, so
     whichever stub lands first wins for the entire session. setdefault is the
     pattern every root-level test file already uses (see test_resolution.py's
     comment for the history) — this file was the sole outlier.

  2. Carry the same PRODUCT_MAP as those files. An empty map is not a neutral
     default: it makes _detect_product() return '' for everything, so the
     root-level tests that assert on product keys failed only in a full-suite
     run and passed in isolation — 15 tests' worth of silent, ordering-
     dependent breakage.
"""

import re
import sys
import types

# ── Stub the config module before data_pipeline imports it ───────────────────
_config_stub = types.ModuleType('config')
_config_stub.CVE_PATTERN        = re.compile(r'(CVE-\d{4}-\d{4,7})', re.IGNORECASE)
# Must match the map used by the repo-root test files — see rule 2 above.
_config_stub.PRODUCT_MAP        = [
    ('google chrome',   'chrome'),
    ('mozilla firefox', 'firefox'),
    ('microsoft edge',  'edge'),
]
_config_stub.FIXED_VERSION_RULES = {}
_config_stub.STATUS_RANK        = {
    'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
    'Pending': 3, 'Missing': 2, 'Failed': 1,
}
_config_stub.STATUS_LABEL       = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
_config_stub.INSTALLED_STATUSES = {'Installed', 'Reboot Required'}
_config_stub._CONFIG            = {}

# setdefault, not assignment — see rule 1 above.
sys.modules.setdefault('config', _config_stub)
