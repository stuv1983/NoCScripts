"""
tests/conftest.py — pytest configuration and shared fixtures.

Stubs out the `config` module so tests can import data_pipeline without
needing config.json, tkinter, or any other production dependency present.
"""

import re
import sys
import types

# ── Stub the config module before data_pipeline imports it ───────────────────
_config_stub = types.ModuleType('config')
_config_stub.CVE_PATTERN        = re.compile(r'(CVE-\d{4}-\d{4,7})', re.IGNORECASE)
_config_stub.PRODUCT_MAP        = []           # no product detection in unit tests
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

sys.modules['config'] = _config_stub
