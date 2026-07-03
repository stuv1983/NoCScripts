"""
test_trends_remediation_summary.py — unit tests for the Month-over-Month
Remediation Summary metrics added to data_pipeline.compute_trends().

Why this file exists: the first version of these metrics reused the
existing resolved_pair_count / new_pair_count (both restricted to products
present in BOTH reports) as the "cleared" / "new" figures, while the
"previous/current unresolved pairs" top-line figures used the full scope
(all products, including ones that appeared or disappeared entirely
between reports). Mixing those two scopes looks harmless but silently
breaks the arithmetic a reader would naturally do:
    Previous - Cleared + New should equal Current
On a real dataset this was off by 244 pairs. These tests pin down that the
four "Month-over-Month" figures are drawn from the same population, so
that identity holds exactly, and lock in the scope distinction from the
pre-existing resolved_pair_count so nobody "simplifies" this back to reusing
it.

Note: uses non-destructive config stubbing (setdefault, not an override +
forced re-import) and real product names already recognised by the
project's config.json (Google Chrome / Microsoft Edge / Mozilla Firefox),
specifically so this file doesn't need its own PRODUCT_MAP and can't poison
the shared data_pipeline/resolution module cache for whichever test file
happens to run after it — see test_resolution.py and test_patch_resolution.py
for the history of why that matters.

Run with: pytest test_trends_remediation_summary.py -v
"""
import os
import sys
import types

os.environ.setdefault('PYTEST_CURRENT_TEST', 'bootstrap')

_fake_config = types.ModuleType('config')
_fake_config.CVE_PATTERN = __import__('re').compile(r'(CVE-\d{4}-\d{4,7})', __import__('re').IGNORECASE)
_fake_config.PRODUCT_MAP = [
    ('google chrome', 'chrome'),
    ('microsoft edge', 'edge'),
    ('mozilla firefox', 'firefox'),
]
_fake_config.FIXED_VERSION_RULES = {}
_fake_config.STATUS_RANK = {}
_fake_config.STATUS_LABEL = {}
_fake_config.INSTALLED_STATUSES = {'Installed', 'Reboot Required'}
_fake_config._CONFIG = {}
sys.modules.setdefault('config', _fake_config)

import pandas as pd  # noqa: E402
from data_pipeline import compute_trends  # noqa: E402


def _df(rows):
    return pd.DataFrame(rows)


# Previous report: Chrome (2 CVEs on DEV1), Edge (1 CVE on DEV2) —
# Edge disappears entirely in the current report.
PREV_ROWS = [
    {'Name': 'DEV1', 'Vulnerability Name': 'CVE-2026-0001', 'Affected Products': 'Google Chrome',
     'Base Product': 'Google Chrome', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
    {'Name': 'DEV1', 'Vulnerability Name': 'CVE-2026-0002', 'Affected Products': 'Google Chrome',
     'Base Product': 'Google Chrome', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
    {'Name': 'DEV2', 'Vulnerability Name': 'CVE-2026-0003', 'Affected Products': 'Microsoft Edge',
     'Base Product': 'Microsoft Edge', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
]

# Current report: CVE-0001 persists, CVE-0002 is gone (cleared), CVE-0004 is
# new on Chrome, and Firefox appears for the first time (brand new product —
# its pair still counts as "new", and its device counts toward "current
# unresolved devices").
CUR_ROWS = [
    {'Name': 'DEV1', 'Vulnerability Name': 'CVE-2026-0001', 'Affected Products': 'Google Chrome',
     'Base Product': 'Google Chrome', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
    {'Name': 'DEV1', 'Vulnerability Name': 'CVE-2026-0004', 'Affected Products': 'Google Chrome',
     'Base Product': 'Google Chrome', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
    {'Name': 'DEV3', 'Vulnerability Name': 'CVE-2026-0005', 'Affected Products': 'Mozilla Firefox',
     'Base Product': 'Mozilla Firefox', 'Vulnerability Score': 9.0, 'Threat Status': 'UNRESOLVED'},
]


def _compute():
    prev_df = _df(PREV_ROWS)
    cur_df  = _df(CUR_ROWS)
    return compute_trends(cur_df, prev_df, threshold=9.0, prev_source_type='dashboard')


def test_previous_and_current_unresolved_pair_counts_use_full_scope():
    """Full scope: all 3 previous pairs and all 3 current pairs count,
    including Edge (previous-only) and Firefox (current-only)."""
    trend = _compute()
    m = trend['metrics']
    assert m['previous_unresolved_pair_count'] == 3
    assert m['current_unresolved_pair_count']  == 3


def test_cleared_count_includes_fully_removed_product():
    """
    Edge's pair (DEV2, CVE-0003) disappeared along with the whole
    product — it must still count as 'cleared' here, even though the
    pre-existing resolved_pair_count (common-product-restricted) would not
    count it, because Edge was never in the current report at all.
    """
    trend = _compute()
    m = trend['metrics']
    # Cleared: (DEV1,CVE-0002,Chrome) and (DEV2,CVE-0003,Edge) — 2 pairs.
    assert m['cleared_previous_unresolved_count'] == 2
    # The pre-existing, common-product-restricted metric only sees Chrome's
    # clearance (Edge is excluded from its population entirely) — this
    # locks in that the two metrics are legitimately different, so nobody
    # "simplifies" the new one back to reusing the old one.
    assert m['resolved_pair_count'] == 1
    assert m['cleared_previous_unresolved_count'] != m['resolved_pair_count']


def test_new_count_includes_brand_new_product():
    """(DEV1,CVE-0004,Chrome) and (DEV3,CVE-0005,Firefox) are new — 2 pairs."""
    trend = _compute()
    m = trend['metrics']
    assert m['new_unresolved_pair_count'] == 2


def test_remediation_summary_arithmetic_reconciles_exactly():
    """
    The property this whole set of metrics exists to guarantee: a reader
    doing Previous - Cleared + New by hand must land exactly on Current.
    This is the exact identity that was off by 244 on a real dataset before
    the fix (mixing full-scope totals with a common-product-restricted
    cleared count).
    """
    trend = _compute()
    m = trend['metrics']
    previous = m['previous_unresolved_pair_count']
    current  = m['current_unresolved_pair_count']
    cleared  = m['cleared_previous_unresolved_count']
    new      = m['new_unresolved_pair_count']
    assert previous - cleared + new == current


def test_cleared_percentage_matches_cleared_over_previous():
    trend = _compute()
    m = trend['metrics']
    expected = m['cleared_previous_unresolved_count'] / m['previous_unresolved_pair_count']
    assert m['cleared_previous_unresolved_pct'] == expected


def test_device_counts():
    """Previous: DEV1, DEV2 unresolved (2). Current: DEV1, DEV3 unresolved (2)."""
    trend = _compute()
    m = trend['metrics']
    assert m['previous_unresolved_device_count'] == 2
    assert m['current_unresolved_device_count']  == 2


def test_zero_previous_pairs_gives_zero_percent_not_a_crash():
    """No previous data at all — cleared percentage must degrade to 0.0,
    not raise a ZeroDivisionError."""
    empty_prev = _df(PREV_ROWS).iloc[0:0]   # same columns, zero rows
    cur_df = _df(CUR_ROWS)
    trend = compute_trends(cur_df, empty_prev, threshold=9.0, prev_source_type='dashboard')
    m = trend['metrics']
    assert m['previous_unresolved_pair_count'] == 0
    assert m['cleared_previous_unresolved_pct'] == 0.0