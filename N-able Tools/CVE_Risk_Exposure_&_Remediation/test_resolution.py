"""
test_resolution.py — unit tests for resolution.py, the single source of
truth for "is this CVE/device row resolved?"

Why this file exists: the Resolution Status table on the Summary sheet once
showed 0% Resolved on a real customer report even though the customer had
been patching all along — because the per-row ☑/☐ logic in
build_product_sheets() and the aggregate math in build_client_summary_sheet()
were two independently-maintained copies of the same rule, and neither one
alone made it obvious that "no --patch file + an export with no RESOLVED
rows" == "Resolved will always be 0". That's exactly the kind of thing a
test should catch before a report goes out, not a person staring at a
workbook wondering if the tool is broken.

Run with:
    pytest test_resolution.py -v

Author : Stu Villanti <s.villanti@kenstra.com>
"""

import os
import sys
import types

os.environ.setdefault('PYTEST_CURRENT_TEST', 'bootstrap')

# ---------------------------------------------------------------------------
# Config stubbing: intentionally non-destructive.
#
# If test_patch_resolution.py (or another test module) already stubbed
# sys.modules['config'] and imported data_pipeline before this file runs,
# reuse that same module object rather than forcing a fresh import. Forcing
# a fresh import here would leave a *different* data_pipeline module object
# in sys.modules than the one other test files' already-bound functions
# were closed over at their own import time — their later
# `import data_pipeline as _dp` (done inside test functions, at run time,
# not at collection time) would then silently mutate a dict on the wrong
# module object. None of the tests below depend on FIXED_VERSION_RULES
# content, so it's always safe to just reuse whatever is already loaded.
# ---------------------------------------------------------------------------

# Same config-stubbing pattern as test_patch_resolution.py, so this file can
# run standalone without a real config.json on disk.
_fake_config = types.ModuleType('config')
_fake_config.CVE_PATTERN = __import__('re').compile(r'(CVE-\d{4}-\d{4,7})', __import__('re').IGNORECASE)
_fake_config.PRODUCT_MAP = [
    ('google chrome', 'chrome'),
    ('mozilla firefox', 'firefox'),
    ('microsoft edge', 'edge'),
]
_fake_config.FIXED_VERSION_RULES = {}
_fake_config.STATUS_RANK = {
    'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
    'Pending': 3,   'Missing': 2,          'Failed': 1,
}
_fake_config.STATUS_LABEL = {
    'Installed':       'Matched - installed',
    'Reboot Required': 'Matched - reboot required',
    'Installing':      'Matched - installing',
    'Pending':         'Matched - pending',
    'Missing':         'Matched - missing',
    'Failed':          'Matched - failed',
}
_fake_config.INSTALLED_STATUSES = {'Installed', 'Reboot Required'}
_fake_config._CONFIG = {}
sys.modules.setdefault('config', _fake_config)

import pandas as pd  # noqa: E402
from resolution import (  # noqa: E402
    compute_resolved_flags,
    split_patch_pairs,
    get_sheet_product_key,
    compute_resolved_series,
)


def _df(rows):
    return pd.DataFrame(rows)


# ── split_patch_pairs ─────────────────────────────────────────────────────────

def test_split_patch_pairs_handles_mixed_2d_and_3d():
    pairs = {
        ('CVLT001', 'CVE-2026-0001'),                    # 2-tuple: applies to any product
        ('CVLT002', 'CVE-2026-0002', 'chrome'),           # 3-tuple: applies only to chrome
    }
    two_d, three_d = split_patch_pairs(pairs)
    assert two_d == {('CVLT001', 'CVE-2026-0001')}
    assert three_d == {'chrome': {('CVLT002', 'CVE-2026-0002')}}


def test_split_patch_pairs_empty_input():
    two_d, three_d = split_patch_pairs(None)
    assert two_d == set()
    assert three_d == {}


# ── compute_resolved_flags: priority rules ────────────────────────────────────

def test_patch_evidence_marks_resolved_even_without_status_column():
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001'}])
    two_d, three_d = split_patch_pairs({('CVLT001', 'CVE-2026-0001')})
    flags = compute_resolved_flags(df, 'chrome', two_d, three_d)
    assert flags == [True]


def test_status_resolved_marks_resolved_with_no_patch_evidence():
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
               'Threat Status': 'RESOLVED'}])
    flags = compute_resolved_flags(df, 'chrome', set(), {})
    assert flags == [True]


def test_all_unresolved_when_neither_source_present():
    """
    The exact scenario that caused the original bug: no --patch file
    (empty patch pairs) and every row's status is UNRESOLVED. This MUST
    produce all-False — that's correct behaviour, not a bug — but it's
    worth pinning down explicitly so nobody 'fixes' it into a false
    positive later.
    """
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002', 'Threat Status': 'UNRESOLVED'},
    ])
    flags = compute_resolved_flags(df, 'chrome', set(), {})
    assert flags == [False, False]


def test_approaching_stale_overrides_patch_evidence():
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001'}])
    two_d, three_d = split_patch_pairs({('CVLT001', 'CVE-2026-0001')})
    flags = compute_resolved_flags(df, 'chrome', two_d, three_d,
                                    approaching_stale_names={'CVLT001'})
    assert flags == [False]


def test_approaching_stale_overrides_status_resolved():
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
               'Threat Status': 'RESOLVED'}])
    flags = compute_resolved_flags(df, 'chrome', set(), {},
                                    approaching_stale_names={'CVLT001'})
    assert flags == [False]


def test_status_column_named_status_instead_of_threat_status():
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Status': 'RESOLVED'}])
    flags = compute_resolved_flags(df, 'chrome', set(), {})
    assert flags == [True]


def test_3tuple_patch_pair_only_applies_within_its_product():
    """A 3-tuple pair scoped to 'chrome' must not resolve the same
    device/CVE pair being evaluated under a different product key."""
    two_d, three_d = split_patch_pairs({('CVLT001', 'CVE-2026-0001', 'chrome')})
    df = _df([{'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001'}])
    assert compute_resolved_flags(df, 'chrome', two_d, three_d) == [True]
    assert compute_resolved_flags(df, 'edge',   two_d, three_d) == [False]


# ── get_sheet_product_key ─────────────────────────────────────────────────────

def test_get_sheet_product_key_prefers_affected_products_over_base_product():
    key = get_sheet_product_key(['Google Chrome'], 'Some Other Base Label')
    assert key == 'chrome'


def test_get_sheet_product_key_falls_back_to_base_product():
    key = get_sheet_product_key([], 'Google Chrome')
    assert key == 'chrome'


# ── Consistency guard ─────────────────────────────────────────────────────────
# This is the test that matters most: build_product_sheets (the ☑/☐ column)
# and build_client_summary_sheet (the Resolution Status table) must derive
# their resolved/unresolved counts from calling compute_resolved_flags() —
# not from two separately-maintained implementations of the same rule. This
# test doesn't inspect excel_builder.py directly (that would require a full
# workbook build), but it locks down that calling the shared function twice,
# the way both call sites do, is deterministic and side-effect free — so a
# future refactor that makes either call site stop using resolution.py would
# be the only way to reintroduce the original class of bug.
def test_same_inputs_same_outputs_across_repeated_calls():
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Threat Status': 'RESOLVED'},
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002', 'Threat Status': 'UNRESOLVED'},
    ])
    two_d, three_d = split_patch_pairs({('CVLT002', 'CVE-2026-0002', 'chrome')})
    first  = compute_resolved_flags(df, 'chrome', two_d, three_d)
    second = compute_resolved_flags(df, 'chrome', two_d, three_d)
    assert first == second == [True, True]


# ── compute_resolved_series: row-misalignment regression ──────────────────────
# A real production bug: building a flat list of per-group flags across a
# groupby('Base Product') and then doing pd.Series(flat_list, index=df.index)
# assumes list position i corresponds to df.index[i] — false whenever a
# product's rows aren't already contiguous in df. That silently attaches
# every flag to the WRONG row while sums/aggregates still come out right
# (which is exactly why it went unnoticed: top-line totals matched, but any
# breakdown correlating "resolved" with another per-row attribute did not).
# These tests use INTERLEAVED products specifically, since a bug like this
# hides completely if a test's input happens to already be grouped by
# product — that coincidence is what let this ship in the first place.

def test_compute_resolved_series_with_interleaved_products():
    """
    Rows deliberately interleaved: chrome, edge, chrome, edge — not grouped.
    Row 0 (chrome, resolved) and row 2 (chrome, unresolved) must each keep
    their own correct flag, not swap with each other or with the edge rows.
    """
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'RESOLVED'},
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002',
         'Base Product': 'Microsoft Edge', 'Affected Products': 'Microsoft Edge',
         'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVLT003', 'Vulnerability Name': 'CVE-2026-0003',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVLT004', 'Vulnerability Name': 'CVE-2026-0004',
         'Base Product': 'Microsoft Edge', 'Affected Products': 'Microsoft Edge',
         'Threat Status': 'RESOLVED'},
    ])
    product_to_sheet = {'Google Chrome': 'Google Chrome', 'Microsoft Edge': 'Microsoft Edge'}
    result = compute_resolved_series(df, product_to_sheet, patch_resolved_pairs=None)

    assert result.tolist() == [True, False, False, True]
    # Also check by index label directly, not just position — a Series that
    # happens to be right positionally but wrong by label would still be a
    # correctness bug the moment a caller does df.loc[result] instead of
    # relying on row order.
    assert result.loc[0] == True
    assert result.loc[1] == False
    assert result.loc[2] == False
    assert result.loc[3] == True


def test_compute_resolved_series_matches_across_many_products():
    """
    Broader check: five products, rows shuffled so no product's rows are
    contiguous, half resolved half not. Cross-checks against a
    straightforward per-row loop (the 'obviously correct but slow' version)
    rather than hand-computing expected output.
    """
    import random
    random.seed(7)
    rows = []
    products = ['Google Chrome', 'Microsoft Edge', '7-Zip', 'Adobe Reader', 'Zoom']
    for i in range(40):
        product = products[i % len(products)]
        rows.append({
            'Name': f'DEV{i:03d}',
            'Vulnerability Name': f'CVE-2026-{i:04d}',
            'Base Product': product,
            'Affected Products': product,
            'Threat Status': 'RESOLVED' if i % 3 == 0 else 'UNRESOLVED',
        })
    random.shuffle(rows)  # break any accidental contiguity by product
    df = _df(rows)
    product_to_sheet = {p: p for p in products}

    result = compute_resolved_series(df, product_to_sheet, patch_resolved_pairs=None)

    for idx, row in df.iterrows():
        expected = row['Threat Status'] == 'RESOLVED'
        assert result.loc[idx] == expected, (
            f"Row {idx} ({row['Name']}, {row['Vulnerability Name']}): "
            f"expected {expected}, got {result.loc[idx]}"
        )


def test_compute_resolved_series_respects_product_scoped_patch_pairs():
    """A 3-tuple patch pair scoped to one product must not resolve the same
    device/CVE combo appearing under a different product."""
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome'},
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Microsoft Edge', 'Affected Products': 'Microsoft Edge'},
    ])
    product_to_sheet = {'Google Chrome': 'Google Chrome', 'Microsoft Edge': 'Microsoft Edge'}
    pairs = {('CVLT001', 'CVE-2026-0001', 'chrome')}
    result = compute_resolved_series(df, product_to_sheet, patch_resolved_pairs=pairs)
    assert result.loc[0] == True    # Chrome row: covered by the scoped pair
    assert result.loc[1] == False   # Edge row: same device+CVE, different product — must stay unresolved