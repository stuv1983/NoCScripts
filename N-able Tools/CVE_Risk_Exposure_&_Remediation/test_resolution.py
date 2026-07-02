"""
test_resolution.py — unit tests for resolution.py.

Run with: pytest test_resolution.py -v

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
    dedup_per_base_product,
    build_all_scope_frame,
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


# ── compute_resolved_series does NOT deduplicate — callers must ───────────────
# A real bug: build_product_sheets deduplicates each product's rows by
# (Name, Vulnerability Name) BEFORE computing ☑/☐ (keep='first'), but the
# Summary sheet's Top At-Risk Devices table was calling compute_resolved_series
# on the raw, un-deduplicated triage_df. When a device had multiple raw rows
# for the same CVE with mixed resolved status (duplicate scan entries, or the
# same CVE surfacing under two Affected Products variants), the two tables
# could legitimately disagree — the product sheet showed one final verdict,
# Top At-Risk counted the CVE as unresolved if ANY duplicate instance was.
# Fixed by pointing Top At-Risk at triage_dedup (the same per-Base-Product
# drop_duplicates(['Name','Vulnerability Name']) frame the product sheets and
# Resolution Status table already use) instead of raw triage_df.
#
# compute_resolved_series() itself is correct either way — it just resolves
# whatever rows it's given. The contract callers must honor is: dedupe your
# input the same way build_product_sheets does, BEFORE calling this, or your
# counts can disagree with what the reader sees on the actual product sheet.
def test_compute_resolved_series_result_depends_on_caller_deduplication():
    """
    Demonstrates why the dedup-timing bug was possible: the same raw data,
    deduplicated vs not, gives different "how many CVEs are unresolved for
    this device" answers when duplicate rows disagree on status. This isn't
    a bug in compute_resolved_series — it's documentation of the contract
    that let the bug happen when a caller skipped the dedup step.
    """
    raw_rows = [
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'RESOLVED'},
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',   # duplicate CVE, conflicting status
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'UNRESOLVED'},
    ]
    product_to_sheet = {'Google Chrome': 'Google Chrome'}

    # Without dedup: nunique() over the unresolved subset still counts this
    # CVE once, because one of the two duplicate rows is unresolved.
    df_raw = _df(raw_rows)
    resolved_raw = compute_resolved_series(df_raw, product_to_sheet, patch_resolved_pairs=None)
    unresolved_cve_count_raw = df_raw.loc[~resolved_raw, 'Vulnerability Name'].nunique()
    assert unresolved_cve_count_raw == 1

    # Deduplicated first (keep='first', matching build_product_sheets): only
    # the first row survives, and it's RESOLVED — so this CVE does not count
    # as unresolved. This is what the product sheet itself would show.
    df_dedup = df_raw.drop_duplicates(subset=['Name', 'Vulnerability Name']).copy()
    resolved_dedup = compute_resolved_series(df_dedup, product_to_sheet, patch_resolved_pairs=None)
    unresolved_cve_count_dedup = df_dedup.loc[~resolved_dedup, 'Vulnerability Name'].nunique()
    assert unresolved_cve_count_dedup == 0

    # The two disagree — which is exactly why every consumer of
    # compute_resolved_series that reports a count meant to match a product
    # sheet MUST dedupe with the same rule first.
    assert unresolved_cve_count_raw != unresolved_cve_count_dedup


# ── build_all_scope_frame / dedup_per_base_product: stale-only-product bug ────
# Real bug: the "All" scope used to be built by deduplicating stale/not-in-RMM
# rows PER Base Product, but only for Base Products that were ALSO a key in
# product_to_sheet (the active scope). A product existing exclusively on
# stale or not-in-RMM devices — never on any active device — would silently
# vanish from "Unique CVE Types (All)" / "Total detection rows (All)" even
# though those exact rows are written, unfiltered, to their own detail sheet.
# Confirmed on a real dataset: 8 stale-only products, 470 rows, 468 unique
# CVEs, silently dropped. These tests lock in that dedup_per_base_product()
# and build_all_scope_frame() include EVERY Base Product, regardless of
# whether it also appears in some other scope.

def test_dedup_per_base_product_includes_products_with_no_active_sheet():
    """A product that only exists in this frame (e.g. only ever seen on
    stale devices) must still be deduplicated and kept — not dropped just
    because some OTHER scope (like product_to_sheet) doesn't know about it."""
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Base Product': 'MySQL Server'},
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Base Product': 'MySQL Server'},  # dup
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002', 'Base Product': 'VMware Tools'},
    ])
    result = dedup_per_base_product(df)
    assert len(result) == 2   # duplicate MySQL row collapsed, VMware row kept
    assert set(result['Base Product']) == {'MySQL Server', 'VMware Tools'}


def test_dedup_per_base_product_handles_empty_and_none():
    assert dedup_per_base_product(None).empty
    assert dedup_per_base_product(_df([])).empty


def test_build_all_scope_frame_includes_stale_only_product():
    """
    The exact scenario that caused the bug: 'MySQL Server' only appears on
    a stale device, never on any active one, so it's not in triage_dedup
    at all. The All-scope frame must still include it.
    """
    triage_dedup = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Base Product': 'Google Chrome'},
    ])
    stale_df = _df([
        {'Name': 'CVLT099', 'Vulnerability Name': 'CVE-2026-9999', 'Base Product': 'MySQL Server'},
    ])
    all_scope = build_all_scope_frame(triage_dedup, stale_excluded_df=stale_df, not_in_rmm_df=None)

    assert len(all_scope) == 2
    assert all_scope['Vulnerability Name'].nunique() == 2
    assert 'CVE-2026-9999' in all_scope['Vulnerability Name'].values


def test_build_all_scope_frame_includes_not_in_rmm_only_product():
    """Same scenario, but for a product only seen on a not-in-RMM device."""
    triage_dedup = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Base Product': 'Google Chrome'},
    ])
    nirm_df = _df([
        {'Name': 'CVFANTOM', 'Vulnerability Name': 'CVE-2026-8888', 'Base Product': 'PuTTY'},
    ])
    all_scope = build_all_scope_frame(triage_dedup, stale_excluded_df=None, not_in_rmm_df=nirm_df)

    assert len(all_scope) == 2
    assert 'CVE-2026-8888' in all_scope['Vulnerability Name'].values


def test_build_all_scope_frame_with_no_exclusions_equals_triage_dedup():
    triage_dedup = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001', 'Base Product': 'Google Chrome'},
    ])
    all_scope = build_all_scope_frame(triage_dedup, stale_excluded_df=None, not_in_rmm_df=None)
    assert len(all_scope) == len(triage_dedup)


# ── compute_resolved_series: health-score scope gap ────────────────────────────
# Real bug: health_triage_df (used for Health Score / Score Lift) is built
# from a broader CVSS threshold (≥7.0) than the report's own threshold
# (e.g. ≥9.0), so it can contain a product with rows in the 7.0-8.9 gap but
# none at the report's own threshold — meaning that product is never a key
# in product_to_sheet (which is built from the narrower triage_df). Callers
# used to pre-filter their scope frame to `bp in product_to_sheet`, or
# compute_resolved_series() itself forced such rows to False — either way,
# a product legitimately in-scope for the Health Score was silently
# dropped from it. Fixed by (a) building scope frames with
# dedup_per_base_product() (no product_to_sheet filtering) and (b) making
# compute_resolved_series() resolve every group it's given rather than
# gating on product_to_sheet membership — inclusion/exclusion is now the
# caller's job, done before calling this function, not this function's job.

def test_compute_resolved_series_resolves_products_not_in_product_to_sheet():
    """
    'MySQL Server' is NOT a key in product_to_sheet (simulating a product
    that only has rows below the report's own threshold, so it never got a
    product sheet) — but it must still be resolved correctly via its own
    raw status, not forced unresolved just because product_to_sheet doesn't
    know about it.
    """
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002',
         'Base Product': 'MySQL Server', 'Affected Products': 'MySQL Server',
         'Threat Status': 'RESOLVED'},
    ])
    product_to_sheet = {'Google Chrome': 'Google Chrome'}   # MySQL Server deliberately absent

    result = compute_resolved_series(df, product_to_sheet, patch_resolved_pairs=None)

    assert result.loc[0] == False   # Chrome row: genuinely unresolved
    assert result.loc[1] == True    # MySQL row: genuinely resolved — must NOT be forced False


def test_compute_resolved_series_resolves_products_not_in_product_to_sheet_via_patch_evidence():
    """Same scope gap, but the out-of-scope product's resolution comes from
    patch evidence rather than raw status — must still apply correctly."""
    df = _df([
        {'Name': 'CVFANTOM', 'Vulnerability Name': 'CVE-2026-9999',
         'Base Product': 'VMware Tools', 'Affected Products': 'VMware Tools'},
    ])
    product_to_sheet = {'Google Chrome': 'Google Chrome'}   # VMware Tools absent
    pairs = {('CVFANTOM', 'CVE-2026-9999')}   # 2-tuple: applies regardless of product

    result = compute_resolved_series(df, product_to_sheet, patch_resolved_pairs=pairs)
    assert result.loc[0] == True


def test_compute_resolved_series_with_empty_product_to_sheet_still_groups_by_base_product():
    """
    Even with product_to_sheet=None/empty entirely, a 'Base Product' column
    is enough to group and resolve correctly — this must NOT degrade to the
    raw-status-only fallback just because product_to_sheet is missing.
    """
    df = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'RESOLVED'},
        {'Name': 'CVLT002', 'Vulnerability Name': 'CVE-2026-0002',
         'Base Product': 'Microsoft Edge', 'Affected Products': 'Microsoft Edge',
         'Threat Status': 'UNRESOLVED'},
    ])
    result = compute_resolved_series(df, product_to_sheet=None, patch_resolved_pairs=None)
    assert result.loc[0] == True
    assert result.loc[1] == False


def test_dedup_per_base_product_then_compute_resolved_series_matches_expected_scope():
    """
    End-to-end shape of the actual fix: build a scope frame with
    dedup_per_base_product() (broader than product_to_sheet), then resolve
    it with compute_resolved_series() — a product outside product_to_sheet
    must survive both steps and be resolved correctly, not vanish at
    either one.
    """
    broad_scope = _df([
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVLT001', 'Vulnerability Name': 'CVE-2026-0001',   # duplicate, should collapse
         'Base Product': 'Google Chrome', 'Affected Products': 'Google Chrome',
         'Threat Status': 'UNRESOLVED'},
        {'Name': 'CVFANTOM', 'Vulnerability Name': 'CVE-2026-7777',
         'Base Product': 'PuTTY', 'Affected Products': 'PuTTY',
         'Threat Status': 'RESOLVED'},
    ])
    product_to_sheet = {'Google Chrome': 'Google Chrome'}   # PuTTY absent — simulates the scope gap

    scope = dedup_per_base_product(broad_scope)
    assert len(scope) == 2   # duplicate Chrome row collapsed, PuTTY row kept

    resolved = compute_resolved_series(scope, product_to_sheet, patch_resolved_pairs=None)
    unresolved_count = int((~resolved).sum())
    assert unresolved_count == 1   # only the Chrome row — PuTTY correctly resolved, not dropped or forced False