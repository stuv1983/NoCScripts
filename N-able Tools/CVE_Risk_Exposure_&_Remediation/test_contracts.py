"""
test_contracts.py — every test class is named for the historical bug (or
failure class) its contract guards. If a contract check ever weakens, the
test that fails tells you which production incident it was protecting
against.

Run: pytest test_contracts.py -v
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import pytest

from contracts import (
    ContractError,
    check_frame, check_vocabulary,
    check_cve_export, check_rmm_inventory, check_merged,
    check_scopes, check_patch_match, run_check,
    PATCH_INTERNAL_COLS,
)


# ── fixtures ──────────────────────────────────────────────────────────────────

def _cve_df(**overrides) -> pd.DataFrame:
    base = {
        'Name': ['WS01', 'WS02'],
        'Vulnerability Name': ['CVE-2026-0001', 'CVE-2026-0002'],
        'Vulnerability Score': [9.8, 7.5],
        'Threat Status': ['UNRESOLVED', 'RESOLVED'],
        'Affected Products': ['Google Chrome', 'Microsoft Edge'],
    }
    base.update(overrides)
    return pd.DataFrame(base)


def _rmm_df(**overrides) -> pd.DataFrame:
    base = {
        'Device': ['WS01', 'WS02'],
        'Last Response': ['2026-07-01 10:00', '2026-07-02 11:00'],
        'Device_Join': ['WS01', 'WS02'],
        'Device Type': ['Workstation', 'Server'],
        'Username': ['alice', 'bob'],
    }
    base.update(overrides)
    return pd.DataFrame(base)


def _merged_df() -> pd.DataFrame:
    return pd.DataFrame({
        'Name': ['WS01', 'WS02', 'WS03', 'WS04'],
        'Vulnerability Name': [f'CVE-2026-000{i}' for i in range(1, 5)],
        'Vulnerability Score': [9.8, 9.1, 7.5, 4.0],
        'Last Response': ['2026-07-01', 'Not Found in RMM', '2026-07-02', '2026-07-03'],
        'Device Type': ['Workstation', 'Unknown', 'Server', 'Workstation'],
        'Username': ['alice', '', 'bob', 'carol'],
    })


def _scopes(merged, threshold=9.0):
    filtered = merged[merged['Vulnerability Score'] >= threshold]
    triage = filtered[filtered['Last Response'] != 'Not Found in RMM']
    return filtered, triage


def _patch_full(**overrides) -> pd.DataFrame:
    base = {
        'Name': ['WS01'],
        'Vulnerability Name': ['CVE-2026-0001'],
        'Status': ['Installed'],
        'Patch Match Result': ['Matched - installed'],
        'Version Check Result': ['Version compliant'],
        'Patch Evidence Status': ['Patch confirmed - pending rescan'],
    }
    base.update(overrides)
    return pd.DataFrame(base)


# ══════════════════════════════════════════════════════════════════════════════
# Generic building blocks
# ══════════════════════════════════════════════════════════════════════════════

class TestCheckFrame:
    def test_passes_on_clean_frame(self):
        check_frame(_cve_df(), 't', required=('Name',))

    def test_missing_required_raises_with_column_name(self):
        with pytest.raises(ContractError, match=r"missing required.*Nope"):
            check_frame(_cve_df(), 't', required=('Nope',))

    def test_forbidden_column_raises(self):
        df = _cve_df()
        df['_ck'] = 'x'
        with pytest.raises(ContractError, match=r"leaked.*_ck"):
            check_frame(df, 't', forbidden=('_ck',))

    def test_duplicate_column_labels_raise(self):
        # The historical 'Client' + 'Customer Name' → two 'Customer' columns
        # alias bug produced exactly this shape.
        df = pd.concat([_cve_df(), _cve_df()[['Name']]], axis=1)
        assert df.columns.duplicated().any()
        with pytest.raises(ContractError, match='duplicate column'):
            check_frame(df, 't')

    def test_non_numeric_in_numeric_column_raises(self):
        df = _cve_df(**{'Vulnerability Score': [9.8, 'high']})
        with pytest.raises(ContractError, match='non-numeric'):
            check_frame(df, 't', numeric=('Vulnerability Score',))

    def test_empty_frame_rejected_when_disallowed(self):
        with pytest.raises(ContractError, match='empty'):
            check_frame(pd.DataFrame(), 't', allow_empty=False)

    def test_non_dataframe_raises(self):
        with pytest.raises(ContractError, match='expected DataFrame'):
            check_frame([1, 2, 3], 't')  # type: ignore[arg-type]


class TestCheckVocabulary:
    def test_unknown_value_warns_not_raises(self):
        df = _cve_df(**{'Threat Status': ['UNRESOLVED', 'IN PROGRESS']})
        issues = check_vocabulary(df, 't', 'Threat Status',
                                  allowed={'RESOLVED', 'UNRESOLVED'})
        assert len(issues) == 1 and 'IN PROGRESS' in issues[0]

    def test_forbidden_value_hard_raises(self):
        df = _patch_full(Status=['RESOLVED'])
        with pytest.raises(ContractError, match='forbidden'):
            check_vocabulary(df, 't', 'Status',
                             forbidden_values=('RESOLVED', 'UNRESOLVED'),
                             hard=True)

    def test_comparison_is_case_and_whitespace_insensitive(self):
        df = _cve_df(**{'Threat Status': ['  unresolved ', 'Resolved']})
        assert check_vocabulary(df, 't', 'Threat Status',
                                allowed={'RESOLVED', 'UNRESOLVED'}) == []

    def test_absent_column_is_a_noop(self):
        assert check_vocabulary(_cve_df(), 't', 'Nope', allowed={'X'}) == []

    def test_blank_values_ignored_by_default(self):
        df = _cve_df(**{'Threat Status': ['UNRESOLVED', '']})
        assert check_vocabulary(df, 't', 'Threat Status',
                                allowed={'RESOLVED', 'UNRESOLVED'}) == []


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 1 — CVE export
# ══════════════════════════════════════════════════════════════════════════════

class TestCveExportContract:
    def test_clean_export_passes_with_no_warnings(self):
        assert check_cve_export(_cve_df()) == []

    def test_missing_vulnerability_name_raises(self):
        df = _cve_df().drop(columns=['Vulnerability Name'])
        with pytest.raises(ContractError, match='Vulnerability Name'):
            check_cve_export(df)

    def test_score_out_of_cvss_range_warns(self):
        # e.g. a percentage column mis-aliased onto Vulnerability Score
        df = _cve_df(**{'Vulnerability Score': [98.0, 7.5]})
        issues = check_cve_export(df)
        assert any('outside 0–10' in m for m in issues)

    def test_novel_threat_status_value_warns_but_does_not_raise(self):
        df = _cve_df(**{'Threat Status': ['UNRESOLVED', 'ACKNOWLEDGED']})
        issues = check_cve_export(df)
        assert any('ACKNOWLEDGED' in m for m in issues)

    def test_blank_device_names_warn(self):
        df = _cve_df(Name=['WS01', '  '])
        issues = check_cve_export(df)
        assert any('blank device Name' in m for m in issues)

    def test_empty_export_raises(self):
        with pytest.raises(ContractError, match='empty'):
            check_cve_export(_cve_df().iloc[0:0])


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 2 — RMM inventory
# ══════════════════════════════════════════════════════════════════════════════

class TestRmmInventoryContract:
    def test_clean_inventory_passes(self):
        assert check_rmm_inventory(_rmm_df()) == []

    def test_duplicate_join_keys_raise(self):
        # load_rmm_data promises drop_duplicates on Device_Join; a duplicate
        # here fans out every matching CVE row in the merge.
        df = _rmm_df(Device_Join=['WS01', 'WS01'])
        with pytest.raises(ContractError, match='duplicate Device_Join'):
            check_rmm_inventory(df)

    def test_unknown_device_type_warns(self):
        df = _rmm_df(**{'Device Type': ['Workstation', 'Laptop']})
        issues = check_rmm_inventory(df)
        assert any('LAPTOP' in m for m in issues)

    def test_missing_username_column_raises(self):
        # v0.8 contract: load_rmm_data must synthesise Username when absent.
        df = _rmm_df().drop(columns=['Username'])
        with pytest.raises(ContractError, match='Username'):
            check_rmm_inventory(df)


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 3 — merged frame
# ══════════════════════════════════════════════════════════════════════════════

class TestMergedContract:
    def test_clean_merge_passes(self):
        assert check_merged(_merged_df()) == []

    def test_missing_last_response_raises(self):
        # merge_data guarantees this column exists even with skip_rmm=True.
        df = _merged_df().drop(columns=['Last Response'])
        with pytest.raises(ContractError, match='Last Response'):
            check_merged(df)

    def test_category_roundtrip_regression_warns(self):
        # v0.12: Device Type should be category dtype after merge_data's
        # final re-downcast. object dtype ⇒ someone added a write after it.
        df = _merged_df()  # object dtype
        issues = check_merged(df, expect_category_cols=('Device Type',))
        assert any('category' in m for m in issues)

        df['Device Type'] = df['Device Type'].astype('category')
        assert check_merged(df, expect_category_cols=('Device Type',)) == []


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 4 — scope frames (v0.3 Triage Scope Bug)
# ══════════════════════════════════════════════════════════════════════════════

class TestScopesContractV03TriageScopeBug:
    def test_correctly_built_scopes_pass(self):
        merged = _merged_df()
        filtered, triage = _scopes(merged)
        assert check_scopes(merged, filtered, triage, 9.0) == []

    def test_not_in_rmm_row_in_triage_raises(self):
        """THE v0.3 bug: a Not-Found-in-RMM row reaching triage scope."""
        merged = _merged_df()
        filtered, _ = _scopes(merged)
        bad_triage = filtered  # forgot the Last Response filter
        with pytest.raises(ContractError, match='Not Found in RMM'):
            check_scopes(merged, filtered, bad_triage, 9.0)

    def test_below_threshold_row_in_filtered_raises(self):
        """v0.3 companion: threshold compared against the wrong column."""
        merged = _merged_df()
        filtered = merged  # forgot the score filter — row at 4.0 slips in
        triage = filtered[filtered['Last Response'] != 'Not Found in RMM']
        with pytest.raises(ContractError, match='below the CVSS threshold'):
            check_scopes(merged, filtered, triage, 9.0)

    def test_triage_not_subset_of_filtered_raises(self):
        merged = _merged_df()
        filtered, _ = _scopes(merged)
        alien = merged.iloc[[3]]  # score 4.0 — not in filtered
        bad_triage = pd.concat([filtered, alien])
        with pytest.raises(ContractError, match='not present in filtered_df'):
            check_scopes(merged, filtered, bad_triage, 9.0)

    def test_health_floor_below_seven_raises(self):
        # The min(threshold, 7.0) regression that made the health score
        # swing with the display threshold.
        merged = _merged_df()
        filtered, triage = _scopes(merged)
        hf = merged[merged['Vulnerability Score'] >= 1.0]
        ht = hf[hf['Last Response'] != 'Not Found in RMM']
        with pytest.raises(ContractError, match='fixed at 7.0'):
            check_scopes(merged, filtered, triage, 9.0,
                         health_filtered=hf, health_triage_df=ht,
                         health_score_threshold=1.0)

    def test_valid_health_scope_passes(self):
        merged = _merged_df()
        filtered, triage = _scopes(merged)
        hf = merged[merged['Vulnerability Score'] >= 7.0]
        ht = hf[hf['Last Response'] != 'Not Found in RMM']
        assert check_scopes(merged, filtered, triage, 9.0,
                            health_filtered=hf, health_triage_df=ht,
                            health_score_threshold=7.0) == []

    def test_empty_scopes_pass(self):
        merged = _merged_df()
        empty = merged.iloc[0:0]
        assert check_scopes(merged, empty, empty, 9.0) == []


# ══════════════════════════════════════════════════════════════════════════════
# Boundary 5 — patch match (v0.4 Status Column Collision)
# ══════════════════════════════════════════════════════════════════════════════

class TestPatchMatchContractV04StatusCollision:
    def test_clean_patch_output_passes(self):
        assert check_patch_match(_patch_full()) == []

    def test_cve_status_in_patch_status_column_raises(self):
        """THE v0.4 bug, both directions of reintroduction: the merge kept
        the CVE export's RESOLVED/UNRESOLVED where the patch install status
        belongs. This is the check that would have caught the 763-row
        incident on run one."""
        bad = _patch_full(Status=['UNRESOLVED'])
        with pytest.raises(ContractError, match='forbidden'):
            check_patch_match(bad)

    def test_lowercase_cve_status_also_caught(self):
        bad = _patch_full(Status=['resolved'])
        with pytest.raises(ContractError, match='forbidden'):
            check_patch_match(bad)

    def test_internal_working_columns_must_be_dropped(self):
        for col in ('_cve_status_orig', '_ck', '_mck', '_patch_status'):
            bad = _patch_full()
            bad[col] = 'x'
            with pytest.raises(ContractError, match='leaked'):
                check_patch_match(bad)

    def test_all_internal_cols_are_underscore_prefixed(self):
        # Keeps the "drop everything underscore-prefixed" cleanup honest.
        assert all(c.startswith('_') for c in PATCH_INTERNAL_COLS)

    def test_confirmed_without_compliant_version_warns(self):
        # Classifier AND-chain: confirmed requires version-compliant.
        bad = _patch_full(**{'Version Check Result': ['Below fixed version']})
        issues = check_patch_match(bad)
        assert any('AND-chain' in m for m in issues)

    def test_none_and_empty_frames_are_noops(self):
        assert check_patch_match(None) == []
        assert check_patch_match(_patch_full().iloc[0:0]) == []


# ══════════════════════════════════════════════════════════════════════════════
# run_check plumbing
# ══════════════════════════════════════════════════════════════════════════════

class TestRunCheck:
    def test_semantic_issues_collected_into_warnings_list(self):
        warnings: list[str] = []
        df = _cve_df(**{'Threat Status': ['UNRESOLVED', 'WEIRD']})
        run_check(check_cve_export, df, warnings=warnings)
        assert len(warnings) == 1
        assert warnings[0].startswith('Data contract:')

    def test_structural_failure_propagates(self):
        warnings: list[str] = []
        with pytest.raises(ContractError):
            run_check(check_cve_export, pd.DataFrame({'Name': ['x']}),
                      warnings=warnings)
        assert warnings == []


# ══════════════════════════════════════════════════════════════════════════════
# Static guard — the itertuples failure class
# ══════════════════════════════════════════════════════════════════════════════

# itertuples() renames space-containing column names ('Device Name',
# 'Patch Evidence Notes') to positional _1/_2..., so getattr silently
# returns defaults forever — the blank-device-names bug (see the NOTE in
# diagnostics.py). Rules enforced here:
#
#   • itertuples(..., name=None) → plain tuples, positional access only,
#     no mangling possible → always allowed.
#   • NAMED itertuples (default namedtuple) → counted against a frozen,
#     reviewed baseline. Existing uses have been audited as safe (see
#     comments below); any NEW named itertuples fails this test and must
#     either use name=None, switch to to_dict('records'), or be reviewed
#     and added to the baseline with a justification comment.
#
# Reviewed baseline as of 2026-07:
#   device_sheets.py:204  — converts to list(row) immediately, positional. Safe.
#   summary_sheet.py:1605 — getattr only on space-free agg columns
#                            (Name, device_type, has_exploit). Safe but
#                            fragile: adding a spaced column to that agg
#                            and reading it via getattr WILL silently break.
NAMED_ITERTUPLES_BASELINE = {
    'summary_sheet.py': 2,   # two calls on the same _sorted frame
    'product_sheets.py': 0,  # uses name=None — not counted
    'trend_sheets.py': 0,
    'device_sheets.py': 1,
    'patch_sheets.py': 0,
    'excel_builder.py': 0,
    'sheet_helpers.py': 0,
    'diagnostics.py': 0,
}

_ITERTUPLES_CALL_RE = re.compile(r'\.itertuples\(([^)]*)\)')


def _count_named_itertuples(source: str) -> int:
    n = 0
    for line in source.splitlines():
        code = line.split('#', 1)[0]  # ignore comments
        for m in _ITERTUPLES_CALL_RE.finditer(code):
            if 'name=None' not in m.group(1).replace(' ', ''):
                n += 1
    return n


@pytest.mark.parametrize('module_name', sorted(NAMED_ITERTUPLES_BASELINE))
def test_no_new_named_itertuples_in_sheet_builders(module_name):
    """Ratchet: named-tuple itertuples count must not grow past the
    reviewed baseline. Skips when the module isn't present so this file
    also runs standalone outside the project tree."""
    path = Path(__file__).resolve().parent / module_name
    if not path.exists():
        pytest.skip(f'{module_name} not found next to test file')
    source = path.read_text(encoding='utf-8', errors='replace')
    count = _count_named_itertuples(source)
    baseline = NAMED_ITERTUPLES_BASELINE[module_name]
    assert count <= baseline, (
        f"{module_name}: {count} named itertuples() call(s), reviewed "
        f"baseline is {baseline}. pandas renames columns containing spaces "
        f"to positional _N names, so getattr access silently breaks. Use "
        f"itertuples(name=None) with positional access, or "
        f"df.to_dict('records') — or review the new call and update "
        f"NAMED_ITERTUPLES_BASELINE with a justification comment."
    )
    if count < baseline:
        # Baseline can be ratcheted DOWN — not a failure, but nudge.
        pytest.skip(
            f'{module_name}: count {count} below baseline {baseline} — '
            f'consider lowering NAMED_ITERTUPLES_BASELINE'
        )