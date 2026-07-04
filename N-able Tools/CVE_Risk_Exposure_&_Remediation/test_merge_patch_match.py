"""
test_merge_patch_match.py — regression tests for data_pipeline.merge_data()
and data_pipeline.process_patch_match().

These two functions carry the highest concentration of past silent-data-
corruption bugs per CHANGELOG.md (RMM join direction, Device Type inference,
and the v0.4 'Status Column Collision' bug), yet had no dedicated test file
before this one — test_patch_resolution.py covers the row-wise
_classify_version_check / _classify_resolution / _classify_baseline_compliance
helpers, but process_patch_match() itself does not call those; it uses its
own vectorised _vec_vcr / _vec_bc / _vec_pes / _vec_fixed_version / _vec_baseline
implementations. This file exercises that actual code path.

Run with: pytest test_merge_patch_match.py -v

Author : Stu Villanti <s.villanti@kenstra.com>
"""

import os
import sys
import types

import pytest
from unittest.mock import patch

os.environ.setdefault('PYTEST_CURRENT_TEST', 'bootstrap')

# ---------------------------------------------------------------------------
# Config stubbing — same pattern as test_resolution.py / test_patch_resolution.py,
# so this file can run standalone without a real config.json on disk.
# See test_resolution.py's comment for why setdefault (not forced assignment)
# matters when multiple test files stub 'config' in the same pytest session.
# ---------------------------------------------------------------------------

import re as _re
_fake_config = types.ModuleType('config')
_fake_config.CVE_PATTERN = _re.compile(r'(CVE-\d{4}-\d{4,7})', _re.IGNORECASE)
_fake_config.PRODUCT_MAP = [
    ('google chrome',  'chrome'),
    ('mozilla firefox', 'firefox'),
    ('microsoft edge', 'edge'),
]
_fake_config.FIXED_VERSION_RULES = {
    'chrome': {
        '_baseline':     '148.0.0.0',
        'CVE-2026-1234': '147.0.0.0',
    },
}
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

project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import pandas as pd  # noqa: E402

import data_pipeline  # noqa: E402
from data_pipeline import merge_data, process_patch_match  # noqa: E402

# ---------------------------------------------------------------------------
# FIXED_VERSION_RULES is a single dict object bound once, whenever
# data_pipeline is FIRST imported anywhere in the pytest session — whichever
# test file's fake config module won the sys.modules race at that moment
# "wins" for every other test file too, since sys.modules.setdefault is a
# no-op once anything has already set 'config'. patch.dict the exact rules
# this file needs around every test and restore afterward, so results here
# are correct regardless of what other test files ran before it in the same
# session (see test_patch_resolution.py's copy of this fixture for the
# full explanation).
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def _isolated_fixed_version_rules():
    with patch.dict(data_pipeline.FIXED_VERSION_RULES, _fake_config.FIXED_VERSION_RULES, clear=True):
        yield


# ===========================================================================
# merge_data()
# ===========================================================================

def _vuln_row(name='WS01', score=9.5, **extra):
    row = {
        'Name': name,
        'Name_Join': name,
        'Vulnerability Name': 'CVE-2026-0001',
        'Vulnerability Score': score,
        'Affected Products': 'Google Chrome',
    }
    row.update(extra)
    return row


class TestMergeDataJoinDirection:
    """
    CHANGELOG v0.3: exclude_missing_rmm=True used to be the default and
    silently dropped devices absent from the RMM inventory via an INNER join,
    before any threshold filtering ran. The fix made exclude_missing_rmm=False
    (LEFT join) the default, marking unmatched devices 'Not Found in RMM'
    instead of deleting them. Both behaviours must keep working correctly,
    since exclude_missing_rmm is still an explicit option.
    """

    def test_left_join_keeps_devices_missing_from_rmm(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01'), _vuln_row('GHOST01')])
        df_rmm = pd.DataFrame([{
            'Device_Join': 'WS01', 'Last Response': '2026-06-01', 'Device Type': 'Workstation',
        }])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=False)

        assert len(merged) == 2, "LEFT join must keep the unmatched device, not drop it"
        ghost = merged[merged['Name'] == 'GHOST01'].iloc[0]
        assert ghost['Last Response'] == 'Not Found in RMM'
        assert ghost['Device Type'] == 'Unknown'
        matched = merged[merged['Name'] == 'WS01'].iloc[0]
        assert matched['Last Response'] == '2026-06-01'
        assert matched['Device Type'] == 'Workstation'

    def test_inner_join_excludes_devices_missing_from_rmm(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01'), _vuln_row('GHOST01')])
        df_rmm = pd.DataFrame([{
            'Device_Join': 'WS01', 'Last Response': '2026-06-01', 'Device Type': 'Workstation',
        }])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=True)

        assert len(merged) == 1, "INNER join (exclude_missing_rmm=True) must drop the unmatched device"
        assert merged.iloc[0]['Name'] == 'WS01'

    def test_skip_rmm_true_never_touches_rmm_frame(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01')])
        # A malformed/irrelevant df_rmm must be ignored entirely when skip_rmm=True.
        df_rmm = pd.DataFrame([{'Device_Join': 'WS01'}])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=True)

        assert len(merged) == 1
        assert merged.iloc[0]['Last Response'] == 'N/A'
        assert merged.iloc[0]['Device Type'] == 'Unknown'

    def test_skip_rmm_true_with_no_rmm_frame_at_all(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01')])
        merged = merge_data(df_vuln, None, skip_rmm=True)
        assert len(merged) == 1
        assert merged.iloc[0]['Last Response'] == 'N/A'


class TestMergeDataDoesNotOverwriteExistingColumns:
    """
    merge_data only pulls 'Last Response' / 'Device Type' from the RMM frame
    when df_vuln itself lacks those columns (vuln_has_lr / vuln_has_dt guards).
    This matters for re-processing a previously-merged export (e.g. a
    dashboard's Raw Data re-fed as input) — the RMM inventory must not
    clobber values the vuln export already carries.
    """

    def test_existing_last_response_is_not_overwritten_by_rmm(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01', **{'Last Response': '2026-01-01'})])
        df_rmm = pd.DataFrame([{'Device_Join': 'WS01', 'Last Response': '2026-06-01',
                               'Device Type': 'Workstation'}])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=False)

        assert merged.iloc[0]['Last Response'] == '2026-01-01', (
            "df_vuln's own Last Response must win — RMM data should not be pulled "
            "in at all when the column already exists on the vuln export"
        )

    def test_existing_device_type_is_not_overwritten_by_rmm(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01', **{'Device Type': 'Server'})])
        df_rmm = pd.DataFrame([{'Device_Join': 'WS01', 'Device Type': 'Workstation',
                               'Last Response': '2026-06-01'}])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=False)

        assert merged.iloc[0]['Device Type'] == 'Server'

    def test_username_is_pulled_from_rmm_when_absent_on_vuln(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01')])
        df_rmm = pd.DataFrame([{
            'Device_Join': 'WS01', 'Last Response': '2026-06-01',
            'Device Type': 'Workstation', 'Username': 'jsmith',
        }])

        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=False)

        assert merged.iloc[0]['Username'] == 'jsmith'


class TestMergeDataDeviceTypeInference:
    """
    Device Type falls back through: RMM Device Type -> Operating System Role
    -> OS text sniffing -> 'Unknown'. Each stage only fires when the
    previous one left the value at 'Unknown'.
    """

    def test_infers_server_from_os_text_when_device_type_unknown(self):
        df_vuln = pd.DataFrame([_vuln_row('SRV01', OS='Windows Server 2022 Standard')])
        merged = merge_data(df_vuln, None, skip_rmm=True)
        assert merged.iloc[0]['Device Type'] == 'Server'

    def test_infers_workstation_from_os_text_when_device_type_unknown(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01', OS='Windows 11 Pro')])
        merged = merge_data(df_vuln, None, skip_rmm=True)
        assert merged.iloc[0]['Device Type'] == 'Workstation'

    def test_operating_system_role_used_when_skipping_rmm(self):
        df_vuln = pd.DataFrame([_vuln_row('SRV01', **{'Operating System Role': 'server'})])
        merged = merge_data(df_vuln, None, skip_rmm=True)
        assert merged.iloc[0]['Device Type'] == 'Server'


class TestMergeDataDaysSinceLastResponse:

    def test_days_since_last_response_uses_as_of_date_not_wall_clock(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01', **{'Last Response': '2026-06-01'})])
        merged = merge_data(df_vuln, None, skip_rmm=True, as_of_date='2026-06-15')
        assert merged.iloc[0]['Days Since Last Response'] == 14

    def test_not_found_in_rmm_has_no_days_value(self):
        df_vuln = pd.DataFrame([_vuln_row('WS01'), _vuln_row('GHOST01')])
        df_rmm = pd.DataFrame([{'Device_Join': 'WS01', 'Last Response': '2026-06-01',
                               'Device Type': 'Workstation'}])
        merged = merge_data(df_vuln, df_rmm, skip_rmm=False, exclude_missing_rmm=False,
                            as_of_date='2026-06-15')
        ghost = merged[merged['Name'] == 'GHOST01'].iloc[0]
        assert ghost['Days Since Last Response'] == '—'


# ===========================================================================
# process_patch_match()
# ===========================================================================

def _write_patch_csv(tmp_path, rows):
    path = tmp_path / 'patch.csv'
    pd.DataFrame(rows).to_csv(path, index=False)
    return str(path)


def _cve_row(name='WS01', cve='CVE-2026-1234', product='Google Chrome',
             score=9.8, status_col=None, status_val=None,
             first_detected='2026-05-01', date_published='2026-05-01'):
    row = {
        'Name': name, 'Vulnerability Name': cve, 'Affected Products': product,
        'Vulnerability Score': score, 'Customer': 'Acme', 'Site': 'HQ',
        'First detected': first_detected, 'Date Published': date_published,
    }
    if status_col:
        row[status_col] = status_val
    return row


class TestProcessPatchMatchColumnValidation:

    def test_missing_patch_columns_raises_value_error(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01'}])
        cve_df = pd.DataFrame([_cve_row()])
        with pytest.raises(ValueError, match='Patch report missing required columns'):
            process_patch_match(patch_csv, cve_df)

    def test_missing_cve_columns_raises_value_error(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.0.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([{'Name': 'WS01'}])  # missing Vulnerability Name / Affected Products
        with pytest.raises(ValueError, match='CVE data missing columns'):
            process_patch_match(patch_csv, cve_df)


class TestProcessPatchMatchStatusColumnCollision:
    """
    CHANGELOG v0.4 'Status Column Collision' bug, reintroduced from the
    opposite side after the vectorised rewrite: some N-able exports carry the
    CVE's own resolution state in a column literally named 'Status'
    (RESOLVED/UNRESOLVED) rather than 'Threat Status' (see
    _active_trend_scope's Threat Status / Status fallback for the same
    variability). The patch report also has its own 'Status' column
    (Installed/Pending/etc.). Merging the two without renaming one of them
    collides, and pandas' suffixes=('', '_p') silently keeps the LEFT (cve)
    column, so downstream code reads the CVE's RESOLVED/UNRESOLVED value
    where it expects the patch's install status.
    """

    def test_installed_and_compliant_is_not_masked_by_a_bare_status_column(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0', 'Discovered / Install Date': '2026-06-01',
        }])
        # CVE export's own resolution status happens to be named 'Status'.
        cve_df = pd.DataFrame([_cve_row(status_col='Status', status_val='UNRESOLVED')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Status'] == 'Installed', (
            "best['Status'] must be the PATCH report's install status after the "
            "merge, never the CVE export's own RESOLVED/UNRESOLVED value"
        )
        assert row['Patch Match Result'] == 'Matched - installed'
        assert row['Version Check Result'] == 'Version compliant'
        assert row['Patch Evidence Status'] == 'Patch confirmed - pending rescan'

    def test_same_scenario_with_threat_status_column_name_also_works(self, tmp_path):
        """Control case: the more common column name must keep working too."""
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([_cve_row(status_col='Threat Status', status_val='UNRESOLVED')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Status'] == 'Installed'
        assert row['Patch Evidence Status'] == 'Patch confirmed - pending rescan'


class TestProcessPatchMatchVersionClassification:

    def test_installed_below_fixed_version_stays_unresolved(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 140.0.0.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([_cve_row()])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Version Check Result'] == 'Below fixed version'
        assert row['Patch Evidence Status'] == 'Unresolved'

    def test_patch_installed_before_cve_detected_does_not_confirm_resolution(self, tmp_path):
        """
        A patch installed BEFORE the CVE was even published/detected cannot be
        the fix for it — Patch Evidence Status must require
        install_date >= max(First detected, Date Published).
        """
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0', 'Discovered / Install Date': '2026-01-01',
        }])
        cve_df = pd.DataFrame([_cve_row(first_detected='2026-05-01', date_published='2026-05-01')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Version Check Result'] == 'Version compliant'
        assert row['Patch Evidence Status'] == 'Unresolved', (
            "an install predating the CVE's own detection/publication date "
            "must not be treated as evidence of resolution"
        )

    def test_device_not_in_patch_report_at_all(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'OTHERDEVICE', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([_cve_row(name='WS01')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Patch Match Result'] == 'Not found in patch report'
        assert row['Patch Evidence Status'] == 'Unresolved'

    def test_device_in_report_but_product_not_matched(self, tmp_path):
        """Device exists in the patch report, but no row there matches this product."""
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Adobe Reader 24.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([_cve_row(product='Google Chrome')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Patch Match Result'] == 'Device in patch report - product not found'

    def test_architecture_mismatch_is_not_treated_as_a_match(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0 (x86)', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([_cve_row(product='Google Chrome (x64)')])

        _, full, _, _, _ = process_patch_match(patch_csv, cve_df, min_score=9.0)
        row = full.iloc[0]

        assert row['Patch Match Result'] == 'Device in patch report - product not found'


class TestProcessPatchMatchScoreFiltering:

    def test_min_score_filters_rows_before_matching(self, tmp_path):
        patch_csv = _write_patch_csv(tmp_path, [{
            'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01', 'Status': 'Installed',
            'Patch': 'Google Chrome 148.0.1.0', 'Discovered / Install Date': '2026-06-01',
        }])
        cve_df = pd.DataFrame([
            _cve_row(name='WS01', cve='CVE-2026-1234', score=9.8),
            _cve_row(name='WS01', cve='CVE-2026-9999', score=5.0),
        ])

        _, full, _, total_rows, filtered_rows = process_patch_match(patch_csv, cve_df, min_score=9.0)

        assert total_rows == 2
        assert filtered_rows == 1
        assert len(full) == 1
        assert full.iloc[0]['Vulnerability Name'] == 'CVE-2026-1234'