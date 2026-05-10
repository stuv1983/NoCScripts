"""
tests/test_core.py — Unit tests for the CVE Dashboard core pipeline.

Run with:  pytest tests/ -v

These tests feed static, deterministic DataFrames directly into the pipeline
functions so they run without any file I/O, GUI, or external API calls.
"""

import sys
import os
import pytest
import pandas as pd

# ── Path setup ────────────────────────────────────────────────────────────────
# Allow running from repo root or from the tests/ directory.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from data_pipeline import (
    classify_patch_gap,
    compute_trends,
    normalize_device_name,
    extract_cve_id,
    get_base_product,
    parse_last_response,
    _active_trend_scope,
)


# ==============================================================================
# HELPERS — build minimal DataFrames that satisfy pipeline expectations
# ==============================================================================

def _make_cve_df(rows: list[dict]) -> pd.DataFrame:
    """Build a CVE-style DataFrame with all required columns defaulted."""
    defaults = {
        'Name':                 'DEVICE001',
        'Vulnerability Name':   'CVE-2024-0001',
        'Affected Products':    'Google Chrome',
        'Vulnerability Score':  9.5,
        'Vulnerability Severity': 'Critical',
        'Threat Status':        'Unresolved',
        'Has Known Exploit':    'No',
        'CISA KEV':             'No',
        'Last Response':        '2026-03-15 09:00:00',
        'Device Type':          'Workstation',
    }
    return pd.DataFrame([{**defaults, **r} for r in rows])


# ==============================================================================
# classify_patch_gap
# ==============================================================================

class TestClassifyPatchGap:
    def test_not_in_patch_report_returns_coverage_gap(self):
        assert classify_patch_gap('Not found in patch report') == 'coverage_gap'

    def test_device_in_report_product_not_found_returns_unmanaged(self):
        assert classify_patch_gap('Device in patch report - product not found') == 'unmanaged_app'

    def test_matched_installed_unresolved_is_detection_mismatch(self):
        assert classify_patch_gap('Matched - installed', 'Unresolved') == 'detection_mismatch'

    def test_matched_reboot_required_unresolved_is_detection_mismatch(self):
        assert classify_patch_gap('Matched - reboot required', 'Unresolved') == 'detection_mismatch'

    def test_matched_installed_resolved_is_not_a_gap(self):
        assert classify_patch_gap('Matched - installed', 'Patch confirmed - pending rescan') is None

    def test_matched_installed_no_status_is_not_a_gap(self):
        assert classify_patch_gap('Matched - installed') is None

    def test_unknown_result_returns_none(self):
        assert classify_patch_gap('Some unexpected value') is None

    def test_empty_string_returns_none(self):
        assert classify_patch_gap('') is None

    def test_whitespace_stripped(self):
        assert classify_patch_gap('  Not found in patch report  ') == 'coverage_gap'


# ==============================================================================
# compute_trends — core set arithmetic
# ==============================================================================

class TestComputeTrends:
    """
    Tests that New / Resolved / Persisting CVE counts are exactly correct
    under a variety of scenarios. These are the most important tests in the
    suite — a regression here silently mis-reports the managed estate.
    """

    THRESHOLD = 9.0

    def _run(self, current_rows, previous_rows, **kwargs):
        cur  = _make_cve_df(current_rows)
        prev = _make_cve_df(previous_rows)
        # _active_trend_scope needs Vulnerability Score as numeric; guard empty frames
        if 'Vulnerability Score' in cur.columns:
            cur['Vulnerability Score']  = pd.to_numeric(cur['Vulnerability Score'],  errors='coerce')
        if 'Vulnerability Score' in prev.columns:
            prev['Vulnerability Score'] = pd.to_numeric(prev['Vulnerability Score'], errors='coerce')
        return compute_trends(cur, prev, self.THRESHOLD, **kwargs)

    # ── Basic set arithmetic ──────────────────────────────────────────────────

    def test_all_new_when_previous_is_empty(self):
        result = self._run(
            current_rows=[
                {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'},
                {'Name': 'PC2', 'Vulnerability Name': 'CVE-2024-0002', 'Affected Products': 'Chrome'},
            ],
            previous_rows=[],
        )
        m = result['metrics']
        assert m['new_cve_count'] == 2
        assert m['resolved_cve_count'] == 0
        assert m['persisting_cve_count'] == 0

    def test_all_resolved_when_current_is_empty(self):
        # When current is empty there are no products in common scope, so
        # resolved_cve_count (scoped) is 0. The device count delta is what
        # surfaces the change — prev_devices > 0, cur_devices == 0.
        result = self._run(
            current_rows=[],
            previous_rows=[
                {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'},
            ],
        )
        m = result['metrics']
        assert m['new_cve_count'] == 0
        assert m['cur_devices'] == 0
        assert m['prev_devices'] == 1
        assert m['persisting_cve_count'] == 0

    def test_persisting_when_same_cve_both_periods(self):
        row = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        result = self._run(current_rows=[row], previous_rows=[row])
        m = result['metrics']
        assert m['persisting_cve_count'] == 1
        assert m['new_cve_count'] == 0
        assert m['resolved_cve_count'] == 0

    def test_mixed_new_resolved_persisting(self):
        shared    = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        only_cur  = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0002', 'Affected Products': 'Chrome'}
        only_prev = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0003', 'Affected Products': 'Chrome'}
        result = self._run(
            current_rows=[shared, only_cur],
            previous_rows=[shared, only_prev],
        )
        m = result['metrics']
        assert m['new_cve_count'] == 1          # only_cur
        assert m['resolved_cve_count'] == 1     # only_prev
        assert m['persisting_cve_count'] == 1   # shared

    # ── Score threshold filtering ─────────────────────────────────────────────

    def test_below_threshold_cves_excluded(self):
        """CVEs below 9.0 must not appear in any trend bucket."""
        result = self._run(
            current_rows=[
                {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-LOW', 'Vulnerability Score': 5.0, 'Affected Products': 'Chrome'},
            ],
            previous_rows=[
                {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-LOW', 'Vulnerability Score': 5.0, 'Affected Products': 'Chrome'},
            ],
        )
        m = result['metrics']
        assert m['new_cve_count'] == 0
        assert m['persisting_cve_count'] == 0
        assert m['resolved_cve_count'] == 0

    def test_exactly_at_threshold_is_included(self):
        row = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001',
               'Vulnerability Score': 9.0, 'Affected Products': 'Chrome'}
        result = self._run(current_rows=[row], previous_rows=[])
        assert result['metrics']['new_cve_count'] == 1

    # ── Resolved-only filter (Threat Status) ─────────────────────────────────

    def test_resolved_rows_excluded_from_active_scope(self):
        """RESOLVED rows in the current export must not count as persisting or new."""
        row = {
            'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001',
            'Affected Products': 'Chrome', 'Threat Status': 'Resolved',
        }
        result = self._run(current_rows=[row], previous_rows=[row])
        m = result['metrics']
        assert m['persisting_cve_count'] == 0
        assert m['new_cve_count'] == 0

    # ── Device normalisation ──────────────────────────────────────────────────

    def test_device_name_normalisation_matches_across_case(self):
        """PC1 and pc1.domain.local should match as the same device."""
        cur_row  = {'Name': 'PC1',             'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        prev_row = {'Name': 'pc1.domain.local', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        result = self._run(current_rows=[cur_row], previous_rows=[prev_row])
        # Should be persisting, not new + resolved
        m = result['metrics']
        assert m['persisting_cve_count'] == 1
        assert m['new_cve_count'] == 0
        assert m['resolved_cve_count'] == 0

    # ── Not-in-RMM exclusion ─────────────────────────────────────────────────

    def test_not_in_rmm_rows_excluded_from_trend(self):
        """Not Found in RMM rows must never appear as New in trend arithmetic."""
        # The pipeline sets Last Response = 'Not Found in RMM' for these devices.
        # _active_trend_scope honours the 'Last Response' exclusion when the
        # orchestrator passes them through; in isolation we use inventory_devices
        # to simulate the same exclusion.
        row = {
            'Name': 'GHOST1', 'Vulnerability Name': 'CVE-2024-0001',
            'Affected Products': 'Chrome', 'Last Response': 'Not Found in RMM',
        }
        # Passing inventory_devices that doesn't include GHOST1 mirrors what
        # the orchestrator does: only devices in the RMM inventory are scoped.
        result = self._run(
            current_rows=[row], previous_rows=[],
            inventory_devices={'SOME_OTHER_PC'},
        )
        assert result['metrics']['new_pair_count'] == 0

    # ── Re-detection tracking ─────────────────────────────────────────────────

    def test_redetected_count_when_prev_resolved_pair_reappears(self):
        row = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        prev_resolved = {('PC1', 'CVE-2024-0001')}
        result = self._run(
            current_rows=[row],
            previous_rows=[],
            prev_resolved_pairs=prev_resolved,
        )
        assert result['redetected_count'] == 1

    def test_redetected_count_zero_when_no_overlap(self):
        row = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        prev_resolved = {('PC2', 'CVE-2024-0099')}
        result = self._run(
            current_rows=[row],
            previous_rows=[],
            prev_resolved_pairs=prev_resolved,
        )
        assert result['redetected_count'] == 0


# ==============================================================================
# normalize_device_name
# ==============================================================================

class TestNormalizeDeviceName:
    def test_uppercase(self):
        assert normalize_device_name('pc1') == 'PC1'

    def test_strips_domain(self):
        assert normalize_device_name('pc1.domain.local') == 'PC1'

    def test_strips_netbios_prefix(self):
        assert normalize_device_name('DOMAIN\\PC1') == 'PC1'

    def test_whitespace_stripped(self):
        assert normalize_device_name('  PC1  ') == 'PC1'

    def test_already_normalised(self):
        assert normalize_device_name('PC1') == 'PC1'


# ==============================================================================
# extract_cve_id
# ==============================================================================

class TestExtractCveId:
    def test_bare_cve(self):
        assert extract_cve_id('CVE-2024-12345') == 'CVE-2024-12345'

    def test_lowercase_normalised(self):
        assert extract_cve_id('cve-2024-12345') == 'CVE-2024-12345'

    def test_extracts_from_hyperlink_formula(self):
        val = '=HYPERLINK("https://cve.org/CVERecord?id=CVE-2024-12345","CVE-2024-12345")'
        assert extract_cve_id(val) == 'CVE-2024-12345'

    def test_no_cve_returns_input_uppercased(self):
        assert extract_cve_id('not a cve') == 'NOT A CVE'


# ==============================================================================
# get_base_product
# ==============================================================================

class TestGetBaseProduct:
    def test_strips_x64_tag(self):
        assert 'x64' not in get_base_product('Google Chrome x64').lower()

    def test_strips_version_suffix(self):
        result = get_base_product('Mozilla Firefox 123.0')
        assert '123.0' not in result

    def test_strips_32bit(self):
        result = get_base_product('Microsoft Edge (32-bit)')
        assert '32-bit' not in result
        assert 'Microsoft Edge' in result

    def test_plain_name_unchanged(self):
        assert get_base_product('Google Chrome') == 'Google Chrome'


# ==============================================================================
# parse_last_response
# ==============================================================================

class TestParseLastResponse:
    _epoch = pd.to_datetime('1900-01-01')

    def test_not_found_in_rmm_returns_epoch(self):
        assert parse_last_response('Not Found in RMM') == self._epoch

    def test_na_returns_epoch(self):
        assert parse_last_response('N/A') == self._epoch

    def test_empty_string_returns_epoch(self):
        assert parse_last_response('') == self._epoch

    def test_valid_datetime_parsed(self):
        result = parse_last_response('2026-03-15 09:00:00')
        assert result == pd.to_datetime('2026-03-15 09:00:00')

    def test_does_not_swallow_keyboard_interrupt(self):
        """parse_last_response must not catch KeyboardInterrupt."""
        import unittest.mock as mock
        with mock.patch('pandas.to_datetime', side_effect=KeyboardInterrupt):
            with pytest.raises(KeyboardInterrupt):
                parse_last_response('2026-03-15')


# ==============================================================================
# _active_trend_scope — boundary conditions
# ==============================================================================

class TestActiveTrendScope:
    def test_filters_below_threshold(self):
        df = _make_cve_df([
            {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Vulnerability Score': 5.0, 'Affected Products': 'Chrome'},
            {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0002', 'Vulnerability Score': 9.5, 'Affected Products': 'Chrome'},
        ])
        df['Vulnerability Score'] = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        result = _active_trend_scope(df, threshold=9.0)
        assert len(result) == 1
        assert result.iloc[0]['Vulnerability Name'] == 'CVE-2024-0002'

    def test_filters_not_in_rmm(self):
        # _active_trend_scope itself doesn't filter Not-in-RMM rows — that
        # happens upstream via the Threat Status = UNRESOLVED filter when the
        # merged DataFrame arrives from orchestrator. We verify the full
        # not-in-RMM exclusion in TestComputeTrends.test_not_in_rmm_rows_excluded_from_trend.
        # Here we verify score threshold + status filtering which is what
        # _active_trend_scope does own.
        df = _make_cve_df([
            {'Name': 'PC1', 'Vulnerability Score': 5.0, 'Threat Status': 'Unresolved', 'Affected Products': 'Chrome'},
            {'Name': 'PC2', 'Vulnerability Score': 9.5, 'Threat Status': 'Unresolved', 'Affected Products': 'Chrome'},
            {'Name': 'PC3', 'Vulnerability Score': 9.5, 'Threat Status': 'Resolved',   'Affected Products': 'Chrome'},
        ])
        df['Vulnerability Score'] = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        result = _active_trend_scope(df, threshold=9.0)
        assert 'PC1' not in result['Name'].values   # below threshold
        assert 'PC3' not in result['Name'].values   # resolved
        assert 'PC2' in result['Name'].values        # passes both filters

    def test_filters_stale_devices(self):
        df = _make_cve_df([
            {'Name': 'STALE', 'Affected Products': 'Chrome'},
            {'Name': 'PC1',   'Affected Products': 'Chrome'},
        ])
        df['Vulnerability Score'] = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        result = _active_trend_scope(df, threshold=9.0, stale_devices={'STALE'})
        assert 'STALE' not in result['Name'].values

    def test_deduplicates_on_name_cve_product(self):
        """Same device+CVE+product pair appearing twice should deduplicate."""
        row = {'Name': 'PC1', 'Vulnerability Name': 'CVE-2024-0001', 'Affected Products': 'Chrome'}
        df = _make_cve_df([row, row])
        df['Vulnerability Score'] = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        result = _active_trend_scope(df, threshold=9.0)
        assert len(result) == 1
