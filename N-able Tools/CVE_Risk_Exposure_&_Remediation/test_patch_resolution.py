"""
tests/test_patch_resolution.py

Unit tests for the Chrome/Edge patch resolution logic in data_pipeline.py.

Each test corresponds to a real scenario that has caused incorrect results
in production.  The test name documents the scenario so failures are
immediately understandable without reading the code.

Run with:
    pytest tests/test_patch_resolution.py -v
"""

import os
import sys
import types
import pytest
import pandas as pd

# ---------------------------------------------------------------------------
# Bootstrap: stub config.json loading so tests run without the full project
# ---------------------------------------------------------------------------

os.environ.setdefault('PYTEST_CURRENT_TEST', 'bootstrap')

# Provide a minimal config module before importing data_pipeline
_fake_config = types.ModuleType('config')
_fake_config.CVE_PATTERN = __import__('re').compile(r'(CVE-\d{4}-\d{4,7})', __import__('re').IGNORECASE)
_fake_config.PRODUCT_MAP = [
    ('google chrome', 'chrome'),
    ('mozilla firefox', 'firefox'),
    ('microsoft edge', 'edge'),
    ('chromium', 'chrome'),
]
_fake_config.FIXED_VERSION_RULES = {
    'chrome': {
        '_baseline':      '148.0.7778.97',
        'CVE-2026-5858':  '147.0.7727.55',
        'CVE-2026-5859':  '147.0.7727.55',
        'CVE-2026-5288':  '146.0.7680.178',
        'CVE-2026-5289':  '146.0.7680.178',
        'CVE-2026-5290':  '146.0.7680.178',
    },
    'edge': {
        '_baseline':      '147.0.3912.87',
        'CVE-2026-5289':  '146.0.3856.97',
        'CVE-2026-5290':  '146.0.3856.97',
        'CVE-2026-5288':  '146.0.3856.97',
    },
    'firefox': {
        '_baseline':      '150.0.1',
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
sys.modules['config'] = _fake_config

# Import the functions under test
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from data_pipeline import (
    _classify_version_check,
    _classify_resolution,
    _resolve_fixed_version,
    FIXED_VERSION_RULES,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_row(
    status: str,
    matched_version: str,
    fixed_version: str,
    install_date,
    first_detected,
    vulnerability_name: str = 'CVE-2026-5858',
    affected_products: str = 'Google Chrome',
    pk: str = 'chrome',
) -> pd.Series:
    """Build a minimal patch-match row for testing resolution functions."""
    return pd.Series({
        'Status':                       status,
        'Matched Patch Version':        matched_version,
        'Fixed Version Used':           fixed_version,
        'Patch Install Date':           pd.Timestamp(install_date) if install_date else pd.NaT,
        'First detected':               pd.Timestamp(first_detected) if first_detected else pd.NaT,
        'Date Published':               pd.NaT,
        'Vulnerability Name':           vulnerability_name,
        'Affected Products':            affected_products,
        '_pk':                          pk,
        'Resolved (from Patch Report)': '',  # not used in these tests directly
    })


# ---------------------------------------------------------------------------
# _resolve_fixed_version — max(per_cve, baseline) logic
# ---------------------------------------------------------------------------

class TestResolveFixedVersion:

    def test_baseline_supersedes_stale_per_cve_rule(self):
        """
        Edge CVE-2026-5289: per-CVE rule is 146.0.3856.97, baseline is 147.0.3912.87.
        The baseline is stricter and must win.
        """
        row = _make_row(
            status='Installed', matched_version='146.0.3856.78',
            fixed_version='', install_date='2026-03-30', first_detected='2026-04-04',
            vulnerability_name='CVE-2026-5289', affected_products='Microsoft Edge 80+',
            pk='edge',
        )
        fv, source = _resolve_fixed_version(row)
        assert fv == '147.0.3912.87', f"Expected baseline 147.0.3912.87, got {fv}"
        assert 'supersedes' in source.lower() or 'baseline' in source.lower()

    def test_per_cve_wins_when_stricter_than_baseline(self):
        """
        If a CVE requires a version above the current baseline, per-CVE wins.
        Hypothetical: CVE requires Chrome 150.0, baseline is 148.0.
        Must patch dp.FIXED_VERSION_RULES directly — the module imports the dict
        at startup and the test fake is a different object.
        """
        import data_pipeline as _dp
        _dp.FIXED_VERSION_RULES.setdefault('chrome', {})
        _dp.FIXED_VERSION_RULES['chrome']['CVE-9999-99999'] = '150.0.0.0'
        old_baseline = _dp.FIXED_VERSION_RULES['chrome'].get('_baseline', '148.0.7778.97')
        _dp.FIXED_VERSION_RULES['chrome']['_baseline'] = '148.0.7778.97'
        try:
            row = _make_row(
                status='Installed', matched_version='149.0.0.0',
                fixed_version='', install_date='2026-05-01', first_detected='2026-04-01',
                vulnerability_name='CVE-9999-99999', pk='chrome',
            )
            fv, source = _resolve_fixed_version(row)
            assert fv == '150.0.0.0', f'Expected per-CVE rule 150.0.0.0, got ' + repr(fv)
        finally:
            del _dp.FIXED_VERSION_RULES['chrome']['CVE-9999-99999']
            _dp.FIXED_VERSION_RULES['chrome']['_baseline'] = old_baseline

    def test_baseline_only_when_no_per_cve_rule(self):
        """When there is no per-CVE rule, the baseline is returned."""
        row = _make_row(
            status='Installed', matched_version='147.0.7727.117',
            fixed_version='', install_date='2026-04-29', first_detected='2026-04-11',
            vulnerability_name='CVE-2026-NOPERRULE', pk='chrome',
        )
        fv, source = _resolve_fixed_version(row)
        assert fv == '148.0.7778.97'
        assert 'baseline' in source.lower()

    def test_explicit_workbook_column_always_wins(self):
        """If 'Fixed Version' column is present in the row, it overrides everything."""
        row = _make_row(
            status='Installed', matched_version='147.0.7727.117',
            fixed_version='', install_date='2026-04-29', first_detected='2026-04-11',
            vulnerability_name='CVE-2026-5858', pk='chrome',
        )
        row['Fixed Version'] = '999.0.0.0'
        fv, source = _resolve_fixed_version(row)
        assert fv == '999.0.0.0'
        assert source == 'CVE workbook column'


# ---------------------------------------------------------------------------
# _classify_version_check
# ---------------------------------------------------------------------------

class TestClassifyVersionCheck:

    def test_version_compliant(self):
        row = _make_row('Installed', '147.0.7727.117', '147.0.7727.55', '2026-04-29', '2026-04-11')
        assert _classify_version_check(row) == 'Version compliant'

    def test_below_fixed_version(self):
        """Chrome 146.0.7680.165 is below fixed 146.0.7680.178."""
        row = _make_row('Installed', '146.0.7680.165', '146.0.7680.178', '2026-03-20', '2026-04-02')
        assert _classify_version_check(row) == 'Below fixed version'

    def test_pending_status_is_not_installed(self):
        """
        Key scenario: Status=Pending with a valid Matched Patch Version.
        The 'Discovered / Install Date' for Pending rows is the discovery date,
        not an install date.  Must NOT be treated as installed.
        """
        row = _make_row('Pending', '147.0.7727.138', '147.0.7727.55', '2026-04-30', '2026-04-11')
        result = _classify_version_check(row)
        assert result == 'Patch not yet installed', (
            f"Pending status must return 'Patch not yet installed', got {result!r}"
        )

    def test_missing_status_is_not_installed(self):
        row = _make_row('Missing', '147.0.7727.138', '147.0.7727.55', '2026-04-30', '2026-04-11')
        result = _classify_version_check(row)
        assert result == 'Patch not yet installed'

    def test_no_fixed_baseline(self):
        row = _make_row('Installed', '147.0.7727.138', '', '2026-04-29', '2026-04-11')
        assert _classify_version_check(row) == 'Installed version found - no fixed baseline'


# ---------------------------------------------------------------------------
# _classify_resolution (end-to-end patch evidence status)
# ---------------------------------------------------------------------------

class TestClassifyResolution:

    def test_chrome_patched_after_detection_version_compliant(self):
        """
        CVLT295 scenario: Chrome 147.0.7727.117 installed 29-Apr, first detected 11-Apr.
        Version >= fixed (147.55) and install post-dates detection → Patch confirmed.
        """
        row = _make_row(
            status='Installed', matched_version='147.0.7727.117',
            fixed_version='147.0.7727.55', install_date='2026-04-29',
            first_detected='2026-04-11', pk='chrome',
        )
        row['Version Check Result'] = 'Version compliant'
        result = _classify_resolution(row)
        assert result == 'Patch confirmed - pending rescan', (
            f"Expected 'Patch confirmed - pending rescan', got {result!r}"
        )

    def test_chrome_pending_with_newer_discovered_date_is_unresolved(self):
        """
        Core Pending confusion: Discovered / Install Date > First detected but Status=Pending.
        Must remain Unresolved.
        """
        row = _make_row(
            status='Pending', matched_version='147.0.7727.138',
            fixed_version='147.0.7727.55', install_date='2026-04-30',
            first_detected='2026-04-11', pk='chrome',
        )
        row['Version Check Result'] = 'Patch not yet installed'
        result = _classify_resolution(row)
        assert result == 'Unresolved', (
            f"Pending status with newer discovered date must be Unresolved, got {result!r}"
        )

    def test_chrome_below_fixed_version_is_unresolved(self):
        """CVLT048: Chrome 146.0.7680.165 installed, fixed 146.0.7680.178 → Unresolved."""
        row = _make_row(
            status='Installed', matched_version='146.0.7680.165',
            fixed_version='146.0.7680.178', install_date='2026-03-20',
            first_detected='2026-04-02', pk='chrome',
        )
        row['Version Check Result'] = 'Below fixed version'
        result = _classify_resolution(row)
        assert result == 'Unresolved'

    def test_edge_below_per_cve_fixed_is_unresolved(self):
        """
        CVLT265: Edge 146.0.3856.78 installed, CVE fixed 146.0.3856.97 → Unresolved.
        This was the original bug where the stale per-CVE rule (136) allowed a false resolve.
        """
        row = _make_row(
            status='Installed', matched_version='146.0.3856.78',
            fixed_version='146.0.3856.97', install_date='2026-03-30',
            first_detected='2026-04-04',
            vulnerability_name='CVE-2026-5289',
            affected_products='Microsoft Edge 80+', pk='edge',
        )
        row['Version Check Result'] = 'Below fixed version'
        result = _classify_resolution(row)
        assert result == 'Unresolved'

    def test_edge_compliant_after_detection(self):
        """
        CVLT295 Edge: 147.0.3912.86 installed 27-Apr, fixed 146.0.3856.97, detected 4-Apr.
        Version compliant and install post-dates detection → Patch confirmed.
        """
        row = _make_row(
            status='Installed', matched_version='147.0.3912.86',
            fixed_version='146.0.3856.97', install_date='2026-04-27',
            first_detected='2026-04-04',
            vulnerability_name='CVE-2026-5289',
            affected_products='Microsoft Edge 80+', pk='edge',
        )
        row['Version Check Result'] = 'Version compliant'
        result = _classify_resolution(row)
        assert result == 'Patch confirmed - pending rescan'

    def test_install_predating_detection_is_unresolved_even_when_version_compliant(self):
        """
        If the patch install date is BEFORE the CVE was first detected, the install
        cannot be evidence that the CVE was remediated — it predates the vulnerability.
        Even if the version is compliant, the timing check must fail.

        This prevents pre-existing browser installs (e.g. Chrome 147 already on device
        when CVE is published) from being auto-resolved without a fresh scan.
        """
        row = _make_row(
            status='Installed', matched_version='147.0.7727.117',
            fixed_version='147.0.7727.55', install_date='2026-04-01',
            first_detected='2026-04-11', pk='chrome',
        )
        row['Version Check Result'] = 'Version compliant'
        result = _classify_resolution(row)
        # Install date (Apr 1) predates first detection (Apr 11) → Unresolved
        assert result == 'Unresolved', (
            f"Pre-detection install must remain Unresolved, got {result!r}"
        )

    def test_install_predating_detection_without_version_check_is_unresolved(self):
        """
        Timing-only path (no version data): install before first detection
        should not be accepted as evidence.
        """
        row = _make_row(
            status='Installed', matched_version='',
            fixed_version='147.0.7727.55', install_date='2026-04-01',
            first_detected='2026-04-11', pk='chrome',
        )
        row['Version Check Result'] = 'Fixed baseline known - installed version not found'
        result = _classify_resolution(row)
        assert result == 'Unresolved'

    def test_reboot_required_compliant_resolves(self):
        """Reboot Required is treated as installed for version checking."""
        row = _make_row(
            status='Reboot Required', matched_version='147.0.7727.117',
            fixed_version='147.0.7727.55', install_date='2026-04-29',
            first_detected='2026-04-11', pk='chrome',
        )
        row['Version Check Result'] = 'Version compliant'
        result = _classify_resolution(row)
        assert result == 'Patch confirmed - pending rescan'
