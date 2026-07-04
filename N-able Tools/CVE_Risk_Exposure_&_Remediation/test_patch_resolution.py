"""
test_patch_resolution.py — unit tests for Chrome/Edge patch resolution logic.

Each test corresponds to a real scenario that has caused incorrect results
in production. Test names document the scenario so failures are immediately
understandable without reading the code.

These tests target _vec_fixed_version / _vec_baseline / _vec_vcr / _vec_bc /
_vec_pes — the functions process_patch_match() actually calls. They were
promoted from nested closures inside process_patch_match to module level
specifically so they could be unit-tested directly (see CHANGELOG). This
file previously tested a parallel row-wise implementation
(_classify_version_check / _classify_resolution / _classify_baseline_compliance /
_resolve_fixed_version / _resolve_baseline) that looked equivalent but was
never actually called by process_patch_match — process_patch_match has used
its own separate vectorised logic since the v0.4-era rewrite. That meant 22
passing tests here gave false confidence about code that wasn't running in
production, while the real path (including a Status-column-collision bug
matching the one v0.4 fixed on the other side of the merge) went untested.
The dead functions have been deleted from data_pipeline.py; these tests now
cover their live replacements with the same scenarios.

One behaviour from the deleted _resolve_fixed_version is intentionally NOT
carried forward: an explicit 'Fixed Version' workbook-column override that
would win over both the per-CVE config rule and the baseline. Nothing in the
current pipeline ever populates a 'Fixed Version' column on the CVE
dataframe passed into process_patch_match, so that override path was dead
even when the row-wise function was live. If a workbook-column override is
still wanted, it needs to be re-added to _vec_fixed_version and wired to an
actual input column — flagging here rather than silently dropping it.

Run with:
    pytest test_patch_resolution.py -v

Author : Stu Villanti <s.villanti@kenstra.com>
"""

import os
import sys
import types
import pytest
import pandas as pd
from unittest.mock import patch

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
sys.modules.setdefault('config', _fake_config)

# Import the functions under test
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import data_pipeline  # noqa: E402
from data_pipeline import (  # noqa: E402
    _vec_fixed_version,
    _vec_baseline,
    _vec_vcr,
    _vec_bc,
    _vec_pes,
    INSTALLED_STATUSES,
)

# ---------------------------------------------------------------------------
# FIXED_VERSION_RULES is a single dict object bound once, whenever
# data_pipeline is FIRST imported anywhere in the pytest session — whichever
# test file's fake config module won the sys.modules race at that moment
# "wins" for every other test file too, since sys.modules.setdefault is a
# no-op once anything has already set 'config'. Rather than depend on this
# file happening to import first, patch.dict the exact rules this file
# needs around every test and restore afterward — the same pattern
# test_integration_summary_reconciliation.py already uses for its one
# monkeypatched rule. This makes every test below correct regardless of
# what other test files ran before it in the same session.
# ---------------------------------------------------------------------------

_REQUIRED_FIXED_VERSION_RULES = _fake_config.FIXED_VERSION_RULES


@pytest.fixture(autouse=True)
def _isolated_fixed_version_rules():
    with patch.dict(data_pipeline.FIXED_VERSION_RULES, _REQUIRED_FIXED_VERSION_RULES, clear=True):
        yield


# ---------------------------------------------------------------------------
# Single-row wrapper helpers around the batch-oriented _vec_* functions.
# process_patch_match calls these on whole columns at once; wrapping each
# call in a single-element list keeps the test scenarios exactly as
# readable as the old row-based versions were.
# ---------------------------------------------------------------------------

def _fixed_version(pk: str, vulnerability_name: str):
    fv_list, fs_list = _vec_fixed_version([pk], [vulnerability_name])
    return fv_list[0], fs_list[0]


def _baseline(pk: str):
    bl_list, bs_list = _vec_baseline([pk])
    return bl_list[0], bs_list[0]


def _vcr(status: str, matched_version: str, fixed_version: str) -> str:
    inst = status in INSTALLED_STATUSES
    return _vec_vcr([status], [matched_version], [fixed_version], [inst])[0]


def _bc(status: str, matched_version: str, baseline: str) -> str:
    inst = status in INSTALLED_STATUSES
    return _vec_bc([status], [matched_version], [baseline], [inst])[0]


def _pes(status: str, version_check_result: str, install_date, first_detected,
         date_published=None) -> str:
    inst_dt = pd.Timestamp(install_date) if install_date else pd.NaT
    cve_dates = [pd.Timestamp(d) for d in (first_detected, date_published) if d]
    cve_max = max(cve_dates) if cve_dates else pd.NaT
    return _vec_pes([status], [version_check_result], [inst_dt], [cve_max])[0]


# ---------------------------------------------------------------------------
# _vec_fixed_version / _vec_baseline — CVE-specific rule vs. rolling baseline
# ---------------------------------------------------------------------------

class TestFixedVersionAndBaseline:

    def test_fixed_version_returns_cve_specific_only(self):
        """
        _vec_fixed_version returns the CVE-specific rule only — not the baseline.
        Edge CVE-2026-5289 fixed at 146.0.3856.97; baseline is 147.0.3912.87.
        The function should return 146.0.3856.97 (CVE rule), NOT 147.0.3912.87
        (baseline). Baseline compliance is tracked separately by _vec_baseline.
        """
        fv, source = _fixed_version('edge', 'CVE-2026-5289')
        assert fv == '146.0.3856.97', (
            f"Expected per-CVE rule 146.0.3856.97, got {fv!r}. "
            f"_vec_fixed_version must return the CVE-specific rule, not the baseline."
        )
        assert 'config rule' in source.lower()

    def test_baseline_returns_rolling_baseline(self):
        """
        _vec_baseline returns the _baseline entry independently of any CVE rule.
        This is how 'Baseline Compliance' is computed separately from CVE patch status.
        """
        bl, bl_src = _baseline('edge')
        assert bl == '147.0.3912.87', f"Expected baseline 147.0.3912.87, got {bl!r}"
        assert 'baseline' in bl_src.lower()

    def test_cve_compliant_but_below_baseline_shows_both(self):
        """
        Core separation test: Chrome 147.0.7727.117 is:
          - CVE-compliant for CVE-2026-5858 (fixed at 147.0.7727.55)  → Patch confirmed
          - Below current baseline (148.0.7778.97)                     → Below baseline

        Both must be independently reportable without one overriding the other.
        """
        status, pv = 'Installed', '147.0.7727.117'
        fv, _ = _fixed_version('chrome', 'CVE-2026-5858')
        bl, _ = _baseline('chrome')

        vcr = _vcr(status, pv, fv)
        assert vcr == 'Version compliant', f"147.117 >= 147.55 → expected compliant, got {vcr!r}"

        pes = _pes(status, vcr, '2026-04-29', '2026-04-11')
        assert pes == 'Patch confirmed - pending rescan', (
            f"Version compliant + install after detection → should be Patch confirmed, got {pes!r}"
        )

        bc = _bc(status, pv, bl)
        assert bc == 'Below baseline', f"147.117 < 148.97 → should be Below baseline, got {bc!r}"

    def test_per_cve_wins_when_stricter_than_baseline(self):
        """
        If a CVE requires a version above the current baseline, per-CVE wins.
        Hypothetical: CVE requires Chrome 150.0, baseline is 148.0.
        Mutates data_pipeline.FIXED_VERSION_RULES directly (the module
        imports the dict at startup and the test fake is a different
        object) — safe because the autouse fixture above restores the
        dict's contents after every test regardless of what happens here.
        """
        import data_pipeline as _dp
        _dp.FIXED_VERSION_RULES.setdefault('chrome', {})
        _dp.FIXED_VERSION_RULES['chrome']['CVE-9999-99999'] = '150.0.0.0'
        _dp.FIXED_VERSION_RULES['chrome']['_baseline'] = '148.0.7778.97'

        fv, source = _fixed_version('chrome', 'CVE-9999-99999')
        assert fv == '150.0.0.0', f'Expected per-CVE rule 150.0.0.0, got {fv!r}'

    def test_no_per_cve_rule_returns_empty_fixed_version(self):
        """
        When there is no per-CVE rule, _vec_fixed_version returns empty.
        The baseline is NOT returned here — it is a separate concern in _vec_baseline.
        """
        fv, source = _fixed_version('chrome', 'CVE-2026-NOPERRULE')
        assert fv == '', (
            f"No per-CVE rule → _vec_fixed_version must return empty, got {fv!r}. "
            f"Use _vec_baseline for baseline tracking."
        )

    def test_baseline_returned_when_no_per_cve_rule(self):
        """_vec_baseline always returns the _baseline regardless of CVE rules."""
        bl, bl_src = _baseline('chrome')
        assert bl == '148.0.7778.97', f"Expected _baseline 148.0.7778.97, got {bl!r}"
        assert 'baseline' in bl_src.lower()


# ---------------------------------------------------------------------------
# _vec_vcr — Version Check Result
# ---------------------------------------------------------------------------

class TestVersionCheckResult:

    def test_version_compliant(self):
        assert _vcr('Installed', '147.0.7727.117', '147.0.7727.55') == 'Version compliant'

    def test_below_fixed_version(self):
        """Chrome 146.0.7680.165 is below fixed 146.0.7680.178."""
        assert _vcr('Installed', '146.0.7680.165', '146.0.7680.178') == 'Below fixed version'

    def test_pending_status_is_not_installed(self):
        """
        Key scenario: Status=Pending with a valid Matched Patch Version.
        The 'Discovered / Install Date' for Pending rows is the discovery date,
        not an install date. Must NOT be treated as installed.
        """
        result = _vcr('Pending', '147.0.7727.138', '147.0.7727.55')
        assert result == 'Patch not yet installed', (
            f"Pending status must return 'Patch not yet installed', got {result!r}"
        )

    def test_missing_status_is_not_installed(self):
        result = _vcr('Missing', '147.0.7727.138', '147.0.7727.55')
        assert result == 'Patch not yet installed'

    def test_no_fixed_baseline(self):
        assert _vcr('Installed', '147.0.7727.138', '') == 'Installed version found - no fixed baseline'


# ---------------------------------------------------------------------------
# _vec_pes — Patch Evidence Status (end-to-end resolution timing)
# ---------------------------------------------------------------------------

class TestPatchEvidenceStatus:

    def test_chrome_patched_after_detection_version_compliant(self):
        """
        Chrome 147.0.7727.117 installed 29-Apr, first detected 11-Apr.
        Version compliant and install post-dates detection → Patch confirmed.
        """
        result = _pes('Installed', 'Version compliant', '2026-04-29', '2026-04-11')
        assert result == 'Patch confirmed - pending rescan', (
            f"Expected 'Patch confirmed - pending rescan', got {result!r}"
        )

    def test_pending_with_newer_discovered_date_is_unresolved(self):
        """
        Core Pending confusion: Discovered / Install Date > First detected but Status=Pending.
        Must remain Unresolved regardless of the date comparison.
        """
        result = _pes('Pending', 'Patch not yet installed', '2026-04-30', '2026-04-11')
        assert result == 'Unresolved', (
            f"Pending status with newer discovered date must be Unresolved, got {result!r}"
        )

    def test_below_fixed_version_is_unresolved(self):
        """Chrome 146.0.7680.165 installed, fixed 146.0.7680.178 → Unresolved."""
        result = _pes('Installed', 'Below fixed version', '2026-03-20', '2026-04-02')
        assert result == 'Unresolved'

    def test_edge_below_per_cve_fixed_is_unresolved(self):
        """
        Edge 146.0.3856.78 installed, CVE fixed 146.0.3856.97 → Unresolved.
        This was the original bug where a stale per-CVE rule allowed a false resolve.
        """
        result = _pes('Installed', 'Below fixed version', '2026-03-30', '2026-04-04')
        assert result == 'Unresolved'

    def test_edge_compliant_after_detection(self):
        """
        Edge: 147.0.3912.86 installed 27-Apr, fixed 146.0.3856.97, detected 4-Apr.
        Version compliant and install post-dates detection → Patch confirmed.
        """
        result = _pes('Installed', 'Version compliant', '2026-04-27', '2026-04-04')
        assert result == 'Patch confirmed - pending rescan'

    def test_install_predating_detection_is_unresolved_even_when_version_compliant(self):
        """
        If the patch install date is BEFORE the CVE was first detected, the install
        cannot be evidence that the CVE was remediated — it predates the vulnerability.
        Even if the version is compliant, the timing check must fail.

        This prevents pre-existing browser installs (e.g. Chrome 147 already on device
        when CVE is published) from being auto-resolved without a fresh scan.
        """
        result = _pes('Installed', 'Version compliant', '2026-04-01', '2026-04-11')
        # Install date (Apr 1) predates first detection (Apr 11) → Unresolved
        assert result == 'Unresolved', (
            f"Pre-detection install must remain Unresolved, got {result!r}"
        )

    def test_install_predating_detection_without_version_check_is_unresolved(self):
        """
        Timing-only path (no version data): install before first detection
        should not be accepted as evidence.
        """
        result = _pes('Installed', 'Fixed baseline known - installed version not found',
                      '2026-04-01', '2026-04-11')
        assert result == 'Unresolved'

    def test_reboot_required_compliant_resolves(self):
        """Reboot Required is treated as installed for version checking."""
        result = _pes('Reboot Required', 'Version compliant', '2026-04-29', '2026-04-11')
        assert result == 'Patch confirmed - pending rescan'

    def test_uses_date_published_when_later_than_first_detected(self):
        """
        Patch Evidence Status must compare against the LATEST of First detected /
        Date Published, not just First detected — an install that postdates
        detection but predates a later-published CVE record is still unresolved.
        """
        result = _pes('Installed', 'Version compliant', '2026-04-15',
                      first_detected='2026-04-11', date_published='2026-04-20')
        assert result == 'Unresolved', (
            "Install (Apr 15) predates Date Published (Apr 20), even though it "
            "postdates First detected (Apr 11) — the later date must win"
        )


# ==============================================================================
# Product trend tests
# ==============================================================================

class TestProductTrend:
    """Tests for compute_trends product_trend construction."""

    def _make_cve_df(self, rows):
        """Build a minimal CVE DataFrame for trend testing."""
        import data_pipeline as dp
        df = pd.DataFrame(rows)
        df['Vulnerability Score'] = pd.to_numeric(df['Vulnerability Score'], errors='coerce')
        df['_Name_Key']  = df['Name'].apply(dp.normalize_device_name)
        df['_CVE_Key']   = df['Vulnerability Name']
        df['Base Product'] = df['Affected Products'].apply(dp.get_base_product)
        return df

    def test_new_product_appears_in_trend_with_prev_zero(self):
        """
        A product present in current but absent from previous must appear in
        product_trend with Previous = 0.  This was the Edge bug: Edge appeared
        this month for the first time and was silently omitted from the Trend
        Summary Top 10 because it wasn't in common_products.
        """
        import data_pipeline as _dp

        cur = self._make_cve_df([
            {'Name': 'D1', 'Vulnerability Name': 'CVE-2026-0001',
             'Affected Products': 'Microsoft Edge 80+', 'Vulnerability Score': 9.6,
             'Last Response': '2026-04-01'},
            {'Name': 'D2', 'Vulnerability Name': 'CVE-2026-0001',
             'Affected Products': 'Microsoft Edge 80+', 'Vulnerability Score': 9.6,
             'Last Response': '2026-04-01'},
            {'Name': 'D1', 'Vulnerability Name': 'CVE-2026-0002',
             'Affected Products': 'Google Chrome', 'Vulnerability Score': 9.6,
             'Last Response': '2026-04-01'},
        ])
        prev = self._make_cve_df([
            {'Name': 'D1', 'Vulnerability Name': 'CVE-2026-0002',
             'Affected Products': 'Google Chrome', 'Vulnerability Score': 9.6,
             'Last Response': '2026-03-01'},
        ])

        result = _dp.compute_trends(cur, prev, threshold=9.0)
        pt = result['product_trend']

        # Edge must appear even though it wasn't in the previous report
        assert 'Microsoft Edge' in pt.index, (
            "Microsoft Edge must appear in product_trend even when absent from previous report. "
            "It was absent before, so Previous should be 0."
        )
        edge_row = pt.loc['Microsoft Edge']
        assert edge_row['Current']  == 2, f"Edge: expected 2 devices, got {edge_row['Current']}"
        assert edge_row['Previous'] == 0, f"Edge: expected Previous=0 (new product), got {edge_row['Previous']}"
        assert edge_row['Change']   == 2, f"Edge: expected Change=+2, got {edge_row['Change']}"

    def test_existing_product_shows_delta(self):
        """Products present in both periods show correct Prev/Current/Change."""
        import data_pipeline as _dp

        def _rows(devices, cve='CVE-2026-0001', product='Google Chrome', score=9.6, date='2026-04-01'):
            return [{'Name': d, 'Vulnerability Name': cve, 'Affected Products': product,
                     'Vulnerability Score': score, 'Last Response': date}
                    for d in devices]

        cur  = self._make_cve_df(_rows(['D1','D2','D3']))
        prev = self._make_cve_df(_rows(['D1','D2'], date='2026-03-01'))

        pt = _dp.compute_trends(cur, prev, threshold=9.0)['product_trend']
        chrome = pt.loc['Google Chrome']
        assert chrome['Current']  == 3
        assert chrome['Previous'] == 2
        assert chrome['Change']   == 1