"""
test_patch_evidence_override.py — regression tests for using the patch report
to resolve CVEs that N-able's detections export still reports as UNRESOLVED.

N-able's scanner can lag a patch install by a full rescan cycle, so a CVE can
still read UNRESOLVED in the detections export well after the update actually
landed. Three separate defects sat on the path that was meant to correct for
this, and every one of them had to be fixed before the patch report could
contribute anything at all:

  1. resolution.reconcile_patch_evidence() — the v0.24 "UNRESOLVED always
     wins" strip removed exactly the pairs the patch report exists to
     contribute (patch evidence only ever changes an outcome where the scanner
     says UNRESOLVED), so its net effect on resolution status was zero.
  2. data_pipeline.process_patch_match() — indexed ['First detected',
     'Date Published'] unconditionally and raised KeyError on exports carrying
     neither, killing the whole patch run before a pair was built.
  3. cve_lookup.enrich_date_published() — nothing in the pipeline ever created
     'Date Published', so _vec_pes' anchor date was always NaT and every row
     fell through to 'Unresolved' even when the crash was avoided.

Run with: pytest test_patch_evidence_override.py -v

Author : Stu Villanti <s.villanti@kenstra.com>
"""

import os
import sys
import types

import pytest
from unittest.mock import patch

os.environ.setdefault('PYTEST_CURRENT_TEST', 'bootstrap')

# ---------------------------------------------------------------------------
# Config stubbing — same pattern (and same setdefault rationale) as
# test_merge_patch_match.py / test_resolution.py, so this file runs standalone
# without a real config.json and cooperates with whichever test file won the
# sys.modules race earlier in the session.
# ---------------------------------------------------------------------------

import re as _re
_fake_config = types.ModuleType('config')
_fake_config.CVE_PATTERN = _re.compile(r'(CVE-\d{4}-\d{4,7})', _re.IGNORECASE)
_fake_config.PRODUCT_MAP = [
    ('google chrome',   'chrome'),
    ('mozilla firefox', 'firefox'),
    ('microsoft edge',  'edge'),
]
_fake_config.FIXED_VERSION_RULES = {
    'chrome': {
        '_baseline':     '148.0.0.0',
        'CVE-2026-1234': '147.0.0.0',
    },
}
_fake_config.STATUS_RANK = {
    'Installed': 6, 'Reboot Required': 5, 'Installing': 4,
    'Pending': 3,   'Missing': 2,         'Failed': 1,
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
from data_pipeline import process_patch_match  # noqa: E402
from resolution import reconcile_patch_evidence  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated_fixed_version_rules():
    with patch.dict(data_pipeline.FIXED_VERSION_RULES,
                    _fake_config.FIXED_VERSION_RULES, clear=True):
        yield


# ===========================================================================
# reconcile_patch_evidence() — who wins when the two sources disagree
# ===========================================================================

class TestReconcilePatchEvidence:

    PAIRS = {('WS01', 'CVE-2026-1234', 'chrome')}
    UNRESOLVED = {('WS01', 'CVE-2026-1234')}

    def test_scanner_wins_by_default(self):
        """Default behaviour is unchanged from v0.24 — contested pairs dropped."""
        pairs, overrides = reconcile_patch_evidence(self.PAIRS, self.UNRESOLVED)

        assert pairs == set(), "contested pair must be dropped when not trusting patch evidence"
        assert overrides == 0

    def test_patch_evidence_wins_when_trusted(self):
        """
        The bug: this pair is the ONLY kind that matters (patch says patched,
        scanner still says UNRESOLVED) and it was unconditionally deleted.
        """
        pairs, overrides = reconcile_patch_evidence(
            self.PAIRS, self.UNRESOLVED, trust_patch_evidence=True)

        assert pairs == self.PAIRS, "patch evidence must survive a stale UNRESOLVED"
        assert overrides == 1

    def test_uncontested_pairs_survive_either_way(self):
        """A pair the scanner doesn't contradict is untouched, and isn't an override."""
        for trust in (False, True):
            pairs, overrides = reconcile_patch_evidence(
                self.PAIRS, {('WS02', 'CVE-2026-9999')}, trust_patch_evidence=trust)

            assert pairs == self.PAIRS
            assert overrides == 0, "uncontested pair is not an override — nothing was overridden"

    def test_contest_is_matched_on_2_tuple_not_product(self):
        """
        Product-string drift must never let a contested pair slip through as
        resolved — the scanner's UNRESOLVED has no product key to compare.
        """
        pairs, _ = reconcile_patch_evidence(
            {('WS01', 'CVE-2026-1234', 'chrome'),
             ('WS01', 'CVE-2026-1234', 'chrome-for-business')},
            self.UNRESOLVED,
        )

        assert pairs == set(), "both product variants of a contested pair must drop"

    def test_override_count_is_distinct_device_cve(self):
        """One CVE spanning several product keys on a device counts once."""
        _, overrides = reconcile_patch_evidence(
            {('WS01', 'CVE-2026-1234', 'chrome'),
             ('WS01', 'CVE-2026-1234', 'edge')},
            self.UNRESOLVED, trust_patch_evidence=True,
        )

        assert overrides == 1

    def test_empty_inputs_are_safe(self):
        assert reconcile_patch_evidence(None, None) == (set(), 0)
        assert reconcile_patch_evidence(set(), self.UNRESOLVED) == (set(), 0)
        assert reconcile_patch_evidence(self.PAIRS, set()) == (self.PAIRS, 0)

    def test_does_not_mutate_caller_set(self):
        original = set(self.PAIRS)
        reconcile_patch_evidence(original, self.UNRESOLVED)

        assert original == self.PAIRS, "input set must not be mutated in place"


# ===========================================================================
# process_patch_match() — missing CVE date columns
# ===========================================================================

def _patch_csv(tmp_path, install_date='1-Mar-2026', status='Installed',
               patch_name='Google Chrome 149.0.0.0'):
    p = tmp_path / 'patch.csv'
    pd.DataFrame([{
        'Client': 'Acme', 'Site': 'HQ', 'Device': 'WS01',
        'Status': status, 'Patch': patch_name,
        'Discovered / Install Date': install_date,
    }]).to_csv(p, index=False)
    return str(p)


def _cve_frame(**extra):
    row = {
        'Name': 'WS01',
        'Vulnerability Name': 'CVE-2026-1234',
        'Affected Products': 'Google Chrome',
        'Threat Status': 'UNRESOLVED',
        'Vulnerability Score': 9.8,
        'Customer': 'Acme',
        'Site': 'HQ',
    }
    row.update(extra)
    return pd.DataFrame([row])


class TestMissingCveDateColumns:
    """
    The real N-able detections export ships 'Last scanned' and neither
    'First detected' nor 'Date Published'. Selecting both unconditionally
    raised KeyError and took down the entire patch run.
    """

    def test_no_date_columns_does_not_crash(self, tmp_path):
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path), _cve_frame(), min_score=0.0)

        assert len(p_full) == 1
        assert 'Patch Evidence Status' in p_full.columns

    def test_no_date_columns_fails_closed(self, tmp_path):
        """
        With no anchor date there is no evidence the patch postdates the CVE,
        so the row must NOT be claimed as patch-confirmed.
        """
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path), _cve_frame(), min_score=0.0)

        assert p_full.iloc[0]['Patch Evidence Status'] == 'Unresolved'

    def test_only_first_detected_present(self, tmp_path):
        """Either column alone must be enough — previously both were required."""
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path), _cve_frame(**{'First detected': '2026-02-01'}),
            min_score=0.0)

        assert p_full.iloc[0]['Patch Evidence Status'] == 'Patch confirmed - pending rescan'

    def test_only_date_published_present(self, tmp_path):
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path), _cve_frame(**{'Date Published': '2026-01-15'}),
            min_score=0.0)

        assert p_full.iloc[0]['Patch Evidence Status'] == 'Patch confirmed - pending rescan'

    def test_install_predating_cve_is_not_evidence(self, tmp_path):
        """An install that predates the CVE cannot be what fixed it."""
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path, install_date='1-Jan-2025'),
            _cve_frame(**{'Date Published': '2026-01-15'}), min_score=0.0)

        assert p_full.iloc[0]['Patch Evidence Status'] == 'Unresolved'


# ===========================================================================
# End-to-end: the reported bug
# ===========================================================================

class TestStaleScannerEndToEnd:

    def test_patched_device_resolves_despite_stale_unresolved(self, tmp_path):
        """
        The whole point. Device patched 1-Mar for a CVE published 15-Jan; the
        detections export still says UNRESOLVED because N-able hasn't rescanned.
        With trust_patch_evidence on, the CVE must come out resolved.
        """
        from data_pipeline import normalize_device_name, extract_cve_id, _detect_product

        cve = _cve_frame(**{'Date Published': '2026-01-15'})
        _ov, p_full, _raw, _t, _f = process_patch_match(
            _patch_csv(tmp_path), cve, min_score=0.0)

        confirmed = p_full[p_full['Patch Evidence Status'] == 'Patch confirmed - pending rescan']
        assert not confirmed.empty, "patch evidence must be produced in the first place"

        pairs = set(zip(
            confirmed['Name'].apply(normalize_device_name),
            confirmed['Vulnerability Name'].apply(extract_cve_id),
            confirmed['Affected Products'].astype(str).apply(_detect_product),
        ))

        # The scanner still reports it unresolved — this is the stale status.
        unresolved_2d = set(zip(
            cve['Name'].apply(normalize_device_name),
            cve['Vulnerability Name'].apply(extract_cve_id),
        ))

        kept, overrides = reconcile_patch_evidence(
            pairs, unresolved_2d, trust_patch_evidence=True)
        assert kept == pairs and overrides == 1

        dropped, no_overrides = reconcile_patch_evidence(pairs, unresolved_2d)
        assert dropped == set() and no_overrides == 0
