"""
test_non_cve_patch_filter.py — dropping non-CVE maintenance patches.

The N-able patch export lists antivirus/Defender signature updates and driver
updates that do not correspond to any CVE the scanner tracks. When such a
patch shares a product key with a CVE row it gets matched and mis-classified
as "Patched but still detected (rescan required)". _is_non_cve_patch flags
these, and process_patch_match drops them before matching — while keeping
Windows/OS cumulative and .NET/Security KB rollups, which DO fix CVEs.
"""
import pandas as pd
import pytest

import data_pipeline as dp


KEEP = [
    "2021-01 Update for Windows Server 2019 for x64-based Systems (KB4589208)",
    "2025-10 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5066835)",
    "2025-10 Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 11 (KB5065789)",
    "2024-11 .NET 6.0.36 Update for x64 Client (KB5047486)",
    "2025-11 Security Update (KB5068861) (26100.7171)",
    "Adobe Acrobat DC (64-bit) (x64) 25.1.20997",
    "Google Chrome (x64) 150.0.7871.18",
    "Mozilla Firefox 128.0 (x64 en-US)",
]

DROP = [
    "Security Intelligence Update for Microsoft Defender Antivirus - KB2267602 (1.4.1234.0)",
    "Dell, Inc. Driver Update (25.0.1.1)",
    "Dell, Inc. Firmware Driver Update (0.1.23.0)",
    "Intel Corporation Bluetooth Driver Update (24.40.0.3)",
    "Realtek Semiconductor Corp. - Extension - Audio Driver Update (6.0.9414.1)",
]


class TestClassifier:
    @pytest.mark.parametrize("name", KEEP)
    def test_keeps_cve_patches(self, name):
        assert dp._is_non_cve_patch(name) is False, f"wrongly flagged: {name}"

    @pytest.mark.parametrize("name", DROP)
    def test_drops_non_cve_patches(self, name):
        assert dp._is_non_cve_patch(name) is True, f"failed to flag: {name}"

    def test_none_is_not_flagged(self):
        assert dp._is_non_cve_patch(None) is False


def _patch_df(rows):
    return pd.DataFrame([
        {'Client': 'Alpha Co', 'Site': 'HQ', 'Device': dev, 'Status': 'Installed',
         'Patch': patch, 'Discovered / Install Date': '2026-07-10'}
        for dev, patch in rows
    ])


class TestProcessPatchMatchFilter:
    def test_driver_patch_not_matched_to_cve(self, tmp_path):
        # A device whose only patch is a driver update, against a Chrome CVE.
        # With the filter, the driver row is dropped, so the CVE must NOT read
        # as matched/installed — it falls through to a coverage gap.
        patch = _patch_df([('DEV-1', 'Intel Corporation Bluetooth Driver Update (24.40.0.3)')])
        pfile = tmp_path / 'patch.csv'; patch.to_csv(pfile, index=False)

        cve = pd.DataFrame([{
            'Customer': 'Alpha Co', 'Site': 'HQ', 'Name': 'DEV-1',
            'Vulnerability Name': 'CVE-2026-0001',
            'Affected Products': 'Google Chrome 120.0 x64',
            'Vulnerability Score': 9.5, 'Threat Status': 'UNRESOLVED',
        }])
        _, full, _, _, _ = dp.process_patch_match(str(pfile), cve, min_score=9.0)
        result = full.iloc[0]['Patch Match Result']
        # Device is still in the report (it has other rows? no — its only row
        # was dropped), so it should read as a coverage gap, never "installed".
        assert 'installed' not in str(result).lower()

    def test_real_patch_still_matches_after_filter(self, tmp_path):
        # Same device has BOTH a driver update (dropped) and a real Chrome
        # patch (kept) — the CVE must still match the real patch.
        patch = _patch_df([
            ('DEV-1', 'Intel Corporation Bluetooth Driver Update (24.40.0.3)'),
            ('DEV-1', 'Google Chrome (x64) 150.0.7871.18'),
        ])
        pfile = tmp_path / 'patch.csv'; patch.to_csv(pfile, index=False)
        cve = pd.DataFrame([{
            'Customer': 'Alpha Co', 'Site': 'HQ', 'Name': 'DEV-1',
            'Vulnerability Name': 'CVE-2026-0001',
            'Affected Products': 'Google Chrome 120.0 x64',
            'Vulnerability Score': 9.5, 'Threat Status': 'UNRESOLVED',
        }])
        _, full, _, _, _ = dp.process_patch_match(str(pfile), cve, min_score=9.0)
        assert 'installed' in str(full.iloc[0]['Patch Match Result']).lower()

    def test_defender_signature_dropped(self, tmp_path):
        patch = _patch_df([
            ('DEV-1', 'Security Intelligence Update for Microsoft Defender Antivirus - KB2267602 (1.4.1.0)'),
            ('DEV-1', 'Google Chrome (x64) 150.0.7871.18'),
        ])
        pfile = tmp_path / 'patch.csv'; patch.to_csv(pfile, index=False)
        loaded = dp.load_data(str(pfile))
        # Prove both rows load; the filter is applied inside process_patch_match.
        assert len(loaded) == 2
        cve = pd.DataFrame([{
            'Customer': 'Alpha Co', 'Site': 'HQ', 'Name': 'DEV-1',
            'Vulnerability Name': 'CVE-2026-0001',
            'Affected Products': 'Google Chrome 120.0 x64',
            'Vulnerability Score': 9.5, 'Threat Status': 'UNRESOLVED',
        }])
        _, full, _, raw_total, _ = dp.process_patch_match(str(pfile), cve, min_score=9.0)
        # Only the Chrome patch should have survived to match.
        assert 'installed' in str(full.iloc[0]['Patch Match Result']).lower()