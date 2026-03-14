# NoC / MSP Script Collection

> ⚠️ **Disclaimer**  
> Run at your own risk. Test thoroughly in a lab environment first. The author is not responsible for any damage, data loss, or unexpected outcomes.

## Overview
A collection of **PowerShell and Python scripts** for **NOC / MSP** environments to automate endpoint health checks, patching, vulnerability auditing, and software lifecycle management.  
All scripts are **RMM-friendly** (clear console output + exit codes) and safe for scheduled automation when properly tested.

---

## 📁 Repository Structure

```
├── N-able Tools/
│   ├── CVEChecks/
│   │   └── CVE-2026-20804.ps1
│   ├── RMMcheck/
│   │   └── RMMcheck.py
│   └── N-able_CVE_Dashboard.py
│
├── intune/
│   ├── chrome/
│   │   ├── Audit-Chrome.ps1
│   │   ├── Detect-Chrome-Audit.ps1
│   │   ├── Detect-Chrome.ps1
│   │   ├── Install-Chrome.ps1
│   │   └── Uninstall-Audit.ps1
│   ├── firefox/
│   │   ├── Audit-Firefox.ps1
│   │   ├── Detect-Firefox-Audit.ps1
│   │   ├── Detect-Firefox.ps1
│   │   ├── Install-Firefox.ps1
│   │   └── Uninstall-Audit.ps1
│   ├── office/
│   │   ├── Detect-OfficeUpdate.ps1
│   │   └── Force-OfficeUpdate.ps1
│   ├── teamviewer/
│   │   ├── Detect-TeamViewer.ps1
│   │   └── Remove-TeamViewer.ps1
│   └── vlc/
│       ├── Detect-VLC.ps1
│       └── Install-VLC.ps1
│
├── AVCheck.ps1
├── SecurityCheck.ps1
├── AutoWindowsUpdate.ps1
├── Update-Chrome.ps1
├── HDDUsageCheck.ps1
└── AutoDiskCleanup.ps1
```

---

# N-able Tools

## CVE-2026-20804.ps1

### Overview
Audits Windows devices for protection against **CVE-2026-20804** using a UBR-first (Update Build Revision) model.  
Registry-only OS detection — no WMI/CIM dependency for build information.

### Design Goals
- UBR is the authoritative compliance indicator (KB presence is not used — cumulative updates are frequently superseded and may not appear in QFE listings even when patched)
- Independent, optional enforcement of Windows 10 ESU Year 1
- Reboot detection is informational only and noise-filtered
- RMM/MSP friendly output and exit codes

### Configuration
| Variable | Default | Description |
|---|---|---|
| `$FailIfWin10MissingESU` | `$false` | `$true` = fail Win10 devices if ESU Year 1 is not active |

### January 2026 Baseline (UBR Gates)
| Build | OS | Min UBR |
|---|---|---|
| 26100 | Windows 11 24H2 | 7623 |
| 26200 | Windows 11 25H2 | 7623 |
| 22631 | Windows 11 23H2 | 6491 |
| 22621 | Windows 11 22H2 | 6491 |
| 19045 | Windows 10 22H2 | 6809 |
| 20348 | Windows Server 2022 | 4648 |

### Exit Codes
| Code | Meaning |
|---|---|
| `0` | Protected / Not in scope / Script error (informational) |
| `1` | Vulnerable (below baseline) or ESU missing (if policy enforced) |

---

## N-able_CVE_Dashboard.py

### Overview
A Python GUI utility that merges **N-able Vulnerability reports** with **RMM Device Inventory exports** to generate a fully formatted, actionable **Excel Executive Risk Dashboard**.

### Requirements
```
pip install pandas xlsxwriter tkcalendar openpyxl
```

### Features
- **Executive risk metrics** — KEV CVE counts, exploitability summaries, server impact percentage
- **Per-product triage sheets** — one tab per affected product with sortable CVE/device data
- **Stale device exclusion** — filters devices below a configurable last-seen cutoff, with a dedicated exclusion sheet
- **Hyperlinked CVEs** — links to CVE.org and NVD per finding
- **Resolution tracking** — checkbox (☐/☑) dropdowns per row with live-counting Overview formulas
- **RMM device type detection** — classifies Server vs Workstation from OS field
- **Background threading** — heavy processing runs off the UI thread to prevent freezing

### GUI Inputs
1. **Vulnerability Report** — N-able vulnerability export (CSV or XLSX)
2. **Device Inventory / RMM Report** — N-able device list (CSV or XLSX) — optional (can be skipped)
3. **Score Threshold** — minimum CVSS score to include in per-product triage tabs (default: 9.0)
4. **RMM Check-in Cutoff Date** — excludes stale devices from triage tabs (can be disabled to show all)

### Output Sheets
| Sheet | Contents |
|---|---|
| Overview | Executive metrics, severity breakdown, top 10 products, resolution status, unsynced devices |
| All Detections | Full merged dataset, all scores, sortable/filterable |
| [Product name] | Per-product triage tabs for findings above the score threshold |
| Stale Excluded Devices | Devices filtered out by the cutoff date |
| Raw Data | Unmodified merged dataset for reference |

---

## RMMcheck.py

### Overview
A Python GUI tool that compares an **Intune/Entra device export** against an **N-able RMM export** to identify corporate devices present in Intune but **missing from RMM** — useful for licence auditing and gap detection.

### Requirements
```
pip install pandas
```

### Device Classification
A device is treated as **corporate** if:
- `joinType` contains `JOINED` (and not `REGISTERED`), **or**
- `Ownership == CORPORATE`

### Matching Logic
1. RMM lookup sets are built from **serial number** and **normalised device name**
2. Serial number matches take precedence; devices without serials fall back to name matching
3. Name normalisation strips `LAPTOP-`/`DESKTOP-` prefixes and non-alphanumeric characters

### Filtering Applied to Intune Devices
- Last sign-in must be **2025-01-01 or later**
- `operatingSystem` must contain `WINDOWS`
- De-duplicated by device name (latest record kept)

### Output
- If missing devices found: CSV with device ID, name, primary user, OS, last check-in, serial, ownership, and join type
- If none found: plain text `_nodata.txt` confirmation file

---

# Intune Scripts

All Intune scripts follow a **two-phase deployment pattern**:

| Phase | Purpose | Intune Assignment |
|---|---|---|
| **Phase 1 – Audit** | Zero-footprint compliance reporting only. No changes made. | Required (all users) |
| **Phase 2 – Enforcement** | Full detection + remediation. Applies fixes. | Available / targeted rollout |

The Detect scripts control whether Intune considers the app compliant. The Install/Remediation scripts run only when detection returns non-compliant.

---

## Chrome

### Phase 1 — Audit (Reporting Only)

**`Detect-Chrome-Audit.ps1`** — Detection  
Reports compliance state without triggering any remediation. Checks for:
- 32-bit Chrome in `Program Files (x86)` (non-compliant)
- Unmanaged per-user AppData installs (non-compliant)
- Missing or disabled Google Update services (non-compliant)
- Chrome simply not installed → compliant (no install performed)

**`Audit-Chrome.ps1`** — Install (No-Op)  
Placeholder satisfying Intune's mandatory Install command. Performs no actions.

**`Uninstall-Audit.ps1`** — Uninstall (No-Op)  
Placeholder satisfying Intune's mandatory Uninstall command. Performs no actions.

---

### Phase 2 — Enforcement

**`Detect-Chrome.ps1`** — Detection  
Full compliance check. Non-compliant if any of the following are true:
- 32-bit installation present
- Unmanaged per-user AppData install found
- Installed version below `$MinimumVersion` floor
- Google Update services missing or disabled
- Google Update scheduled tasks missing

**`Install-Chrome.ps1`** — Remediation  
Surgical remediation targeting only what is broken:

| Condition Detected | Action Taken |
|---|---|
| 32-bit architecture | Purge + MSI install |
| AppData shadow install | Purge (binary only, user data preserved) |
| Missing 64-bit system install | MSI install |
| Missing scheduled update tasks | MSI install (rebuilds update engine) |
| Version below floor | Synchronous MSI in-place upgrade |
| Disabled update service | Silent `Set-Service` repair (no MSI needed) |

Additional behaviours:
- **Patient Process Gate** — waits up to 45 minutes for Chrome to close before making destructive changes; exits `1618` on timeout so Intune retries later
- **Desktop sanitisation** — removes per-user Chrome shortcuts and `.exe` stubs from desktops
- **Shortcut rewiring** — updates Start Menu, taskbar pins, and Public Desktop shortcuts to the 64-bit path

Exit Codes: `0` = success, `1618` = deferred (Chrome running), any MSI code on failure.

---

## Firefox

### Phase 1 — Audit (Reporting Only)

**`Detect-Firefox-Audit.ps1`** — Detection  
Reports compliance state without triggering remediation. Checks for:
- 32-bit Firefox in `Program Files (x86)` (non-compliant)
- Unmanaged per-user AppData installs (non-compliant)
- Missing or disabled Mozilla Maintenance Service (non-compliant)
- Firefox not installed → compliant

**`Audit-Firefox.ps1`** — Install (No-Op)  
Placeholder. No actions performed.

**`Uninstall-Audit.ps1`** — Uninstall (No-Op)  
Placeholder. No actions performed.

---

### Phase 2 — Enforcement

**`Detect-Firefox.ps1`** — Detection  
Full compliance check. Non-compliant if:
- 32-bit Firefox directory exists
- Per-user rogue install found in `AppData\Local\Mozilla Firefox` or `AppData\Local\Programs\Mozilla Firefox`
- Installed version below `$MinimumVersion` floor (currently `148.0`)
- Version string cannot be parsed

> **Note:** Constants (`$MinimumVersion`, `$Firefox64Exe`, `$Firefox86Dir`, `$RoguePaths`) are intentionally duplicated between Detect and Install scripts — Intune evaluates them in separate processes and cannot share state.

**`Install-Firefox.ps1`** — Remediation  
Full enterprise deployment with profile preservation. Key stages:

| Stage | Description |
|---|---|
| Process Gate | Waits up to 45 minutes for Firefox to close; exits `1` on timeout |
| x86 Uninstall | Uses registered `UninstallString` (MSI GUID or helper.exe) for clean removal |
| HKCU Cleanup | Loads each user's NTUSER.DAT hive to remove per-user AppData registrations and disable per-user Firefox scheduled tasks |
| Rogue folder removal | Removes `AppData\Local` and `AppData\Local\Programs` Firefox installs |
| Shortcut cleanup | Removes all Firefox shortcuts pre-install (MSI recreates them at the correct path) |
| Profile snapshot | Captures each user's active profile name and path *before* the MSI runs |
| MSI install | `ALLUSERS=1 /qn /norestart` — upgrades in-place |
| Enterprise policies | Writes registry policies (no first-run page, no telemetry, silent auto-updates, disables Default Browser Agent) |
| Profile restore | Restores `profiles.ini`, `installs.ini`, and `compatibility.ini` for each user post-install; writes `user.js` to suppress onboarding on next launch |
| installs.ini re-assertion | Re-writes `installs.ini` as the final step to override any MSI post-install activity |
| Validation | Two-tier: critical failures exit `1`; per-user profile check failures are logged as warnings |

Both known Firefox install path CRC hashes are written to `installs.ini`/`profiles.ini`:
- `308046B0AF4A39CB` — `C:\Program Files\Mozilla Firefox\firefox.exe`
- `E7CF176E110C211B` — `C:\Program Files (x86)\Mozilla Firefox\firefox.exe`

Exit Codes: `0` = success, `1` = failure or timeout, `3010` = success with reboot pending.

---

## Office (Click-to-Run)

**`Detect-OfficeUpdate.ps1`** — Detection  
Reads `VersionToReport` from the Click-to-Run registry configuration key.  
- Office not installed → compliant (exit 0)
- Version at or above `$TargetVersion` → compliant (exit 0)
- Version below `$TargetVersion` → non-compliant (exit 1)

Update `$TargetVersion` to set the minimum acceptable build.

**`Force-OfficeUpdate.ps1`** — Remediation  
Triggers the native `OfficeC2RClient.exe` updater asynchronously with:
- `displaylevel=false` — hides all update UI from the user
- `forceappshutdown=false` — does **not** kill open Office apps; update stages in the background and applies when apps close

> Intune may temporarily report the device as failed while the background update downloads and installs. Detection will return compliant on the next Intune sync cycle after the update completes.

---

## TeamViewer

Designed for **eradication** deployments — removes TeamViewer from managed endpoints via Intune uninstall assignment.

> ⚠️ **Inverted exit code logic:** In an uninstall assignment, exit 0 means "app is present — run the removal script." Exit 1 means "nothing found — machine is clean."

**`Detect-TeamViewer.ps1`** — Detection  
Scans three vectors:
1. **HKLM registry** — Uninstall keys matching `^TeamViewer\b` (avoids `TeamViewerMeeting` false positives)
2. **Services** — Any running service matching `*TeamViewer*`
3. **Per-user AppData** — `AppData\Local\TeamViewer`, `AppData\Roaming\TeamViewer`, and `AppData\Local\Programs\TeamViewer` across all real user profiles (identified by `NTUSER.DAT`)

Configuration toggle `$AllowOnServers` — when `$false` (default), skips detection on Servers and Domain Controllers.

**`Remove-TeamViewer.ps1`** — Removal  
Full eradication across 9 stages:

| Stage | Action |
|---|---|
| 1 | Server guardrail (skips non-workstations unless overridden) |
| 2 | Optional active-session wait loop (up to 45 min); exits `1618` on timeout |
| 3 | Stop all TeamViewer services; kill processes with taskkill fallback |
| 4 | Execute official uninstallers from registry (MSI GUID or EXE silent switch) |
| 5 | Remove TeamViewer scheduled tasks |
| 6 | Delete system and per-user filesystem remnants |
| 7 | Remove HKLM and all loaded user hive registry keys |
| 8 | Delete any orphaned services via `sc.exe delete` |
| 9 | Post-uninstall verification across all vectors; exits `1` if any footprint remains |

Exit Codes: `0` = eradicated, `1` = failure/footprint remains, `1618` = active session (retry), `3010` = success with reboot pending.  
Transcript logged to `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\TeamViewer-Removal.log`.

---

## VLC

Update-only deployment — devices without VLC are considered compliant (no install performed).

**`Detect-VLC.ps1`** — Detection  
1. Checks for 64-bit then 32-bit VLC installation
2. If not installed → compliant (exit 0)
3. Queries VideoLAN's live status endpoint for the latest version
4. If local version ≥ online version → compliant (exit 0); otherwise → non-compliant (exit 1)
5. If the status endpoint is unreachable → compliant (exit 0, skip to avoid noise)

**`Install-VLC.ps1`** — Update  
Downloads and installs the latest VLC silently using VideoLAN's `last/` stable endpoint.
- Skips if VLC is running (avoids disrupting the user; detection remains non-compliant until a later sync)
- Skips if already at or above `$TargetVersion`
- Download via BITS with `Invoke-WebRequest` fallback
- Validates download size before executing installer
- Cleans up the installer on success

Exit Codes: `0` = success or no action required, `1` = failure.

---

# N-able NOC Scripts

## SecurityCheck.ps1

Performs a full endpoint security posture audit in a single pass.

### What It Checks
| Check | Description |
|---|---|
| Antivirus | Detects Defender / third-party AV; checks service state + signatures |
| MDE | Verifies Defender for Endpoint onboarding (Sense service) |
| Firewall | Confirms Windows Firewall state across all profiles |
| Windows Update | Reports date and age of last installed update |
| Pending Reboot | Detects pending reboot state |
| Uptime | Reports current device uptime |
| VBS / HVCI | Validates Virtualization-Based Security and HVCI state |
| LAPS | Checks Local Administrator Password Solution status |
| PS Script Logging | Reports PowerShell script block logging state |

### Parameters
| Parameter | Purpose |
|---|---|
| `-RequireFirewallOn` | Fail if Windows Firewall is off |
| `-RequireRealTime` | Fail if no real-time AV is enabled |
| `-RequireMDE` | Fail if MDE is not onboarded |
| `-RequireVBS` | Fail if VBS/HVCI is not enabled |
| `-RequireLAPS` | Fail if LAPS is not configured |
| `-RequireScriptLogging` | Fail if PS script logging is not enabled |
| `-Strict` | Elevate warnings to failures |
| `-Full` | Detailed multi-line output for ticket notes |
| `-AsJson` | Structured JSON output for ingestion |

### Exit Codes
| Code | Meaning |
|---|---|
| `0` | Secure |
| `1` | Warning |
| `2` | Critical |
| `4` | Script error |

---

## AVCheck.ps1

Dedicated antivirus posture check. Narrower than SecurityCheck.ps1 — ideal for high-frequency AV-only monitoring or replacing unreliable built-in RMM AV checks.

### What It Checks
| Area | Description |
|---|---|
| Installed AV | Enumerates AV products via Windows Security Center (WSC) |
| Active / Real-Time AV | Confirms which AV engine has real-time protection enabled |
| Defender Health | Service state, real-time protection, engine & platform versions |
| Signature Currency | Flags stale Defender signatures (configurable threshold) |
| MDE (Sense) | Detects Defender for Endpoint onboarding via Sense service |
| AV Conflicts | Detects multiple real-time AV engines enabled simultaneously |

### Parameters
| Parameter | Purpose |
|---|---|
| `-RequireRealTime` | Fail if no real-time AV is enabled |
| `-RequireMDE` | Fail if MDE is not onboarded |
| `-SigFreshHours` | Max allowed Defender signature age (default: 48h) |
| `-Full` | Detailed multi-line output |
| `-AsJson` | Structured JSON output |
| `-DebugMode` | Extra diagnostics for troubleshooting |

### Exit Codes
| Code | Meaning |
|---|---|
| `0` | OK / Secure |
| `1` | Warning |
| `2` | Critical |
| `4` | Script error |

---

## AutoWindowsUpdate.ps1

Lightweight script for checking and installing Windows Updates. Designed for RMM scheduled patch windows.

> Does **not** upgrade Windows 10 to Windows 11.

### Parameters
| Parameter | Purpose |
|---|---|
| `-CheckOnly` | List pending updates without installing |
| `-Install` | Install all available updates |
| `-Reboot` | Automatically restart if updates require it |

Automatically installs `PSWindowsUpdate` module if missing.

---

## Update-Chrome.ps1 (N-able / RMM)

Updates Google Chrome via Winget. Handles the common "Pending Relaunch" scenario where Chrome has staged an update but the user hasn't closed the browser.

### Logic
1. Is `chrome.exe` running?
   - **No** → run `winget upgrade Google.Chrome`
   - **Yes** → compare registry staged version against running version
2. If registry version > running version → report "Pending Relaunch", skip update
3. If no reboot pending → check Winget for new version; if available, report and skip (preserves active session)

### Example Output
```
[!] PENDING REBOOT DETECTED
    Running Version:  120.0.6099.109
    Staged Version:   120.0.6099.130
    Action: Skipped. Chrome needs a relaunch to apply the staged update.
```

---

## HDDUsageCheck.ps1

Audits system disk usage for MSP environments — drive-level statistics, per-user profile space, and large folder/file detection.

### What It Reports
| Category | Description |
|---|---|
| Drive capacity | Total, used, and free space per fixed drive |
| Per-user profile size | Recursive space usage per `C:\Users` profile |
| Large folders | Subfolders ≥ 5 GB within each profile |
| Large files | All files ≥ 5 GB including OST files |

### Configuration
| Variable | Default | Description |
|---|---|---|
| `$LowSpaceThreshold` | 10% | Free space % below which to flag a warning |
| `$LargeThresholdGB` | 5 | Size in GB to report as large |
| `$UserRoot` | `C:\Users` | Root path for user profiles |

---

## AutoDiskCleanup.ps1

Production-ready silent disk cleanup automation for MSP/RMM environments. Runs `cleanmgr.exe` via a temporary hidden SYSTEM scheduled task with before/after disk space reporting.

### Features
- Fully hidden execution (no user-visible UI)
- Test / dry-run mode
- Configurable cleanup categories
- Before / after / delta space reporting
- Automatic removal of the temporary scheduled task
- Mutex prevents concurrent execution
- No on-disk logs (stdout only — captured by RMM)

### Usage
```powershell
# Live run
.\DiskCleanup.ps1

# Dry-run
.\DiskCleanup.ps1 -Mode test
```

### Cleanup Categories
Temporary Files, Temporary Setup Files, Recycle Bin, Windows Error Reporting Files, System error memory dump files, System error minidump files, Update Cleanup, Device Driver Packages, Old ChkDsk Files, Setup Log Files, Thumbnail Cache.  
Missing categories on a given OS are silently skipped.

---

## Author
Developed and maintained by **Stu Villanti** for NOC/MSP automation and patch lifecycle management.

## Versioning
```bash
git add .
git commit -m "your message"
git push
```
