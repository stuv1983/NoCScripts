# NoC / MSP Script Collection

> ⚠️ **Disclaimer**  
> Run at your own risk. Test thoroughly in a lab environment first. The author is not responsible for any damage, data loss, or unexpected outcomes.

## Overview

A collection of PowerShell and Python scripts for NOC / MSP environments to automate endpoint health checks, patching, vulnerability auditing, browser governance, software lifecycle management, and operational reporting.

The repository is designed around MSP/RMM use cases:

- Clear console output suitable for RMM capture
- Exit codes that can drive monitoring, alerting, Intune detection, or remediation workflows
- Conservative behaviour where user interruption should be avoided
- Audit-first patterns before enforcement
- Reporting tools for vulnerability, patch, and inventory reconciliation

---

## Repository Structure

```text
NoCScripts/
├── N-able Tools/
│   ├── CVEChecks/
│   │   └── CVE-2026-20804.ps1
│   ├── CVE_Risk_Exposure_&_Remediation/
│   │   ├── config.json
│   │   ├── config.py
│   │   ├── data_pipeline.py
│   │   ├── diagnostics.py
│   │   ├── excel_builder.py
│   │   ├── main.py
│   │   ├── orchestrator.py
│   │   ├── requirements.txt
│   │   ├── requirements-dev.txt
│   │   ├── run_dashboard.py
│   │   ├── snapshot.py
│   │   └── version_sync.py
│   ├── RMMcheck/
│   │   └── RMMcheck.py
│   ├── N-able_CVE_Dashboard.py
│   ├── N-able_PatchReport_Dashboard.py
│   └── cve_risk_exposure_remediation_dashboard.py
│
├── RMM/
│   ├── AVCheck/
│   │   ├── AVCheck.ps1
│   │   └── ACCheck(beta).ps1
│   ├── AutoWindows/
│   │   └── AutoWindowsUpdate.ps1
│   ├── Browser-Audit/
│   │   └── Browser-Audit.ps1
│   ├── CheckTVInstall/
│   │   └── checkTVInstall.ps1
│   ├── DiskCleanUp/
│   │   └── DiskCleanup.ps1
│   ├── HDDCheck/
│   │   └── HDDUsageCheck.ps1
│   ├── Office Version and Update Channel Audit/
│   │   └── OfficeVersionandUpdateChannelAudit.ps1
│   ├── PaperCut Follow Me/
│   │   └── PaperCutFollowMe.ps1
│   ├── PrinterCheck/
│   │   └── printerCheck.ps1
│   ├── SecurityCheck/
│   │   └── SecurityCheck.ps1
│   ├── UpdateFirefoxAndOffice/
│   │   ├── adobeCheckReport.ps1
│   │   ├── chromeCheckUpdate.ps1
│   │   ├── firefoxCheckUpdate.ps1
│   │   ├── main.ps1
│   │   ├── officeCheckUpdate.ps1
│   │   └── UpdateFirefox.ps1
│   └── UpdateMS365C2R/
│       └── UpdateMS365C2R.ps1
│
├── handyPSScripts/
│   └── CalanderPermissions.ps1
│
└── intune/
    ├── chrome/
    │   ├── Audit-Chrome.ps1
    │   ├── Detect-Chrome-Audit.ps1
    │   ├── Detect-Chrome.ps1
    │   ├── Install-Chrome.ps1
    │   └── Uninstall-Audit.ps1
    ├── firefox/
    │   ├── Audit-Firefox.ps1
    │   ├── Detect-Firefox-Audit.ps1
    │   ├── Detect-Firefox.ps1
    │   └── Install-Firefox.ps1
    ├── office/
    │   ├── Detect-OfficeUpdate.ps1
    │   └── Force-OfficeUpdate.ps1
    ├── teamviewer/
    │   ├── Detect-TeamViewer.ps1
    │   └── Remove-TeamViewer.ps1
    └── vlc/
        ├── Detect-VLC.ps1
        └── Install-VLC.ps1
```

---

# Quick Start

## PowerShell Scripts

Most scripts can be run directly in PowerShell or through an RMM remote/background task.

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\ScriptName.ps1
```

For RMM usage, prefer running as **SYSTEM** where the script is designed for device-level checks, and test output/exit codes before deploying broadly.

## Python Tools

Create a virtual environment where possible:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

For the modular CVE dashboard:

```bash
cd "N-able Tools/CVE_Risk_Exposure_&_Remediation"
pip install -r requirements.txt
python main.py
```

---

# N-able Tools

## CVE-2026-20804.ps1

### Overview

Audits Windows devices for protection against **CVE-2026-20804** using a UBR-first model.

The script uses registry-based OS/build detection and treats Update Build Revision (UBR) as the authoritative compliance indicator rather than relying on KB presence.

### Design Goals

- UBR-first compliance model
- No WMI/CIM dependency for OS build detection
- Optional Windows 10 ESU Year 1 enforcement
- Reboot state reported as informational
- RMM/MSP-friendly output and exit codes

### Configuration

| Variable | Default | Description |
|---|---:|---|
| `$FailIfWin10MissingESU` | `$false` | Set to `$true` to fail Windows 10 devices if ESU Year 1 is not active |

### January 2026 Baseline Gates

| Build | OS | Minimum UBR |
|---:|---|---:|
| 26100 | Windows 11 24H2 | 7623 |
| 26200 | Windows 11 25H2 | 7623 |
| 22631 | Windows 11 23H2 | 6491 |
| 22621 | Windows 11 22H2 | 6491 |
| 19045 | Windows 10 22H2 | 6809 |
| 20348 | Windows Server 2022 | 4648 |

### Exit Codes

| Code | Meaning |
|---:|---|
| `0` | Protected / not in scope / informational script condition |
| `1` | Vulnerable, below baseline, or ESU missing when policy is enforced |

---

## CVE Risk Exposure & Remediation Dashboard

Path:

```text
N-able Tools/CVE_Risk_Exposure_&_Remediation/
```

### Overview

A modular Python CVE analysis pipeline for N-able/MSP vulnerability reporting.

It correlates vulnerability exports, RMM device inventory, patch reports, and previous dashboard outputs to produce an actionable Excel workbook.

This tool is intended to answer:

> Are devices actually remediated, or is the patch/vulnerability tooling giving conflicting evidence?

### Production Dependencies

Install from the included requirements file:

```bash
pip install -r requirements.txt
```

Production dependencies include:

- `pandas`
- `xlsxwriter`
- `openpyxl`
- `tkcalendar`

### Architecture

| File | Purpose |
|---|---|
| `main.py` | Tkinter GUI entrypoint |
| `run_dashboard.py` | CLI/headless entrypoint |
| `orchestrator.py` | Coordinates the pipeline |
| `data_pipeline.py` | Loads, normalises, merges, filters, and matches report data |
| `diagnostics.py` | Root-cause and patch evidence classification |
| `excel_builder.py` | Excel workbook/sheet rendering |
| `snapshot.py` | Local JSON snapshot history for trend tracking |
| `version_sync.py` | Syncs product baselines from vendor APIs |
| `config.json` | Product mapping, fixed-version rules, remediation rules |

### Main Capabilities

- Vulnerability/RMM inventory merge
- Optional patch report matching
- Optional patch failure report integration
- Month-over-month trend comparison
- Stale device exclusion
- CVSS threshold filtering
- Product-level triage tabs
- Patch diagnostics and evidence notes
- Health score and recommended actions
- Snapshot storage for historical tracking
- CLI mode for scheduled or headless execution

### Patch Evidence Classification

The dashboard separates findings into operationally useful categories:

| Category | Meaning |
|---|---|
| Patch required | Installed version is below the fixed baseline |
| Patched but still detected | Patch evidence exists, but the scanner still reports the CVE |
| Device missing from patch report | Device is not present in the patch report scope |
| Product not tracked | Device exists in the patch report, but the affected product is not tracked |
| Installed but version unknown | Patch status exists, but version evidence is insufficient |
| No patch baseline defined | The tool lacks a fixed baseline for that product/CVE |

### CLI Usage

Minimal run, skipping RMM merge:

```bash
python run_dashboard.py ^
  --input reports/april_cve.xlsx ^
  --output output/April_Dashboard.xlsx ^
  --skip-rmm
```

With RMM inventory and patch matching:

```bash
python run_dashboard.py ^
  --input reports/april_cve.xlsx ^
  --rmm reports/device_inventory.xlsx ^
  --patch reports/patch_report.csv ^
  --output output/April_Dashboard.xlsx ^
  --threshold 9.0 ^
  --since 2026-04-01
```

With previous dashboard comparison:

```bash
python run_dashboard.py ^
  --input reports/april_cve.xlsx ^
  --rmm reports/device_inventory.xlsx ^
  --output output/April_Dashboard.xlsx ^
  --previous output/March_Dashboard.xlsx
```

### Output Sheets

| Sheet | Purpose |
|---|---|
| Trend Summary | Month-over-month high-level movement |
| New This Month | New CVE types compared to previous report |
| Resolved | CVEs no longer detected |
| Persisting CVEs | CVEs still present from the previous report |
| Monthly Detections | Executive overview and risk summary |
| All Detections | Full filtered detection set |
| Product tabs | Per-product triage sheets |
| Patch Match Overview | Patch evidence roll-up |
| Patch Match Full Data | Detailed patch matching output |
| Patch Report (Full) | Raw patch report evidence |
| Diagnostics | Patch evidence/root-cause diagnostics |
| Stale Excluded Devices | Devices excluded by last-response cutoff |
| Raw Data | Unmodified merged dataset |

---

## N-able_CVE_Dashboard.py

### Overview

Legacy/single-file Python GUI utility that merges N-able vulnerability reports with RMM Device Inventory exports to generate a formatted Excel Executive Risk Dashboard.

This remains useful for simpler reporting workflows, but the modular dashboard folder is the preferred path for expanded patch matching and diagnostics.

### Requirements

```bash
pip install pandas xlsxwriter tkcalendar openpyxl
```

### Features

- Executive risk metrics
- KEV and known-exploit summaries
- Server impact percentage
- Per-product triage tabs
- Stale device exclusion
- Hyperlinked CVE.org and NVD references
- Checkbox-style resolution tracking
- Background processing to keep the GUI responsive

### Inputs

| Input | Purpose |
|---|---|
| Vulnerability Report | N-able vulnerability export, CSV or XLSX |
| Device Inventory / RMM Report | N-able device list, CSV or XLSX |
| Score Threshold | Minimum CVSS score for triage tabs |
| RMM Check-in Cutoff Date | Optional stale device filter |
| Output Path | Destination Excel workbook |

---

## N-able_PatchReport_Dashboard.py

### Overview

Python-based dashboard tooling for N-able patch report analysis.

Use this when the focus is patch-report visibility and actionability rather than CVE-to-patch correlation.

### Typical Use Cases

- Summarise N-able patch status exports
- Identify failed, missing, pending, or installed patch states
- Prepare patch reporting for internal review
- Support monthly patch evidence workflows

---

## RMMcheck.py

### Overview

Python GUI tool that compares an Intune/Entra device export against an N-able RMM export to identify corporate Windows devices that appear to be missing from RMM.

### Requirements

```bash
pip install pandas
```

### Device Classification

A device is treated as corporate if:

- `joinType` contains `JOINED` and not `REGISTERED`, or
- `Ownership` is `CORPORATE`

### Matching Logic

- Builds RMM lookup sets from serial number and normalised device name
- Serial number matches take precedence
- Devices without serials fall back to device-name matching
- Device-name normalisation strips common prefixes and non-alphanumeric characters

### Output

| Result | Output |
|---|---|
| Missing devices found | CSV with device ID, name, user, OS, check-in, serial, ownership, and join type |
| No missing devices | `_nodata.txt` confirmation file |

---

# RMM Scripts

## SecurityCheck.ps1

Path:

```text
RMM/SecurityCheck/SecurityCheck.ps1
```

### Overview

Full endpoint security posture audit in a single pass.

### Checks

| Check | Description |
|---|---|
| Antivirus | Detects Defender/third-party AV and checks service/signature state |
| Defender for Endpoint | Verifies onboarding via Sense service |
| Firewall | Checks Windows Firewall state across profiles |
| Windows Update | Reports last installed update and age |
| Pending Reboot | Detects pending reboot indicators |
| Uptime | Reports current uptime |
| VBS / HVCI | Checks virtualisation-based security and memory integrity |
| LAPS | Checks Local Administrator Password Solution status |
| PowerShell Logging | Reports script block logging state |

### Common Parameters

| Parameter | Purpose |
|---|---|
| `-RequireFirewallOn` | Fail if firewall is off |
| `-RequireRealTime` | Fail if no real-time AV is enabled |
| `-RequireMDE` | Fail if Defender for Endpoint is not onboarded |
| `-RequireVBS` | Fail if VBS/HVCI is not enabled |
| `-RequireLAPS` | Fail if LAPS is not configured |
| `-RequireScriptLogging` | Fail if PowerShell script logging is not enabled |
| `-Strict` | Elevate warnings to failures |
| `-Full` | Detailed output for ticket notes |
| `-AsJson` | Structured JSON output |

### Exit Codes

| Code | Meaning |
|---:|---|
| `0` | Secure/OK |
| `1` | Warning |
| `2` | Critical |
| `4` | Script error |

---

## AVCheck.ps1

Path:

```text
RMM/AVCheck/AVCheck.ps1
```

### Overview

Dedicated antivirus posture check. Useful for high-frequency AV monitoring or replacing noisy RMM AV checks.

### Checks

| Area | Description |
|---|---|
| Installed AV | Enumerates AV products via Windows Security Center |
| Real-time AV | Confirms active real-time protection |
| Defender Health | Checks Defender service, engine, platform, and protection state |
| Signature Currency | Flags stale Defender signatures |
| Defender for Endpoint | Checks Sense/MDE state |
| AV Conflicts | Detects multiple real-time engines |

### Common Parameters

| Parameter | Purpose |
|---|---|
| `-RequireRealTime` | Fail if no real-time AV is enabled |
| `-RequireMDE` | Fail if MDE is not onboarded |
| `-SigFreshHours` | Maximum Defender signature age |
| `-Full` | Detailed output |
| `-AsJson` | JSON output |
| `-DebugMode` | Extra diagnostics |

---

## Browser-Audit.ps1

Path:

```text
RMM/Browser-Audit/Browser-Audit.ps1
```

### Overview

RMM browser governance audit covering Chrome, Edge, and Firefox footprints.

Use this to detect browser drift, unmanaged installs, architecture sprawl, and broken updater states across endpoints.

### What It Helps Find

- Per-user browser installs under user profiles
- System-level browser installs
- 32-bit vs 64-bit browser footprints
- Chrome update service/task health
- Edge update health
- Firefox maintenance service/task health
- Installed and running browser version context
- Devices requiring migration to managed/system-level installs

### Typical Use

Run via RMM Remote Background to gather evidence before enforcing Intune/RMM browser remediation.

---

## AutoWindowsUpdate.ps1

Path:

```text
RMM/AutoWindows/AutoWindowsUpdate.ps1
```

### Overview

Lightweight Windows Update script for RMM scheduled patch windows.

> This script does not upgrade Windows 10 to Windows 11.

### Parameters

| Parameter | Purpose |
|---|---|
| `-CheckOnly` | List pending updates without installing |
| `-Install` | Install available updates |
| `-Reboot` | Restart automatically if required |

Automatically installs the `PSWindowsUpdate` module if missing.

---

## HDDUsageCheck.ps1

Path:

```text
RMM/HDDCheck/HDDUsageCheck.ps1
```

### Overview

Audits disk usage for MSP environments.

### Reports

| Category | Description |
|---|---|
| Drive capacity | Total, used, and free space |
| User profile size | Recursive usage per `C:\Users` profile |
| Large folders | Profile subfolders over configured threshold |
| Large files | Large files including OST files |

### Common Configuration

| Variable | Default | Purpose |
|---|---:|---|
| `$LowSpaceThreshold` | `10%` | Warn below this free-space percentage |
| `$LargeThresholdGB` | `5` | Report folders/files over this size |
| `$UserRoot` | `C:\Users` | User profile root |

---

## DiskCleanup.ps1

Path:

```text
RMM/DiskCleanUp/DiskCleanup.ps1
```

### Overview

Silent disk cleanup automation for MSP/RMM environments.

Runs `cleanmgr.exe` via a temporary hidden SYSTEM scheduled task and reports before/after disk usage.

### Features

- Hidden execution
- Dry-run/test mode
- Configurable cleanup categories
- Before/after/delta reporting
- Temporary scheduled task cleanup
- Mutex to prevent concurrent execution
- No persistent on-disk logs by default

### Usage

```powershell
# Live run
.\DiskCleanup.ps1

# Dry run
.\DiskCleanup.ps1 -Mode test
```

---

## UpdateMS365C2R.ps1

Path:

```text
RMM/UpdateMS365C2R/UpdateMS365C2R.ps1
```

### Overview

Triggers Microsoft 365 Apps Click-to-Run update processing from RMM.

Useful where Office updates need to be nudged without forcibly closing Office applications.

### Operational Notes

- Uses the native Office Click-to-Run update engine
- Suitable for background remediation
- Avoids user disruption where configured not to force app shutdown
- Complements Microsoft 365 Apps Cloud Update / Intune policy workflows

---

## OfficeVersionandUpdateChannelAudit.ps1

Path:

```text
RMM/Office Version and Update Channel Audit/OfficeVersionandUpdateChannelAudit.ps1
```

### Overview

Audits Microsoft Office Click-to-Run version and update channel.

Useful for identifying devices on the wrong channel, devices lagging behind target builds, or devices not aligned with tenant update policy.

---

## UpdateFirefoxAndOffice Scripts

Path:

```text
RMM/UpdateFirefoxAndOffice/
```

### Scripts

| Script | Purpose |
|---|---|
| `main.ps1` | Combined/update orchestration script |
| `UpdateFirefox.ps1` | Firefox update workflow |
| `chromeCheckUpdate.ps1` | Chrome update check/report workflow |
| `firefoxCheckUpdate.ps1` | Firefox update check/report workflow |
| `officeCheckUpdate.ps1` | Office update check/report workflow |
| `adobeCheckReport.ps1` | Adobe update/check reporting workflow |

### Typical Use

Use these scripts where RMM needs individual browser/application update status or a targeted update action without deploying a full Intune Win32 package.

---

## PrinterCheck.ps1

Path:

```text
RMM/PrinterCheck/printerCheck.ps1
```

### Overview

Printer and network print troubleshooting helper.

Useful for RMM-side checks where printer reachability, configuration, or network path evidence needs to be captured for ticket notes.

---

## PaperCutFollowMe.ps1

Path:

```text
RMM/PaperCut Follow Me/PaperCutFollowMe.ps1
```

### Overview

PaperCut Follow Me print support script.

Use for PaperCut/Follow Me print deployment or support workflows where print queue setup or validation is required.

---

## checkTVInstall.ps1

Path:

```text
RMM/CheckTVInstall/checkTVInstall.ps1
```

### Overview

Checks for TeamViewer installation/footprints from RMM.

Useful as a lightweight audit before or after TeamViewer eradication workflows.

---

# Intune Scripts

These scripts follow an audit-first pattern where possible.

| Phase | Purpose | Assignment Style |
|---|---|---|
| Phase 1 – Audit | Report state only, no changes | Broad/required audit |
| Phase 2 – Enforcement | Detection plus remediation | Targeted rollout |

The detection script decides compliance. Remediation/install scripts run only when detection returns non-compliant.

---

## Intune Chrome

Path:

```text
intune/chrome/
```

### Audit Phase

| Script | Purpose |
|---|---|
| `Detect-Chrome-Audit.ps1` | Reports compliance without remediation |
| `Audit-Chrome.ps1` | No-op install placeholder |
| `Uninstall-Audit.ps1` | No-op uninstall placeholder |

Audit checks include:

- 32-bit Chrome in `Program Files (x86)`
- Per-user AppData Chrome installs
- Missing/disabled Google Update services
- Missing update tasks
- Chrome absent = compliant/no action

### Enforcement Phase

| Script | Purpose |
|---|---|
| `Detect-Chrome.ps1` | Full compliance detection |
| `Install-Chrome.ps1` | Remediation / enterprise MSI enforcement |

Non-compliance examples:

- 32-bit Chrome present
- Per-user unmanaged Chrome found
- Installed version below minimum floor
- Google Update services disabled/missing
- Google Update scheduled tasks missing

Remediation behaviour:

- Removes unmanaged binary footprints while preserving user data where designed
- Installs or repairs 64-bit system-level Chrome
- Rebuilds update engine where required
- Waits for Chrome to close before destructive actions
- Uses retry-friendly exit behaviour for active sessions

---

## Intune Firefox

Path:

```text
intune/firefox/
```

### Audit Phase

| Script | Purpose |
|---|---|
| `Detect-Firefox-Audit.ps1` | Reports compliance without remediation |
| `Audit-Firefox.ps1` | No-op install placeholder |

Audit checks include:

- 32-bit Firefox
- Per-user rogue Firefox installs
- Missing/disabled Mozilla Maintenance Service
- Firefox absent = compliant/no action

### Enforcement Phase

| Script | Purpose |
|---|---|
| `Detect-Firefox.ps1` | Full compliance detection |
| `Install-Firefox.ps1` | Enterprise remediation with profile preservation |

Remediation behaviour:

- Process gate before destructive work
- x86 uninstall where required
- HKCU/user-hive cleanup
- Rogue folder removal
- Shortcut cleanup
- MSI install using system-level deployment
- Enterprise policy configuration
- Profile restore and onboarding suppression

---

## Intune Office

Path:

```text
intune/office/
```

### Scripts

| Script | Purpose |
|---|---|
| `Detect-OfficeUpdate.ps1` | Detects whether Office C2R is at or above a target version |
| `Force-OfficeUpdate.ps1` | Triggers native Office C2R updater |

### Behaviour

- Office missing = compliant/no action
- Office below target = non-compliant
- Update runs silently through `OfficeC2RClient.exe`
- Avoids force-closing Office apps where configured

---

## Intune TeamViewer

Path:

```text
intune/teamviewer/
```

### Overview

TeamViewer detection and eradication workflow for managed endpoints.

### Scripts

| Script | Purpose |
|---|---|
| `Detect-TeamViewer.ps1` | Detects TeamViewer footprints |
| `Remove-TeamViewer.ps1` | Removes TeamViewer across services, registry, tasks, files, and user profiles |

### Notes

- Designed for uninstall/eradication assignments
- Includes guardrails for servers unless overridden
- Scans HKLM uninstall keys, services, and per-user AppData
- Removes scheduled tasks and orphaned services
- Performs post-removal verification

---

## Intune VLC

Path:

```text
intune/vlc/
```

### Overview

Update-only VLC deployment.

Devices without VLC are treated as compliant.

### Scripts

| Script | Purpose |
|---|---|
| `Detect-VLC.ps1` | Detects installed VLC and compares to online latest version |
| `Install-VLC.ps1` | Downloads and installs latest VLC silently |

### Behaviour

- Checks 64-bit then 32-bit VLC
- Does not install VLC where it is absent
- Skips update if VLC is running to avoid user disruption
- Validates download before execution

---

# Handy PowerShell Scripts

## CalanderPermissions.ps1

Path:

```text
handyPSScripts/CalanderPermissions.ps1
```

### Overview

Exchange Online / mailbox calendar permission helper script.

Useful for reviewing or applying calendar folder permissions as part of help desk or Microsoft 365 administration workflows.

> Note: File name currently appears as `CalanderPermissions.ps1` in the repo.

---

# Operational Guidance

## Testing

Before running across production:

1. Test on a lab machine.
2. Confirm expected stdout/stderr output.
3. Confirm exit code behaviour.
4. Confirm whether the script should run as user or SYSTEM.
5. Confirm whether the script is audit-only or remediation-capable.
6. For Intune scripts, validate detection logic independently before assigning remediation broadly.

## RMM Deployment Tips

- Prefer Remote Background execution where no UI is required.
- Avoid destructive remediation on SCADA/OT or other change-controlled endpoints without approval.
- Capture output into ticket notes for auditability.
- Use JSON output where available for ingestion or dashboarding.
- For user-facing apps, prefer non-disruptive update behaviour unless there is an approved maintenance window.

## Git Workflow

```bash
git status
git add .
git commit -m "type: short description"
git pull --rebase origin main
git push origin main
```

If the branch has diverged:

```bash
git fetch origin
git log --oneline --graph --decorate --all -n 10
git pull --rebase origin main
git push origin main
```

If local uncommitted changes block a rebase:

```bash
git stash push -u -m "temp before rebase"
git pull --rebase origin main
git stash pop
git push origin main
```

---

# Author

Developed and maintained by **Stu Villanti** for NOC/MSP automation, vulnerability reporting, patch lifecycle management, and endpoint governance.
