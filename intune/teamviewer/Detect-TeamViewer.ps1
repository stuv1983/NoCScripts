<#
.SYNOPSIS
    Hardened Detection for TeamViewer Eradication via Intune.
.DESCRIPTION
    Scans HKLM registry, background services, and per-user AppData for TeamViewer footprints.

    INTUNE DETECTION LOGIC (UNINSTALL ASSIGNMENT):
    - Exit 0 (Success): TeamViewer IS detected. Intune sees the app as present
      and triggers the uninstall assignment to execute the removal script.
    - Exit 1 (Fail):    TeamViewer is NOT detected. Intune considers the device
      clean and the removal script does not run.

    OUTPUT:
    Intune captures stdout from this script. Full output is visible in the
    Intune Management Extension log at:
    C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log
#>

# --- CONFIGURATION TOGGLES ---
$AllowOnServers = $false    # Set to $true to allow detection on Servers and Domain Controllers

# --- STEP 1: WORKSTATION GUARDRAIL ---
# Intune primarily manages workstations. We skip Servers (ProductType 3) and
# Domain Controllers (ProductType 2) by default to prevent accidental removal
# from critical infrastructure.
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

if (-not $os) {
    Write-Output "[Warning] Could not determine OS type via WMI/CIM. Proceeding with detection anyway."
} elseif ($os.ProductType -ne 1) {
    if (-not $AllowOnServers) {
        Write-Output "COMPLIANT: Target is not a workstation (ProductType=$($os.ProductType)). Skipping detection."
        exit 1
    } else {
        Write-Output "[Warning] Target is a server (ProductType=$($os.ProductType)), but `$AllowOnServers is enabled. Proceeding with detection..."
    }
}

$targetFound = $false

# --- STEP 2: REGISTRY DETECTION (HKLM) ---
# Anchored regex '^TeamViewer\b' strictly matches "TeamViewer" and "TeamViewer Host"
# while avoiding false positives from "TeamViewerMeeting" or vendor plugins.
$pattern = "^TeamViewer\b"
$registryPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$installedApps = Get-ItemProperty -Path $registryPaths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match $pattern }

if ($installedApps) {
    Write-Output "NON-COMPLIANT: Found TeamViewer registry entry: $($installedApps.DisplayName)"
    $targetFound = $true
}

# --- STEP 3: SERVICE DETECTION ---
# Catches "ghost" installations where a botched uninstall removed the binaries
# and registry keys but left the background service running.
if (Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue) {
    Write-Output "NON-COMPLIANT: TeamViewer background service detected."
    $targetFound = $true
}

# --- STEP 4: PER-USER APPDATA DETECTION ---
# TeamViewer's consumer installer requires no admin rights and bypasses Program Files
# and HKLM entirely. We scan all real user profiles (identified by NTUSER.DAT) to
# catch these stealth deployments.
#
# Checked paths cover:
#   - AppData\Local\TeamViewer          (standard consumer install)
#   - AppData\Roaming\TeamViewer        (roaming/legacy variant)
#   - AppData\Local\Programs\TeamViewer (TV15+ per-user Programs variant)
$userProfiles = Get-ChildItem -LiteralPath "C:\Users" -Directory -ErrorAction SilentlyContinue |
                Where-Object { Test-Path (Join-Path $_.FullName "NTUSER.DAT") }

foreach ($u in $userProfiles) {
    $checkPaths = @(
        (Join-Path $u.FullName "AppData\Local\TeamViewer\TeamViewer.exe"),
        (Join-Path $u.FullName "AppData\Roaming\TeamViewer\TeamViewer.exe"),
        (Join-Path $u.FullName "AppData\Local\Programs\TeamViewer\TeamViewer.exe")
    )
    foreach ($cp in $checkPaths) {
        if (Test-Path -LiteralPath $cp) {
            Write-Output "NON-COMPLIANT: Per-user binary found at: $cp"
            $targetFound = $true
            break  # Stop checking paths for this user; continue scanning other profiles
        }
    }
}

# --- STEP 5: INTUNE EXIT LOGIC ---
if ($targetFound) {
    # Exit 0 tells Intune: "App is present — run the uninstaller."
    exit 0
} else {
    # Exit 1 tells Intune: "Nothing found — machine is clean."
    Write-Output "COMPLIANT: No TeamViewer footprint detected on this system."
    exit 1
}
