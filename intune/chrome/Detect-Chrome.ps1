<#
.SYNOPSIS
    Google Chrome Phase 2 Enforcement Detection
.DESCRIPTION
    Validates Chrome compliance for Microsoft Intune.
    Fails (Exit 1) if it detects x86 architecture, unmanaged AppData installs, 
    outdated versions, or broken/missing servicing engines (Services & Tasks).
#>

$ChromeSystem64 = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$Chromex86 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
$MinimumVersion = [version]"145.0.7632.117" # Update this to your organizational security floor
$AppDataInstallFound = $false

# ------------------------------------------------------------------------
# 1. ARCHITECTURE & SHADOW IT TRAPS
# ------------------------------------------------------------------------
# Trap 1: Legacy 32-bit Architecture
if (Test-Path $Chromex86) {
    Write-Output "Non-Compliant (Exit 1): 32-bit installation detected in x86 folder."
    exit 1
}

# Trap 2: Unmanaged per-user AppData installs
# We strictly exclude default system profiles to avoid regex matching errors
$ExcludedProfiles = @('Public', 'Default', 'Default User', 'All Users')
$UserProfiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin $ExcludedProfiles }

foreach ($Profile in $UserProfiles) {
    if (Test-Path (Join-Path $Profile.FullName "AppData\Local\Google\Chrome\Application\chrome.exe")) { 
        $AppDataInstallFound = $true; break 
    }
}

if ($AppDataInstallFound) {
    Write-Output "Non-Compliant (Exit 1): Rogue AppData installation detected."
    exit 1
}

# ------------------------------------------------------------------------
# 2. ESCAPE HATCH & VERSION FLOOR
# ------------------------------------------------------------------------
# If no bad binaries exist, and no good binaries exist, the device is clean.
if (-not (Test-Path $ChromeSystem64)) {
    Write-Output "Compliant (Exit 0): Chrome is not installed."
    exit 0
}

# Cast the raw file version to a [version] object to ensure accurate mathematical comparison
$CurrentVersion = [version](Get-Item $ChromeSystem64).VersionInfo.ProductVersion
if ($CurrentVersion -lt $MinimumVersion) {
    Write-Output "Non-Compliant (Exit 1): Chrome version $CurrentVersion is below floor $MinimumVersion."
    exit 1
}

# ------------------------------------------------------------------------
# 3. SERVICING ENGINE VALIDATION (Services & Tasks)
# ------------------------------------------------------------------------
# Chrome services naturally sit at "Automatic" but "Stopped". We only fail if Disabled or Missing.
$UpdateServices = @(Get-Service -Name "gupdate", "gupdatem", "GoogleUpdater*" -ErrorAction SilentlyContinue)

if ($UpdateServices.Count -eq 0) {
    Write-Output "Non-Compliant (Exit 1): Update service is missing entirely."
    exit 1
}

if (-not ($UpdateServices | Where-Object { $_.StartType -ne 'Disabled' })) {
    Write-Output "Non-Compliant (Exit 1): Update services are present but Disabled."
    exit 1
}

# CRITICAL FIX: Wrapped in @() to ensure .Count works even if only 1 task is found.
# Matches both Legacy (GoogleUpdateTaskMachine) and Modern (GoogleUpdaterTaskSystem) names.
$Tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { 
    $_.TaskName -match "^GoogleUpdateTaskMachine" -or 
    $_.TaskName -match "^GoogleUpdaterTaskSystem" 
})

if ($Tasks.Count -eq 0) {
    Write-Output "Non-Compliant (Exit 1): Google Update Scheduled Tasks are missing."
    exit 1
}

Write-Output "Compliant (Exit 0): Chrome $CurrentVersion is fully governed and healthy."
exit 0