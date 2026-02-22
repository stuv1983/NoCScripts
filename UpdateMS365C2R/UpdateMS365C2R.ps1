<#
.SYNOPSIS
  Forces Microsoft 365 Apps (Office Click-to-Run) to check/download/install updates and reports build/channel before/after.

.DESCRIPTION
  - Detects whether Office is Click-to-Run (C2R) via registry + OfficeC2RClient.exe
  - Captures VersionToReport + UpdateChannel before update
  - Runs OfficeC2RClient.exe /update user (works under SYSTEM or user context in most RMMs)
  - Waits and polls for version change (best-effort)
  - Outputs a simple summary suitable for RMM logs

.NOTES
  - Office updates often complete only after apps are closed; version may not change until Office apps are restarted.
  - Some tenants enforce update policies via GPO/Intune; this script triggers update but cannot override policy blocks.
#>

[CmdletBinding()]
param(
  [int]$MaxWaitMinutes = 30,
  [int]$PollSeconds = 20,
  [switch]$RebootIfRequired
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-OfficeC2RInfo {
  $reg = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  if (-not (Test-Path $reg)) { return $null }

  $p = Get-ItemProperty $reg -ErrorAction Stop
  [pscustomobject]@{
    VersionToReport = $p.VersionToReport
    UpdateChannel   = $p.UpdateChannel
    CDNBaseUrl      = $p.CDNBaseUrl
    Platform        = $p.Platform
    ClientCulture   = $p.ClientCulture
  }
}

function Get-OfficeC2RClientPath {
  $paths = @(
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
  )
  foreach ($p in $paths) {
    if (Test-Path $p) { return $p }
  }
  return $null
}

function Test-PendingReboot {
  # Common reboot-pending indicators (best-effort)
  $keys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
  )
  foreach ($k in $keys) {
    if (Test-Path $k) { return $true }
  }

  # PendingFileRenameOperations
  $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
  if ($pfro.PendingFileRenameOperations) { return $true }

  return $false
}

Write-Output "=== Microsoft 365 Apps (Click-to-Run) Update ==="

$infoBefore = Get-OfficeC2RInfo
if (-not $infoBefore) {
  Write-Output "Office Click-to-Run registry not found. This device may not be using Microsoft 365 Apps (C2R)."
  exit 1
}

$client = Get-OfficeC2RClientPath
if (-not $client) {
  Write-Output "OfficeC2RClient.exe not found. Cannot trigger C2R update."
  exit 1
}

Write-Output "Before:"
Write-Output ("  VersionToReport : {0}" -f $infoBefore.VersionToReport)
Write-Output ("  UpdateChannel   : {0}" -f $infoBefore.UpdateChannel)
Write-Output ("  Platform        : {0}" -f $infoBefore.Platform)
Write-Output ("  CDNBaseUrl      : {0}" -f $infoBefore.CDNBaseUrl)

# Optional: close common Office apps to help updates apply (comment out if you don't want this)
# NOTE: Be careful in production - this will close user apps.
#$officeProcs = "WINWORD","EXCEL","POWERPNT","OUTLOOK","ONENOTE","MSACCESS","VISIO","WINPROJ"
#Get-Process -Name $officeProcs -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Output ""
Write-Output "Triggering update: OfficeC2RClient.exe /update user"
$proc = Start-Process -FilePath $client -ArgumentList "/update user" -PassThru -WindowStyle Hidden
try {
  # Wait for the updater process to exit (doesn't guarantee install complete, but is a good first signal)
  $proc.WaitForExit()
} catch {
  Write-Output "Warning: WaitForExit failed: $($_.Exception.Message)"
}

# Poll for version change (best-effort)
$deadline = (Get-Date).AddMinutes($MaxWaitMinutes)
$changed = $false

do {
  Start-Sleep -Seconds $PollSeconds
  $infoNow = Get-OfficeC2RInfo
  if ($null -eq $infoNow) { break }

  if ($infoNow.VersionToReport -and ($infoNow.VersionToReport -ne $infoBefore.VersionToReport)) {
    $changed = $true
    break
  }
} while ((Get-Date) -lt $deadline)

$infoAfter = Get-OfficeC2RInfo

Write-Output ""
Write-Output "After:"
Write-Output ("  VersionToReport : {0}" -f $infoAfter.VersionToReport)
Write-Output ("  UpdateChannel   : {0}" -f $infoAfter.UpdateChannel)

if ($changed) {
  Write-Output "Result: Office C2R version changed (update likely applied)."
} else {
  Write-Output "Result: Office C2R version did not change within wait window."
  Write-Output "Note: Office updates may require apps to be closed and reopened, and sometimes a reboot, to reflect the new build."
}

$pending = Test-PendingReboot
Write-Output ("Pending reboot detected: {0}" -f $pending)

if ($RebootIfRequired -and $pending) {
  Write-Output "RebootIfRequired specified and reboot is pending. Rebooting now..."
  shutdown /r /t 0
}

Write-Output "=== Done ==="
