<#
.SYNOPSIS
    Office C2R Active Remediation & Validation Engine
.NOTES
    Name:       OfficeVersionandUpdateChannelAudit.ps1
    Author:     Stu Villanti (s.villanti@kenstra.com.au)
    Version:    5.0
.DESCRIPTION
    Silently triggers Office C2R updates in the background (zero user disruption) 
    and validates the post-trigger state (service health, ADMX blockers, pending reboots).
.EXITCODES
    0 = Success / Compliant / Staged / Not Applicable
    1 = Script Failure (Missing files, errors)
    2 = Reboot Pending (Servicing state is dirty)
#>

[CmdletBinding()]
param ()

# ==============================================================================
# 1. Initialization & Guardrails
# ==============================================================================
$configPath  = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$updatesPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates"
$policyPath  = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate"

if (-not (Test-Path $configPath)) {
    Write-Output "STATUS=NOT_APPLICABLE"
    Write-Output "Message: Office Click-to-Run configuration not found on this device."
    exit 0
}

$initialConfig  = Get-ItemProperty -Path $configPath -ErrorAction SilentlyContinue
$initialVersion = if ($initialConfig.VersionToReport) { $initialConfig.VersionToReport } else { "Unknown" }

$c2rPaths = @(
    "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
)
$updaterExe = $c2rPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $updaterExe) {
    Write-Output "STATUS=FAILED"
    Write-Output "Message: OfficeC2RClient.exe not found in standard directories."
    exit 1
}

# ==============================================================================
# 2. Trigger Background Update
# ==============================================================================
Write-Output "Message: Initial Build: $initialVersion. Triggering silent background update..."
$updateArgs = "/update user updatepromptuser=False forceappshutdown=False displaylevel=False"

try {
    Start-Process -FilePath $updaterExe -ArgumentList $updateArgs -NoNewWindow -Wait
} catch {
    Write-Output "STATUS=FAILED"
    Write-Output "Message: Start-Process failed to execute the updater. Exception: $_"
    exit 1
}

# ==============================================================================
# 3. Bounded Wait Loop
# ==============================================================================
$maxWaitMinutes = 3
$sleepIntervalSeconds = 15
$maxIterations = ($maxWaitMinutes * 60) / $sleepIntervalSeconds
$iteration = 0

$updateTriggered = $false
$updateData = $null

Write-Output "Message: Polling registry for update targeting (Timeout: $maxWaitMinutes min)..."

while ($iteration -lt $maxIterations) {
    Start-Sleep -Seconds $sleepIntervalSeconds
    $iteration++
    
    try {
        $updateData = Get-ItemProperty -Path $updatesPath -ErrorAction Stop
    } catch {
        $updateData = $null
    }
    
    $hasNewTarget = ($updateData -and $null -ne $updateData.UpdateToVersion -and $updateData.UpdateToVersion -ne $initialVersion)
    $hasValidPayload = ($updateData -and $null -ne $updateData.UpdatesReadyToApply -and $updateData.UpdatesReadyToApply -notin @("", "0"))

    if ($hasNewTarget -or $hasValidPayload) {
        $updateTriggered = $true
        break
    }
}

# ==============================================================================
# 4. Post-Update Validation Checks
# ==============================================================================
$finalConfig  = Get-ItemProperty -Path $configPath -ErrorAction SilentlyContinue
$finalVersion = if ($finalConfig.VersionToReport) { $finalConfig.VersionToReport } else { "Unknown" }
$versionChanged = ($finalVersion -and $initialVersion -and $finalVersion -ne "Unknown" -and $finalVersion -ne $initialVersion)

$svc = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue
$svcStatus = if ($svc) { $svc.Status.ToString() } else { "NotFound" }

$policyBlocked = "No"
if (Test-Path $policyPath) {
    $policies = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
    $activePolicies = @()
    if ($null -ne $policies.EnableAutomaticUpdates) { $activePolicies += "EnableAutoUpdates=$($policies.EnableAutomaticUpdates)" }
    if ($null -ne $policies.HideEnableDisableUpdates) { $activePolicies += "HideUI=$($policies.HideEnableDisableUpdates)" }
    if ($null -ne $policies.UpdateBranch) { $activePolicies += "Branch=$($policies.UpdateBranch)" }
    
    $policyBlocked = if ($activePolicies.Count -gt 0) { "Yes ($($activePolicies -join ', '))" } else { "Key exists, no blocking values." }
}

# Granular Reboot Checks
$cbsReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
$wuReboot  = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
$sysReboot = $false

try {
    $pendVals = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue).PendingFileRenameOperations
    if ($null -ne $pendVals -and $pendVals.Count -gt 0) { $sysReboot = $true }
} catch {}

$rebootReasons = @()
if ($cbsReboot) { $rebootReasons += "CBS" }
if ($wuReboot) { $rebootReasons += "WindowsUpdate" }
if ($sysReboot) { $rebootReasons += "PendingFileRename" }

$rebootPending = ($rebootReasons.Count -gt 0)
$rebootDisplay = if ($rebootPending) { "True ($($rebootReasons -join ', '))" } else { "False" }

# ==============================================================================
# 5. Final Output & Exit Codes
# ==============================================================================
Write-Output "--- Validation Details ---"
Write-Output "Service State  : $svcStatus"
Write-Output "Policy Blocker : $policyBlocked"
Write-Output "Reboot Pending : $rebootDisplay"
Write-Output "--------------------------"

if ($versionChanged) {
    Write-Output "STATUS=UPDATED"
    Write-Output "Message: Office successfully updated from $initialVersion to $finalVersion."
    if ($rebootPending) { exit 2 } else { exit 0 }
}

if ($updateTriggered) {
    if ($updateData -and $null -ne $updateData.UpdateToVersion -and $updateData.UpdateToVersion -ne "") {
        $target = $updateData.UpdateToVersion
    } elseif ($updateData -and $null -ne $updateData.UpdatesReadyToApply -and $updateData.UpdatesReadyToApply -ne "") {
        $target = $updateData.UpdatesReadyToApply
    } else {
        $target = "Unknown"
    }
    
    Write-Output "STATUS=ACTION_REQUIRED_PENDING_APPLY"
    Write-Output "Message: Update to $target is staged but NOT YET APPLIED. Device is still running $finalVersion. Office apps must be closed and/or system rebooted to finalize patching."
    if ($rebootPending) { exit 2 } else { exit 0 }
}

Write-Output "STATUS=NO_CHANGE"
Write-Output "Message: No update signal observed within ${maxWaitMinutes}m. Device is either already fully patched or background download requires more time. Current build: $finalVersion."
if ($rebootPending) { exit 2 } else { exit 0 }