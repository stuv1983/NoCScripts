# ==============================================================================
# 1. Initialization & Guardrails
# ==============================================================================
$configPath  = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$updatesPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates"
$policyPath  = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate"

if (-not (Test-Path $configPath)) {
    Write-Output "STATUS=NOT_APPLICABLE"
    Write-Output "Message: Office Click-to-Run configuration not found."
    exit 0
}

$initialConfig = Get-ItemProperty -Path $configPath -ErrorAction SilentlyContinue
$initialVersion = $initialConfig.VersionToReport

$c2rPaths = @(
    "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
    "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
)
$updaterExe = $c2rPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $updaterExe) {
    Write-Output "STATUS=FAILED"
    Write-Output "Message: OfficeC2RClient.exe not found."
    exit 1
}

# ==============================================================================
# 2. Trigger Background Update
# ==============================================================================
Write-Output "Message: Initial Build: $initialVersion. Triggering silent update..."
$updateArgs = "/update user updatepromptuser=False forceappshutdown=False displaylevel=False"

try {
    Start-Process -FilePath $updaterExe -ArgumentList $updateArgs -NoNewWindow -Wait
} catch {
    Write-Output "STATUS=FAILED"
    Write-Output "Message: Failed to execute updater. Exception: $_"
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
# Re-check actual version
$finalConfig  = Get-ItemProperty -Path $configPath -ErrorAction SilentlyContinue
$finalVersion = if ($finalConfig) { $finalConfig.VersionToReport } else { "Unknown" }
$versionChanged = ($finalVersion -and $initialVersion -and $finalVersion -ne $initialVersion)

# Check Service
$svc = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue
$svcStatus = if ($svc) { $svc.Status.ToString() } else { "NotFound" }

# Check Policies (Deep Inspection)
$policyBlocked = "No"
if (Test-Path $policyPath) {
    $policies = Get-ItemProperty -Path $policyPath -ErrorAction SilentlyContinue
    $activePolicies = @()
    if ($null -ne $policies.EnableAutomaticUpdates) { $activePolicies += "EnableAutoUpdates=$($policies.EnableAutomaticUpdates)" }
    if ($null -ne $policies.HideEnableDisableUpdates) { $activePolicies += "HideUI=$($policies.HideEnableDisableUpdates)" }
    if ($null -ne $policies.UpdateBranch) { $activePolicies += "Branch=$($policies.UpdateBranch)" }
    
    if ($activePolicies.Count -gt 0) {
        $policyBlocked = "Yes ($($activePolicies -join ', '))"
    } else {
        $policyBlocked = "Key exists, but no blocking values found."
    }
}

# Check Pending Reboots
$cbsReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
$wuReboot  = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
$sysReboot = $false
try {
    if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue) { $sysReboot = $true }
} catch {}
$rebootPending = ($cbsReboot -or $wuReboot -or $sysReboot)

# ==============================================================================
# 5. Final Output & Exit Codes
# ==============================================================================
Write-Output "--- Validation Details ---"
Write-Output "Service State  : $svcStatus"
Write-Output "Policy Blocker : $policyBlocked"
Write-Output "Reboot Pending : $rebootPending"
Write-Output "--------------------------"

if ($versionChanged) {
    Write-Output "STATUS=UPDATED"
    Write-Output "Message: Office immediately updated from $initialVersion to $finalVersion."
    if ($rebootPending) { exit 2 } else { exit 0 }
}

if ($updateTriggered) {
    $target = if ($updateData -and $updateData.UpdateToVersion) { $updateData.UpdateToVersion } else { "Unknown" }
    Write-Output "STATUS=UPDATE_TRIGGERED"
    Write-Output "Message: Update staged/indicated (target: $target). Current build: $finalVersion"
    if ($rebootPending) { exit 2 } else { exit 0 }
}

Write-Output "STATUS=NO_CHANGE"
Write-Output "Message: No update signal observed within ${maxWaitMinutes}m. Current build: $finalVersion (initial: $initialVersion)."
if ($rebootPending) { exit 2 } else { exit 0 }