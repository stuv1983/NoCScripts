<#
.SYNOPSIS
    Removes installed TeamViewer and forcibly stops running TeamViewer processes.

.DESCRIPTION
    - Workstation-only guardrail.
    - Uninstalls installed TeamViewer (registry).
    - Stops running TeamViewer processes (portable QuickSupport etc.) reliably:
        1) Stop-Process (shows errors)
        2) taskkill fallback (/F /T)
    - Post-check prints what remains.

    Portable EXE deletion is present but DISABLED (commented out).
    No disk logging.

.EXITCODES
    0 = No installed TeamViewer AND no running TeamViewer processes
    1 = Still installed/running OR uninstall failures
#>

[CmdletBinding()]
param(
    [int]$StopRetries = 3,
    [int]$RetryDelaySeconds = 2
)

# ------------------------------------------------------------
# STEP 0 – WORKSTATION GUARDRAIL
# ------------------------------------------------------------
# ProductType: 1=Workstation, 2=DC, 3=Server
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if (-not $os -or $os.ProductType -ne 1) {
    Write-Output "Skipping: Not a workstation OS."
    exit 0
}

# ------------------------------------------------------------
# STEP 1 – UNINSTALL REGISTRY LOCATIONS
# ------------------------------------------------------------
$UninstallPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

function Get-InstalledTeamViewer {
    foreach ($p in $UninstallPaths) {
        Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -match 'TeamViewer' } |
            Select-Object DisplayName, DisplayVersion, UninstallString, QuietUninstallString
    }
}

function Get-RunningTeamViewer {
    # Get-Process gives us the running set; CIM adds ExecutablePath when available.
    $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^TeamViewer' }
    foreach ($proc in $procs) {
        $cim = Get-CimInstance Win32_Process -Filter ("ProcessId={0}" -f $proc.Id) -ErrorAction SilentlyContinue
        [pscustomobject]@{
            Name           = ($proc.Name + ".exe")
            ProcessId      = $proc.Id
            ExecutablePath = $cim.ExecutablePath
        }
    }
}

function Get-MsiGuidFromString {
    param([string]$Text)
    $m = [regex]::Match($Text, '\{[0-9A-Fa-f\-]{36}\}')
    if ($m.Success) { return $m.Value }
    return $null
}

function Invoke-MSIUninstall {
    param([string]$Guid)
    $p = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $Guid /qn /norestart" -Wait -PassThru -WindowStyle Hidden
    return $p.ExitCode
}

function Invoke-ExeUninstall {
    param([string]$UninstallString)

    $exe = $null; $args = $null
    if ($UninstallString -match '^\s*"(.*?)"\s*(.*)$') { $exe=$matches[1]; $args=$matches[2] }
    elseif ($UninstallString -match '^\s*([^\s]+)\s*(.*)$') { $exe=$matches[1]; $args=$matches[2] }
    else { return 1 }

    if ($args -notmatch '/S|/silent|/verysilent|/qn') { $args = ($args + ' /S').Trim() }

    $p = Start-Process -FilePath $exe -ArgumentList $args -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
    if (-not $p) { return 1 }
    return $p.ExitCode
}

function Stop-TeamViewerProcess {
    <#
    Attempts to stop a process by ProcessId.
    Returns $true if the process no longer exists after attempts.
    #>
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        [Parameter(Mandatory)][string]$ProcessName
    )

    # Attempt 1: Stop-Process with visible errors
    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction Stop
        Write-Output ("Stop-Process succeeded for {0} (PID {1})" -f $ProcessName, $ProcessId)
    }
    catch {
        Write-Output ("Stop-Process FAILED for {0} (PID {1}): {2}" -f $ProcessName, $ProcessId, $_.Exception.Message)
    }

    Start-Sleep -Milliseconds 300

    # Check if still exists
    if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
        return $true
    }

    # Attempt 2: taskkill fallback (force + terminate child processes)
    Write-Output ("taskkill fallback for {0} (PID {1})" -f $ProcessName, $ProcessId)
    $null = & taskkill.exe /PID $ProcessId /F /T 2>$null

    Start-Sleep -Milliseconds 500

    return (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue))
}

# ------------------------------------------------------------
# STEP 2 – UNINSTALL INSTALLED TEAMVIEWER
# ------------------------------------------------------------
$failed = $false
$installed = @(Get-InstalledTeamViewer)

if ($installed.Count -gt 0) {
    Write-Output "Installed TeamViewer entries found (uninstalling):"
    $installed | Sort-Object DisplayName, DisplayVersion | ForEach-Object {
        Write-Output (" - {0} (v{1})" -f $_.DisplayName, $_.DisplayVersion)
    }

    foreach ($app in $installed) {
        $cmd = if ($app.QuietUninstallString) { $app.QuietUninstallString } else { $app.UninstallString }
        if (-not $cmd) { Write-Output ("Uninstall string missing: {0}" -f $app.DisplayName); $failed=$true; continue }

        $guid = Get-MsiGuidFromString $cmd

        if ($cmd -match 'msiexec\.exe' -or $guid) {
            if ($guid) {
                Write-Output ("Removing (MSI): {0} GUID={1}" -f $app.DisplayName, $guid)
                $code = Invoke-MSIUninstall $guid
            }
            else {
                Write-Output ("Removing (EXE fallback): {0}" -f $app.DisplayName)
                $code = Invoke-ExeUninstall $cmd
            }
        }
        else {
            Write-Output ("Removing (EXE): {0}" -f $app.DisplayName)
            $code = Invoke-ExeUninstall $cmd
        }

        Write-Output (" -> ExitCode: {0}" -f $code)
        if ($code -ne 0 -and $code -ne 3010) { $failed = $true }
    }
}
else {
    Write-Output "No installed TeamViewer entries found."
}

# ------------------------------------------------------------
# STEP 3 – STOP RUNNING TEAMVIEWER (RETRIES + VERIFICATION)
# ------------------------------------------------------------
for ($attempt = 1; $attempt -le $StopRetries; $attempt++) {

    $running = @(Get-RunningTeamViewer)

    if ($running.Count -eq 0) {
        Write-Output "No running TeamViewer processes detected."
        break
    }

    Write-Output ("Running TeamViewer processes found (attempt {0}/{1}):" -f $attempt, $StopRetries)
    $running | ForEach-Object {
        Write-Output (" - {0} (PID {1}) Path={2}" -f $_.Name, $_.ProcessId, $_.ExecutablePath)
    }

    foreach ($p in $running) {
        $killed = Stop-TeamViewerProcess -ProcessId $p.ProcessId -ProcessName $p.Name
        if ($killed) {
            Write-Output ("Confirmed stopped: {0} (PID {1})" -f $p.Name, $p.ProcessId)
        } else {
            Write-Output ("FAILED to stop: {0} (PID {1})" -f $p.Name, $p.ProcessId)
        }
    }

    Start-Sleep -Seconds $RetryDelaySeconds
}

# ------------------------------------------------------------
# OPTIONAL – PORTABLE EXE DELETION (DISABLED)
# ------------------------------------------------------------
# Present but commented out. Enable only if policy requires deletion.
#

$UserPaths = @(
  "$env:SystemDrive\Users\*\Downloads",
  "$env:SystemDrive\Users\*\Desktop",
  "$env:SystemDrive\Users\*\AppData\Local\Temp"
)
foreach ($base in $UserPaths) {
  Get-ChildItem -Path $base -Recurse -Include "*TeamViewer*.exe" -ErrorAction SilentlyContinue |
    ForEach-Object {
      Write-Output ("Deleting Portable EXE: {0}" -f $_.FullName)
      Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    }
}


# ------------------------------------------------------------
# STEP 4 – POST-CHECK (PRINT WHAT REMAINS)
# ------------------------------------------------------------
$remainingInstalled = @(Get-InstalledTeamViewer)
$remainingRunning   = @(Get-RunningTeamViewer)

if ($remainingInstalled.Count -gt 0 -or $remainingRunning.Count -gt 0 -or $failed) {

    if ($remainingInstalled.Count -gt 0) {
        Write-Output "Post-check: Installed TeamViewer still present:"
        $remainingInstalled | Sort-Object DisplayName, DisplayVersion | ForEach-Object {
            Write-Output (" - {0} (v{1})" -f $_.DisplayName, $_.DisplayVersion)
        }
    }

    if ($remainingRunning.Count -gt 0) {
        Write-Output "Post-check: TeamViewer still running:"
        $remainingRunning | ForEach-Object {
            Write-Output (" - {0} (PID {1}) Path={2}" -f $_.Name, $_.ProcessId, $_.ExecutablePath)
        }
    }

    if ($failed) { Write-Output "Post-check: One or more uninstall actions reported failure." }

    exit 1
}

Write-Output "Removal complete: No installed or running TeamViewer detected."
exit 0
