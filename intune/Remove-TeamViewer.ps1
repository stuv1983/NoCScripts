<#
.SYNOPSIS
  Intune Win32 app scripts for enforcing a "remove TeamViewer" policy.

.DESCRIPTION
  These scripts are designed for removal deployments:
    - Detection: If TeamViewer is present, return exit 1 so Intune runs the removal script.
    - Removal: Force-close TeamViewer (processes + services) and uninstall silently.
    - If TeamViewer is not present, device is compliant.

.NOTES
  Intended execution context: SYSTEM (Intune Win32 app).
  Exit codes:
    - Detection: 0 = compliant (not present), 1 = non-compliant (present)
    - Removal:   0 = removed / not present, 1 = removal failed or still detected

#>

[CmdletBinding()]
param(
    # Remove residual files after uninstall (recommended).
    [bool]$RemoveResidualFiles = $true,

    # Disable TeamViewer services before uninstall (recommended).
    [bool]$DisableServices = $true
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Write-Log([string]$Msg) {
    Write-Output "[TeamViewer-Remove] $Msg"
}

# Force-close any running TeamViewer components.
# This is intentional: RMM is the approved remote tool.
function Stop-TeamViewerProcesses {
    Write-Log "Force-closing TeamViewer processes (best-effort)..."
    Get-Process -Name "TeamViewer*" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Write-Log "Stopping process: $($_.Name) PID=$($_.Id)"
            Stop-Process -Id $_.Id -Force -ErrorAction Stop
        } catch {
            Write-Log "Process stop failed: $($_.Name) - $($_.Exception.Message)"
        }
    }
}

# Stop/disable services to prevent auto-restart and file locks during uninstall.
function Stop-Disable-TeamViewerServices {
    Write-Log "Stopping/disabling TeamViewer services (best-effort)..."
    Get-Service -Name "TeamViewer*" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if ($_.Status -ne 'Stopped') {
                Write-Log "Stopping service: $($_.Name)"
                Stop-Service -Name $_.Name -Force -ErrorAction Stop
            }
        } catch {
            Write-Log "Service stop failed: $($_.Name) - $($_.Exception.Message)"
        }

        if ($DisableServices) {
            try {
                Write-Log "Disabling service: $($_.Name)"
                Set-Service -Name $_.Name -StartupType Disabled -ErrorAction Stop
            } catch {
                Write-Log "Service disable failed: $($_.Name) - $($_.Exception.Message)"
            }
        }
    }
}

# Pull uninstall registry entries for TeamViewer products (32-bit + 64-bit).
function Get-TeamViewerUninstallEntries {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    # Starts-with match keeps it scoped to TeamViewer products
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match '^TeamViewer' -and ($_.UninstallString -or $_.QuietUninstallString) } |
        Sort-Object DisplayName -Unique
}

# Execute a silent uninstall for a single registry entry.
# Preference order:
#   1) QuietUninstallString
#   2) MSI uninstall via msiexec /x {GUID} /qn
#   3) EXE uninstall via UninstallString + /S fallback
function Invoke-SilentUninstall {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)]$App)

    $name = [string]$App.DisplayName
    $ver  = [string]$App.DisplayVersion
    $un   = [string]$App.UninstallString
    $qun  = [string]$App.QuietUninstallString

    Write-Log "Uninstalling: $name $ver"

    # Prefer QuietUninstallString if present (often already silent)
    if ($qun) {
        Write-Log "Using QuietUninstallString."
        $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$qun`"" -Wait -PassThru -WindowStyle Hidden
        return $p.ExitCode
    }

    if (-not $un) {
        Write-Log "No UninstallString present for $name (skipping)."
        return 0
    }

    # MSI installs: enforce /x /qn /norestart
    if ($un -match '(?i)msiexec(\.exe)?\s') {
        $guid = $null
        if ($un -match '\{[0-9A-Fa-f\-]{36}\}') { $guid = $Matches[0] }

        if ($guid) {
            $msiArgs = "/x $guid /qn /norestart"
            Write-Log "MSI uninstall: msiexec.exe $msiArgs"
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -WindowStyle Hidden
            return $p.ExitCode
        }

        # Fallback: transform /I -> /X and append /qn /norestart
        $cmd = $un -replace '(?i)\s/I\s', ' /X '
        if ($cmd -notmatch '(?i)/q(n|uiet)') { $cmd += ' /qn' }
        if ($cmd -notmatch '(?i)/norestart') { $cmd += ' /norestart' }

        Write-Log "MSI fallback command: $cmd"
        $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$cmd`"" -Wait -PassThru -WindowStyle Hidden
        return $p.ExitCode
    }

    # EXE uninstallers: run full commandline and append silent flags if missing
    $cmdLine = $un

    # TeamViewer commonly uses Nullsoft /S; some bundles use /SILENT
    if ($cmdLine -notmatch '(?i)\s/(S|silent|verysilent|qn|quiet)\b') {
        $cmdLine = "$cmdLine /S"
    }

    Write-Log "EXE uninstall command: $cmdLine"
    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$cmdLine`"" -Wait -PassThru -WindowStyle Hidden
    return $p.ExitCode
}

function Remove-ResidualPaths {
    if (-not $RemoveResidualFiles) { return }

    Write-Log "Removing residual TeamViewer folders (best-effort)..."

    $paths = @(
        "$env:ProgramFiles\TeamViewer",
        "${env:ProgramFiles(x86)}\TeamViewer",
        "$env:ProgramData\TeamViewer"
    )

    foreach ($p in $paths) {
        try {
            if (Test-Path $p) {
                Write-Log "Removing: $p"
                Remove-Item $p -Recurse -Force -ErrorAction Stop
            }
        } catch {
            Write-Log "Residual removal failed for ${p}: $($_.Exception.Message)"
        }
    }
}

try {
    Write-Log "Starting removal policy."

    Stop-TeamViewerProcesses
    Stop-Disable-TeamViewerServices

    $apps = Get-TeamViewerUninstallEntries
    if (-not $apps -or $apps.Count -eq 0) {
        Write-Log "No TeamViewer uninstall entries found. Already removed."
        Remove-ResidualPaths
        exit 0
    }

    $failures = @()

    foreach ($app in $apps) {
        try {
            $code = Invoke-SilentUninstall -App $app
            Write-Log "Exit code for '$($app.DisplayName)': $code"

            # 0 success, 3010 reboot required (treat as success)
            if ($code -ne 0 -and $code -ne 3010) {
                $failures += [pscustomobject]@{ Name=$app.DisplayName; ExitCode=$code }
            }
        } catch {
            Write-Log "Exception uninstalling '$($app.DisplayName)': $($_.Exception.Message)"
            $failures += [pscustomobject]@{ Name=$app.DisplayName; ExitCode='Exception' }
        }
    }

    # Close again in case a helper process restarted mid-uninstall
    Stop-TeamViewerProcesses
    Stop-Disable-TeamViewerServices

    Remove-ResidualPaths

    # Final verification: if still present, treat as failure so Intune keeps retrying
    $remaining = Get-TeamViewerUninstallEntries
    if ($remaining -and $remaining.Count -gt 0) {
        Write-Log "TeamViewer still detected after uninstall attempt:"
        $remaining | ForEach-Object { Write-Log " - $($_.DisplayName) $($_.DisplayVersion)" }
        exit 1
    }

    if ($failures.Count -gt 0) {
        Write-Log "One or more uninstallers returned non-success codes:"
        $failures | ForEach-Object { Write-Log " - $($_.Name): $($_.ExitCode)" }
        exit 1
    }

    Write-Log "Removal complete. TeamViewer not detected."
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)"
    exit 1
}