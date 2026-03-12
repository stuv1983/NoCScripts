<#
.SYNOPSIS
    Hardened TeamViewer Removal for Intune (SYSTEM Context)
.DESCRIPTION
    Forcibly removes all TeamViewer footprints: kills processes, runs official
    uninstallers, wipes per-user AppData, clears registry keys across loaded
    user hives, removes scheduled tasks, and deletes orphaned services.

    EXIT CODES:
    - 0    : Success — TeamViewer fully eradicated.
    - 1618 : Retry   — Active UI session persisted beyond wait limit. Intune will retry.
    - 3010 : Success with pending reboot — Intune will schedule a restart.
    - 1    : Failure — Post-uninstall verification found remaining footprint.

    LOGGING:
    Transcript written to:
    C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\TeamViewer-Removal.log
#>

# --- Transcript Logging ---
$LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\TeamViewer-Removal.log"
try {
    Start-Transcript -Path $LogPath -Append -Force | Out-Null
} catch {
    Write-Output "[Warning] Could not start transcript at $LogPath. Continuing without file logging."
}

# Continue on non-terminating errors so one failed step doesn't abort the entire
# removal. Critical steps log failures explicitly into $failed.
$ErrorActionPreference = "Continue"

# --- CONFIGURATION TOGGLES ---
$AllowOnServers = $false    # Set to $true to allow uninstall on Servers and Domain Controllers
$WaitIfActive   = $true     # Set to $true to wait if a remote session UI is active

$rebootRequired = $false
$failed         = $false

try {

    # --- STEP 1: WORKSTATION GUARDRAIL ---
    # Prevent accidental execution on Windows Servers unless explicitly overridden.
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

    if (-not $os) {
        Write-Output "[Warning] Could not determine OS type via WMI/CIM. Proceeding with removal anyway."
    } elseif ($os.ProductType -ne 1) {
        if (-not $AllowOnServers) {
            Write-Output "Target is not a workstation (ProductType=$($os.ProductType)). Aborting for safety."
            exit 0
        } else {
            Write-Output "[Warning] Target is a server (ProductType=$($os.ProductType)), but `$AllowOnServers is enabled. Proceeding..."
        }
    }

    Write-Output "=== Starting TeamViewer Eradication ==="

    # --- STEP 2: OPTIONAL ACTIVE-SESSION WAIT LOOP ---
    # Waits for the TeamViewer UI to close before proceeding. Ignores the background
    # service (TeamViewer_Service) — only the interactive UI processes are checked.
    if ($WaitIfActive) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        while (Get-Process -Name "TeamViewer","tv_w32","tv_x64" -ErrorAction SilentlyContinue) {
            if ($timer.Elapsed.TotalMinutes -ge 45) {
                Write-Output "Wait limit reached. Active UI session persisted for 45 minutes. Exiting 1618 for Intune retry."
                exit 1618
            }
            $elapsed = [math]::Round($timer.Elapsed.TotalMinutes, 1)
            Write-Output "Active TeamViewer UI detected. Waiting for disconnect... (Elapsed: $elapsed / 45 minutes)"
            Start-Sleep -Seconds 60
        }
        if ($timer.Elapsed.TotalSeconds -gt 0) {
            Write-Output "TeamViewer UI closed. Proceeding with uninstall."
        }
    }

    # --- STEP 3: PROCESS & SERVICE TERMINATION ---
    # Services must be stopped first to release file locks on the binaries.
    # Three-stage kill: Stop-Service → Stop-Process → taskkill fallback.
    Write-Output "[Step 3] Stopping all TeamViewer services and killing processes..."

    $svcStop = Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue
    if ($svcStop) {
        $svcStop | Stop-Service -Force -ErrorAction SilentlyContinue
        Write-Output "   -> Stopped $($svcStop.Count) service(s)."
    }

    $tvProcs = @("TeamViewer","TeamViewer_Service","tv_w32","tv_x64")
    foreach ($p in $tvProcs) {
        if (Get-Process -Name $p -ErrorAction SilentlyContinue) {
            Stop-Process -Name $p -Force -ErrorAction SilentlyContinue
            Wait-Process -Name $p -Timeout 2 -ErrorAction SilentlyContinue
            # taskkill fallback for processes that resist standard PS termination
            if (Get-Process -Name $p -ErrorAction SilentlyContinue) {
                Write-Output "   [Fallback] taskkill on resistant process: $p"
                & taskkill.exe /IM "$p.exe" /F /T 2>$null
            }
        }
    }

    # --- STEP 4: REGISTRY UNINSTALLATION ---
    # Parses HKLM uninstall keys for official uninstall strings and executes them silently.
    # Handles both MSI (GUID extraction) and EXE (quoted/unquoted path parsing) variants.
    Write-Output "[Step 4] Executing official uninstallers..."
    $regPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $apps = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match "^TeamViewer\b" }

    foreach ($app in $apps) {
        $cmd = if ($app.QuietUninstallString) { $app.QuietUninstallString } else { $app.UninstallString }
        if (-not $cmd) {
            Write-Output "   [Warning] No uninstall string found for: $($app.DisplayName)"
            $failed = $true
            continue
        }

        try {
            $guidMatch = [regex]::Match($cmd, '\{[0-9A-Fa-f\-]{36}\}')
            if ($guidMatch.Success) {
                # MSI uninstall via GUID
                Write-Output "   -> MSI uninstall: $($guidMatch.Value)"
                $proc = Start-Process "msiexec.exe" -ArgumentList "/x $($guidMatch.Value) /qn /norestart" -Wait -PassThru
            } else {
                # EXE uninstall: handle quoted paths, unquoted paths with spaces, missing silent switches
                if ($cmd -match '^\s*"(.*?)"\s*(.*)$') {
                    $exe = $matches[1]; $args = $matches[2]
                } elseif ($cmd -match '^(.:\\[^\s]+\.exe)\s*(.*)$') {
                    $exe = $matches[1]; $args = $matches[2]
                } else {
                    $parts = $cmd -split ' ', 2; $exe = $parts[0]; $args = $parts[1]
                }
                if ($args -notmatch '(?i)/S|/silent|/qn') { $args = "$args /S" }
                Write-Output "   -> EXE uninstall: $exe $($args.Trim())"
                $proc = Start-Process $exe -ArgumentList $args.Trim() -Wait -PassThru
            }

            # 3010 = success with pending reboot; 1605 = product not installed (acceptable)
            if     ($proc.ExitCode -eq 3010) { $rebootRequired = $true; Write-Output "   -> Exit 3010: Reboot pending." }
            elseif ($proc.ExitCode -eq 1605) { Write-Output "   -> Exit 1605: Product already absent (acceptable)." }
            elseif ($proc.ExitCode -ne 0)    { Write-Output "   [Error] Uninstaller exited with code: $($proc.ExitCode)"; $failed = $true }
            else                             { Write-Output "   -> Uninstaller completed successfully." }

        } catch {
            Write-Output "   [Error] Critical failure during uninstaller execution: $($_.Exception.Message)"
            $failed = $true
        }
    }

    # --- STEP 5: SCHEDULED TASK CLEANUP ---
    Write-Output "[Step 5] Removing TeamViewer scheduled tasks..."
    $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*TeamViewer*" }
    if ($tasks) {
        $tasks | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        Write-Output "   -> Removed $($tasks.Count) scheduled task(s)."
    }

    # --- STEP 6: FILESYSTEM CLEANUP ---
    # Deletes remnant directories to prevent ghost detections on subsequent Intune syncs.
    # Per-user paths cover standard consumer install, roaming variant, and TV15+ Programs variant.
    Write-Output "[Step 6] Cleaning up filesystem remnants..."

    $systemFolders = @(
        "$env:ProgramFiles\TeamViewer",
        "${env:ProgramFiles(x86)}\TeamViewer",
        "$env:ProgramData\TeamViewer"
    )
    foreach ($f in $systemFolders) {
        if (Test-Path $f) {
            Remove-Item $f -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "   -> Removed: $f"
        }
    }

    # Per-user AppData: NTUSER.DAT guard ensures we only touch real user profiles,
    # skipping Default, Public, and other system directories.
    $userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
                    Where-Object { Test-Path (Join-Path $_.FullName "NTUSER.DAT") }

    foreach ($u in $userProfiles) {
        $userPaths = @(
            (Join-Path $u.FullName "AppData\Local\TeamViewer"),
            (Join-Path $u.FullName "AppData\Roaming\TeamViewer"),
            (Join-Path $u.FullName "AppData\Local\Programs\TeamViewer")
        )
        foreach ($up in $userPaths) {
            if (Test-Path $up) {
                Remove-Item $up -Recurse -Force -ErrorAction SilentlyContinue
                Write-Output "   -> Removed per-user path: $up"
            }
        }
    }

    # --- STEP 7: REGISTRY REMNANT CLEANUP ---
    # Cleans HKLM and all actively loaded user hives.
    # HKCU is the SYSTEM profile in this context — iterate HKEY_USERS instead.
    # SID filter 'S-1-5-21-[\d\-]+$' targets real user accounts, skipping
    # system/service accounts and _Classes subkeys.
    Write-Output "[Step 7] Removing leftover TeamViewer registry keys..."

    $hklmKey = "HKLM:\SOFTWARE\TeamViewer"
    if (Test-Path $hklmKey) {
        Remove-Item -Path $hklmKey -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "   -> Removed: $hklmKey"
    }

    $userHives = Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue |
                 Where-Object { $_.PSChildName -match 'S-1-5-21-[\d\-]+$' }

    foreach ($hive in $userHives) {
        $userKey = "Registry::HKEY_USERS\$($hive.PSChildName)\Software\TeamViewer"
        if (Test-Path $userKey) {
            Remove-Item -Path $userKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "   -> Removed registry key for SID: $($hive.PSChildName)"
        }
    }

    # --- STEP 8: ORPHAN SERVICE DELETION ---
    # Deletes services that the official uninstaller may have missed.
    # sc.exe delete requires the service to be stopped — Step 3 handled that.
    # We log the sc.exe return code explicitly since it can fail silently.
    Write-Output "[Step 8] Checking for orphaned services..."
    $orphanSvcs = Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue
    if ($orphanSvcs) {
        foreach ($svc in $orphanSvcs) {
            Write-Output "   -> Deleting orphaned service: $($svc.Name)"
            $scResult = & sc.exe delete $svc.Name
            if ($LASTEXITCODE -ne 0) {
                Write-Output "   [Warning] sc.exe delete returned $LASTEXITCODE for '$($svc.Name)': $scResult"
            } else {
                Write-Output "   -> Service deleted successfully."
            }
        }
    } else {
        Write-Output "   -> No orphaned services found."
    }

    # --- STEP 9: FINAL VERIFICATION ---
    # Checks all vectors: services, processes, Program Files, AND per-user AppData.
    # Any remaining footprint combined with $failed flags the removal as incomplete.
    Write-Output "[Step 9] Running post-uninstall verification..."

    $remainingService = Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue
    $remainingProcess = Get-Process -Name "TeamViewer" -ErrorAction SilentlyContinue
    $remainingPF      = Test-Path "$env:ProgramFiles\TeamViewer"
    $remainingAppData = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
                        Where-Object { Test-Path (Join-Path $_.FullName "NTUSER.DAT") } |
                        Where-Object {
                            (Test-Path (Join-Path $_.FullName "AppData\Local\TeamViewer\TeamViewer.exe")) -or
                            (Test-Path (Join-Path $_.FullName "AppData\Roaming\TeamViewer\TeamViewer.exe")) -or
                            (Test-Path (Join-Path $_.FullName "AppData\Local\Programs\TeamViewer\TeamViewer.exe"))
                        }

    if ($remainingService) { Write-Output "   [Fail] Remaining service(s): $($remainingService.Name -join ', ')" }
    if ($remainingProcess) { Write-Output "   [Fail] Remaining process(es) still running." }
    if ($remainingPF)      { Write-Output "   [Fail] Program Files\TeamViewer still present." }
    if ($remainingAppData) { Write-Output "   [Fail] Per-user AppData remnants in: $($remainingAppData.Name -join ', ')" }

    if ($remainingService -or $remainingProcess -or $remainingPF -or $remainingAppData -or $failed) {
        Write-Output "[Fatal] Post-uninstall verification failed. TeamViewer footprint still remains."
        exit 1
    }

    Write-Output "=== TeamViewer successfully eradicated. ==="
    if ($rebootRequired) {
        Write-Output "Reboot required to clear pending file locks. Returning exit 3010."
        exit 3010
    }
    exit 0

} catch {
    Write-Output "[Unhandled Error] Script terminated unexpectedly: $_"
    exit 1

} finally {
    try { Stop-Transcript | Out-Null } catch {}
}
