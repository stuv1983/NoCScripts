<#
.SYNOPSIS
    Hardened TeamViewer Removal for Intune (SYSTEM Context)
.DESCRIPTION
    Forcibly removes all TeamViewer footprints, cleans up orphan services,
    wipes per-user AppData, clears raw registry keys across loaded user hives, 
    and deletes scheduled tasks.
#>

# Restore error visibility to ensure failures are accurately logged in Intune/RMM
$ErrorActionPreference = "Continue" 

# --- CONFIGURATION TOGGLES ---
$AllowOnServers = $false     # Set to $true to allow uninstall on Servers and Domain Controllers
$WaitIfActive = $true        # Set to $true to delay if a remote session is active
$rebootRequired = $false
$failed = $false

# --- STEP 1: WORKSTATION GUARDRAIL ---
# Prevent accidental execution on Windows Servers unless explicitly overridden.
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

if ((-not $os -or $os.ProductType -ne 1) -and -not $AllowOnServers) { 
    Write-Output "Target is not a workstation. Aborting uninstall for safety."
    exit 0 
} elseif ($os -and $os.ProductType -ne 1 -and $AllowOnServers) {
    Write-Output "Target is a server, but `$AllowOnServers is enabled. Proceeding with eradication..."
}

# --- STEP 2: OPTIONAL WAIT LOOP ---
# Prevents killing the process while in use.
if ($WaitIfActive) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    
    # Loop as long as the UI processes are running (ignores the background service)
    while ((Get-Process -Name "TeamViewer","tv_w32","tv_x64" -ErrorAction SilentlyContinue)) {
        if ($timer.Elapsed.TotalMinutes -ge 45) { 
            Write-Output "Wait limit reached. Active UI session persisted for 45 minutes. Exiting for Intune retry (Exit 1618)."
            exit 1618 
        }
        
        $elapsed = [math]::Round($timer.Elapsed.TotalMinutes, 1)
        Write-Output "Active TeamViewer UI detected. Waiting for user to disconnect... (Elapsed: $elapsed / 45 minutes)"
        Start-Sleep -Seconds 60
    }
    
    if ($timer.Elapsed.TotalSeconds -gt 0) {
        Write-Output "TeamViewer UI closed. Proceeding with uninstall."
    }
}

# --- STEP 3: PROCESS & SERVICE TERMINATION ---
# Services must be stopped first to release file locks on the binaries.
Write-Output "Stopping all TeamViewer services and killing processes..."
Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue | Stop-Service -Force -ErrorAction SilentlyContinue 

$tvProcs = @("TeamViewer", "TeamViewer_Service", "tv_w32", "tv_x64")
foreach ($p in $tvProcs) {
    if (Get-Process -Name $p -ErrorAction SilentlyContinue) {
        Stop-Process -Name $p -Force -ErrorAction SilentlyContinue
        
        # Give the OS a moment to gracefully close handles
        Wait-Process -Name $p -Timeout 2 -ErrorAction SilentlyContinue
        
        # taskkill fallback handles locked processes that resist standard PS termination
        if (Get-Process -Name $p -ErrorAction SilentlyContinue) { 
            & taskkill.exe /IM "$p.exe" /F /T 2>$null 
        }
    }
}

# --- STEP 4: REGISTRY UNINSTALLATION EXECUTION ---
# Parses the local machine registry to find the official uninstall strings and executes them silently.
Write-Output "Executing official uninstallers..."
$regPaths = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$apps = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match "^TeamViewer\b" }

foreach ($app in $apps) {
    $cmd = if ($app.QuietUninstallString) { $app.QuietUninstallString } else { $app.UninstallString }
    if (-not $cmd) { $failed = $true; continue }

    try {
        $guidMatch = [regex]::Match($cmd, '\{[0-9A-Fa-f\-]{36}\}')
        if ($guidMatch.Success) {
            # MSI Execution
            $proc = Start-Process "msiexec.exe" -ArgumentList "/x $($guidMatch.Value) /qn /norestart" -Wait -PassThru
        } else {
            # EXE Execution: Advanced parsing to handle quoted paths, unquoted paths with spaces, and missing silent switches
            if ($cmd -match '^\s*"(.*?)"\s*(.*)$') { 
                $exe = $matches[1]; $args = $matches[2] 
            } elseif ($cmd -match '^(.:\\[^\s]+\.exe)\s*(.*)$') {
                $exe = $matches[1]; $args = $matches[2]
            } else { 
                $parts = $cmd -split ' ',2; $exe = $parts[0]; $args = $parts[1] 
            }
            
            if ($args -notmatch '(?i)/S|/silent|/qn') { $args = "$args /S" }
            $proc = Start-Process $exe -ArgumentList $args.Trim() -Wait -PassThru
        }
        
        # Track 3010 to inform Intune that a reboot is pending to clear file locks
        if ($proc.ExitCode -eq 3010) { $rebootRequired = $true }
        elseif ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 1605) { $failed = $true }
    } catch {
        Write-Error "Critical failure during uninstaller execution: $($_.Exception.Message)"
        $failed = $true
    }
}

# --- STEP 5: SCHEDULED TASK CLEANUP ---
Write-Output "Cleaning up TeamViewer scheduled tasks..."
Get-ScheduledTask -ErrorAction SilentlyContinue | 
    Where-Object { $_.TaskName -like "*TeamViewer*" } | 
    Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

# --- STEP 6: FILESYSTEM CLEANUP ---
# Manual deletion of remnant folders prevents "Ghost" detections on subsequent Intune syncs.
Write-Output "Cleaning up remaining directories and AppData..."
$remnantFolders = @("$env:ProgramFiles\TeamViewer", "${env:ProgramFiles(x86)}\TeamViewer", "$env:ProgramData\TeamViewer")
foreach ($f in $remnantFolders) { 
    if (Test-Path $f) { Remove-Item $f -Recurse -Force -ErrorAction SilentlyContinue } 
}

# Clean per-user AppData to eliminate Consumer installs. Traversing C:\Users works in SYSTEM context.
Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    Remove-Item (Join-Path $_.FullName "AppData\Local\TeamViewer") -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $_.FullName "AppData\Roaming\TeamViewer") -Recurse -Force -ErrorAction SilentlyContinue
}

# --- STEP 7: REGISTRY REMNANT CLEANUP (SYSTEM CONTEXT AWARE) ---
Write-Output "Removing leftover TeamViewer registry keys (HKLM and loaded User Hives)..."

# Clean Machine-wide key
Remove-Item -Path "HKLM:\SOFTWARE\TeamViewer" -Recurse -Force -ErrorAction SilentlyContinue

# Because Intune runs as SYSTEM, 'HKCU' is the SYSTEM profile. 
# We iterate through HKEY_USERS to clean actively loaded human user profiles.
# The regex 'S-1-5-21-[\d\-]+$' ensures we only target actual user SIDs, skipping system/service accounts.
$userHives = Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue | 
             Where-Object { $_.PSChildName -match 'S-1-5-21-[\d\-]+$' }

foreach ($hive in $userHives) {
    $userKey = "Registry::HKEY_USERS\$($hive.PSChildName)\Software\TeamViewer"
    if (Test-Path $userKey) {
        Write-Output "Removing TeamViewer registry key for user SID: $($hive.PSChildName)"
        Remove-Item -Path $userKey -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- STEP 8: ORPHAN SERVICE DELETION ---
# Force-deleting services that registry uninstallers might have missed. 
Write-Output "Checking for orphaned services..."
$svcs = Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue
if ($svcs) {
    foreach ($svc in $svcs) {
        Write-Output "Deleting orphaned service: $($svc.Name)"
        & sc.exe delete $svc.Name
    }
}

# --- STEP 9: FINAL VERIFICATION ---
# If any of these conditions are true, the eradication failed and Intune should mark it as an error.
if ((Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue) -or 
    (Get-Process -Name "TeamViewer" -ErrorAction SilentlyContinue) -or 
    (Test-Path "$env:ProgramFiles\TeamViewer") -or 
    $failed) {
    Write-Error "Post-uninstall verification failed. TeamViewer footprint still remains."
    exit 1
}

# Return the appropriate code for Intune's reboot-handling engine.
Write-Output "TeamViewer successfully eradicated."
if ($rebootRequired) { exit 3010 }
exit 0