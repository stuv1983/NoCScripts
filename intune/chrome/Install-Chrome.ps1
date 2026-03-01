<#
.SYNOPSIS
    Google Chrome Phase 2 Surgical Remediation & Sanitization
.DESCRIPTION
    Executes a smart migration from unmanaged/32-bit Chrome to Enterprise 64-bit MSI.
    - Accepts System-Level EXEs as valid; focuses strictly on enabling auto-updates.
    - Synchronously deploys the MSI over existing installs if the version falls below the security floor.
    - Purges AppData Shadow IT binaries (Protects User Data).
    - Includes a 45-minute "Patient Process Gate" to avoid interrupting active users.
#>

$ErrorActionPreference = "Stop"
$MsiName = "googlechromestandaloneenterprise64.msi"
$ChromeSystem64 = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$Chromex86 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
$MinimumVersion = [version]"145.0.7632.117" # Match this exactly to your Detect-Chrome.ps1 floor

$NeedsCleanup = $false
$NeedsMSI = $false

Write-Output "=== Starting Google Chrome Phase 2 Remediation ==="

# ------------------------------------------------------------------------
# 1. EVALUATE REMEDIATION REQUIREMENTS
# ------------------------------------------------------------------------
Write-Output "[Check] Evaluating remediation scope..."

# Trap 1: Legacy 32-bit architecture detected (Requires Purge AND Install)
if (Test-Path $Chromex86) { 
    Write-Output " -> Flagged: 32-bit architecture found."
    $NeedsCleanup = $true; $NeedsMSI = $true 
}

# Trap 2: AppData Shadow IT detected (Requires Purge only)
$ExcludedProfiles = @('Public', 'Default', 'Default User', 'All Users')
$UserProfiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin $ExcludedProfiles }

foreach ($Profile in $UserProfiles) {
    if (Test-Path (Join-Path $Profile.FullName "AppData\Local\Google\Chrome\Application\chrome.exe")) { 
        Write-Output " -> Flagged: Unmanaged AppData binary found in $($Profile.Name)."
        $NeedsCleanup = $true; break 
    }
}

# Trap 3: Missing Core Application (Requires Install)
if (-not (Test-Path $ChromeSystem64)) { 
    Write-Output " -> Flagged: 64-bit System application is missing."
    $NeedsMSI = $true 
}

# Trap 4: Missing Scheduled Tasks (Requires MSI to rebuild the update engine)
$Tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { 
    $_.TaskName -match "^GoogleUpdateTaskMachine" -or 
    $_.TaskName -match "^GoogleUpdaterTaskSystem" 
})
if ($Tasks.Count -eq 0 -and (Test-Path $ChromeSystem64)) { 
    Write-Output " -> Flagged: Google Update Scheduled Tasks are missing."
    $NeedsMSI = $true 
}

# Trap 5: Outdated Version (Requires Synchronous MSI Patching for Intune Compliance)
# Note: Intune's Post-Detect runs immediately. We must force the MSI to install 
# synchronously so the new version is fully applied before the script exits.
if (Test-Path $ChromeSystem64) {
    $CurrentVersion = [version](Get-Item $ChromeSystem64).VersionInfo.ProductVersion
    if ($CurrentVersion -lt $MinimumVersion) {
        Write-Output " -> Flagged: Chrome version ($CurrentVersion) is below floor ($MinimumVersion)."
        $NeedsMSI = $true
    }
}

# ------------------------------------------------------------------------
# 2. SILENT SERVICE REPAIR (Background fix, no user impact)
# ------------------------------------------------------------------------
Write-Output "[Check] Evaluating Google Updater Service health..."
$UpdateServices = @(Get-Service -Name "gupdate", "gupdatem", "GoogleUpdater*" -ErrorAction SilentlyContinue)

foreach ($Service in $UpdateServices) {
    # If the service exists but was disabled by malware/user, silently flip it back on
    if ($Service.StartType -eq 'Disabled') {
        Write-Output "[Repair] Surgical Fix: Re-enabling disabled service $($Service.Name)..."
        Set-Service -Name $Service.Name -StartupType Automatic
        Start-Service -Name $Service.Name -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------------------
# 3. DESTRUCTIVE FIX / THE GATE / MSI EXECUTION
# ------------------------------------------------------------------------
if ($NeedsCleanup -or $NeedsMSI) {
    Write-Output "[Remediation] Environmental changes required. Entering Patient Process Gate..."
    
    # --- THE PATIENT PROCESS GATE ---
    # We wait up to 45 minutes for the user to close Chrome so we don't force-kill their work.
    $timer = [Diagnostics.Stopwatch]::StartNew()
    while (Get-Process -Name "chrome" -ErrorAction SilentlyContinue) {
        $elapsed = [math]::Round($timer.Elapsed.TotalMinutes, 1)
        Write-Output "   [Wait] Chrome is actively running. Waiting for user to close it (Elapsed: $elapsed / 45 mins)..."
        if ($timer.Elapsed.TotalMinutes -ge 45) { 
            Write-Output "   [Timeout] Maximum wait time reached. Exiting 1618 to defer to next Intune sync."
            exit 1618 
        } 
        Start-Sleep -Seconds 6
    }
    Write-Output "   [Clear] Chrome is closed. Proceeding."

    # --- SCORCHED EARTH PURGE ---
    if ($NeedsCleanup) {
        Write-Output "[Cleanup] Removing unmanaged binaries and legacy architecture..."
        
        # Remove AppData executables (Targets \Application only, protects \User Data)
        foreach ($Profile in $UserProfiles) {
            $AppDir = Join-Path $Profile.FullName "AppData\Local\Google\Chrome\Application"
            if (Test-Path $AppDir) { Remove-Item $AppDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
        
        # Remove 32-bit System architecture
        if (Test-Path $Chromex86) { Remove-Item $Chromex86 -Recurse -Force -ErrorAction SilentlyContinue }
    }

    # --- MSI EXECUTION ---
    $ExitCode = 0
    if ($NeedsMSI) {
        Write-Output "[Deployment] Executing Enterprise MSI Installation to forcefully patch device..."
        $msiPath = Join-Path $PSScriptRoot $MsiName
        # Executes synchronously (-Wait) over the top of any existing install
        $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /qn /norestart" -Wait -PassThru
        $ExitCode = $Process.ExitCode
    } else {
        Write-Output "[Deployment] 64-bit installation is already present and secure. Skipping MSI reinstall."
    }

    # ------------------------------------------------------------------------
    # 4. DESKTOP SANITIZATION & SHORTCUT REWIRING
    # ------------------------------------------------------------------------
    if ($ExitCode -eq 0 -or $ExitCode -eq 3010) {
        Write-Output "[Sanitization] Rewiring shortcuts and cleaning desktop stubs..."
        $shell = New-Object -ComObject WScript.Shell
        $newTarget = $ChromeSystem64
        $newWorkingDir = Split-Path -Path $newTarget -Parent

        # Phase A: Purge all personal user desktop icons (both fake .exe stubs and .lnk files)
        foreach ($Profile in $UserProfiles) {
            $UserDesktop = Join-Path $Profile.FullName "Desktop"
            if (Test-Path $UserDesktop) {
                # Hunt down .exe files masquerading as shortcuts
                if (Test-Path (Join-Path $UserDesktop "chrome.exe")) { Remove-Item -Path (Join-Path $UserDesktop "chrome.exe") -Force -ErrorAction SilentlyContinue }
                if (Test-Path (Join-Path $UserDesktop "Google Chrome.exe")) { Remove-Item -Path (Join-Path $UserDesktop "Google Chrome.exe") -Force -ErrorAction SilentlyContinue }
                
                # Hunt down .lnk shortcuts to prevent duplicates with the Public Desktop shortcut
                $links = Get-ChildItem -Path $UserDesktop -Filter "*.lnk" -File -ErrorAction SilentlyContinue
                foreach ($link in $links) {
                    try {
                        $shortcut = $shell.CreateShortcut($link.FullName)
                        if ($shortcut.TargetPath -match "(?i)Google\\Chrome\\Application\\chrome\.exe") {
                            Remove-Item -LiteralPath $link.FullName -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                }
            }
        }

        # Phase B: Rewire Start Menu, Taskbar, and Public Desktop to the new 64-bit path
        $searchPaths = @(
            "C:\Users\Public\Desktop\*.lnk", 
            "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\*.lnk", 
            "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\*.lnk", 
            "C:\Users\*\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*.lnk"
        )
        foreach ($link in (Get-ChildItem -Path $searchPaths -ErrorAction SilentlyContinue)) {
            try {
                $shortcut = $shell.CreateShortcut($link.FullName)
                if ($shortcut.TargetPath -match "(?i)Google\\Chrome\\Application\\chrome\.exe") {
                    $shortcut.TargetPath = $newTarget
                    $shortcut.WorkingDirectory = $newWorkingDir
                    $shortcut.Save()
                }
            } catch {}
        }
    } else {
        Write-Output "[Error] MSI Failed with Exit Code: $ExitCode"
        exit $ExitCode
    }
} else {
    Write-Output "[Complete] No destructive changes or patching required. Device is compliant."
}

Write-Output "=== Chrome Remediation Completed Successfully ==="
exit 0