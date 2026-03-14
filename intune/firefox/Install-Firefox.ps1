<#
.SYNOPSIS
    Mozilla Firefox Remediation
.DESCRIPTION
    Upgrades Firefox to the latest Enterprise MSI, preserves each user's active
    profile, suppresses first-run onboarding, and applies enterprise policies.
    Handles both Program Files and per-user AppData installs.
#>

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
$Firefox64Exe = "C:\Program Files\Mozilla Firefox\firefox.exe"
$Firefox86Dir = "C:\Program Files (x86)\Mozilla Firefox"

# Per-user self-install locations treated as unmanaged and removed during cleanup
$RoguePaths = @(
    "AppData\Local\Mozilla Firefox",
    "AppData\Local\Programs\Mozilla Firefox"
)

# CRC hashes of the Firefox exe path - used as section headers in installs.ini/profiles.ini.
# Hardcoded because they are deterministic for standard install paths and never change.
# Ref: https://support.mozilla.org/en-US/kb/understanding-depth-profile-installation
#   308046B0AF4A39CB = C:\Program Files\Mozilla Firefox\firefox.exe       (64-bit)
#   E7CF176E110C211B = C:\Program Files (x86)\Mozilla Firefox\firefox.exe (32-bit)
$KnownHashes = @("308046B0AF4A39CB", "E7CF176E110C211B")

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

function Get-FirefoxActiveProfile {
    # Returns a hashtable { Name, Path } for the currently active Firefox profile.
    # Reads the [InstallHASH] Default= value first (Firefox 67+), then resolves it
    # to a [ProfileN] Name=. Falls back to the Default=1 [ProfileN] for older installs.
    # Returning Path= allows deterministic restore without relying on name-matching.
    param([string]$IniContent)

    $activePath = $null

    # Modern: read Default= from the first [InstallHASH] section found
    foreach ($sec in [regex]::Matches($IniContent, '(?ms)^\[Install[^\]]+\][^\[]+')) {
        if ($sec.Value -match '(?m)^Default=(.+)') { $activePath = $matches[1].Trim(); break }
    }

    # Resolve to Name and Path from the matching [ProfileN] block
    foreach ($sec in [regex]::Matches($IniContent, '(?ms)^\[Profile\d+\][^\[]+')) {
        if ($activePath -and $sec.Value -match "(?m)^Path=$([regex]::Escape($activePath))") {
            if ($sec.Value -match '(?m)^Name=(.+)') { return @{ Name = $matches[1].Trim(); Path = $activePath } }
        }
        # Legacy fallback: no [InstallHASH] present, use Default=1
        if (-not $activePath -and $sec.Value -match '(?m)^Default=1') {
            $name = if ($sec.Value -match '(?m)^Name=(.+)') { $matches[1].Trim() } else { $null }
            $path = if ($sec.Value -match '(?m)^Path=(.+)') { $matches[1].Trim() } else { $null }
            if ($name) { return @{ Name = $name; Path = $path } }
        }
    }

    return $null
}

function Build-InstallHashBlock {
    # Builds formatted INI text for [InstallHASH] sections.
    # Locked=1 prevents Firefox from reassigning Default= on next launch.
    # Both profiles.ini and installs.ini must contain identical entries -
    # Firefox reads installs.ini first and creates default-release if a hash is missing.
    param([string[]]$Hashes, [string]$ProfileRelPath)

    $block = ""
    foreach ($hash in $Hashes) {
        $block += "[$hash]`r`nDefault=$ProfileRelPath`r`nLocked=1`r`n`r`n"
    }
    return $block.TrimEnd()
}

function Update-ProfilesIni {
    # Updates profiles.ini in a single read/write:
    #   1. Moves Default=1 to the target [ProfileN] block
    #   2. Sets StartWithLastProfile=1 (inserts if absent) to skip the profile picker
    #   3. Replaces all [InstallHASH] sections with fresh ones for all known hashes
    # Ref: https://support.mozilla.org/en-US/kb/understanding-depth-profile-installation
    param([string]$IniPath, [string]$ProfileRelPath, [string[]]$Hashes)

    $ini = Get-Content $IniPath -Raw

    # Strip Default=1 from all blocks, then set it on the target block only
    $ini = [regex]::Replace($ini, '(?m)^Default=1\r?\n', '')
    foreach ($sec in [regex]::Matches($ini, '(?ms)^\[Profile\d+\][^\[]+')) {
        if ($sec.Value -match "(?m)^Path=$([regex]::Escape($ProfileRelPath))") {
            $ini = $ini.Replace($sec.Value, ($sec.Value.TrimEnd() + "`r`nDefault=1`r`n"))
            break
        }
    }

    # Update StartWithLastProfile in-place, or insert it under [General] if missing
    if ($ini -match '(?m)^StartWithLastProfile=\d') {
        $ini = [regex]::Replace($ini, '(?m)^StartWithLastProfile=\d', 'StartWithLastProfile=1')
    } elseif ($ini -match '(?m)^\[General\]') {
        $ini = [regex]::Replace($ini, '(?m)^\[General\]', "[General]`r`nStartWithLastProfile=1")
    }

    # Replace all existing [InstallHASH] sections with fresh ones for all known hashes
    $ini = [regex]::Replace($ini, '(?ms)^\[Install[^\]]+\][^\[]+', '')
    $ini = [regex]::Replace($ini, '(\r?\n){3,}', "`r`n`r`n")
    $ini = $ini.TrimEnd() + "`r`n`r`n" + (Build-InstallHashBlock -Hashes $Hashes -ProfileRelPath $ProfileRelPath)

    Set-Content -Path $IniPath -Value $ini.TrimEnd() -Encoding UTF8 -NoNewline
}

# ---------------------------------------------------------------------------
# INIT
# ---------------------------------------------------------------------------

$LogDir = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
Start-Transcript -Path "$LogDir\Firefox_Remediation.log" -Append -Force

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$MsiFile = Get-ChildItem -Path $ScriptDir -Filter "Firefox Setup*.msi" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $MsiFile) { Write-Output "[Error] No Firefox MSI found."; Stop-Transcript; exit 1 }
$MsiPath = $MsiFile.FullName

Write-Output "=== Starting Firefox Remediation ==="

# ---------------------------------------------------------------------------
# 1. GATE: Wait for Firefox to close
#    MSI will fail if Firefox is running. Exit 1 (not 1618) on timeout -
#    1618 means "another MSI is running" which is misleading in this context.
# ---------------------------------------------------------------------------
Write-Output "[Check] Checking for active Firefox processes..."
$timer = [Diagnostics.Stopwatch]::StartNew()
while (Get-Process -Name "firefox" -ErrorAction SilentlyContinue) {
    $elapsed = [math]::Round($timer.Elapsed.TotalMinutes, 1)
    Write-Output "   [Wait] Firefox running ($elapsed / 45 mins)..."
    if ($timer.Elapsed.TotalMinutes -ge 45) {
        Write-Output "   [Timeout] Firefox still running after 45 minutes - exiting."
        Stop-Transcript; exit 1
    }
    Start-Sleep -Seconds 5
}

# ---------------------------------------------------------------------------
# 2. CLEANUP
# ---------------------------------------------------------------------------
Write-Output "[Cleanup] Scanning for unmanaged binaries..."

# --- 2a: Uninstall registered x86 Firefox ---
# Use the UninstallString rather than deleting the folder directly - this removes
# registry entries cleanly and prevents Windows Installer repair/self-heal.
$x86UninstallKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                                  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" `
    -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*Mozilla Firefox*" -and $_.InstallLocation -like "*Program Files (x86)*" } |
    Select-Object -First 1

if ($x86UninstallKey) {
    Write-Output " -> Found registered x86 Firefox: $($x86UninstallKey.DisplayName)"
    $uninstallStr = $x86UninstallKey.UninstallString
    if ($uninstallStr) {
        if ($uninstallStr -match '\{[A-F0-9\-]+\}') {
            # MSI product - uninstall via GUID
            $guid = $matches[0]
            Write-Output " -> Uninstalling via MSI GUID: $guid"
            $u = Start-Process msiexec.exe -ArgumentList "/x $guid /qn /norestart" -Wait -PassThru
            if ($u.ExitCode -eq 0 -or $u.ExitCode -eq 3010) {
                Write-Output " -> Uninstalled cleanly (exit $($u.ExitCode))"
            } else {
                Write-Output " -> Warning: MSI uninstall exit $($u.ExitCode) - will attempt folder removal"
            }
        } elseif ($uninstallStr -match '(?i)helper\.exe') {
            # Exe-based product (AppData self-installs use helper.exe /S)
            $exePath = if ($uninstallStr -match '^"([^"]+)"') { $matches[1] } else { $uninstallStr.Split(' ')[0] }
            if (Test-Path $exePath) {
                Write-Output " -> Uninstalling via exe: $exePath"
                Start-Process $exePath -ArgumentList "/S" -Wait | Out-Null
                Write-Output " -> Exe uninstaller completed"
            } else {
                Write-Output " -> Warning: uninstaller not found at $exePath"
            }
        } else {
            Write-Output " -> Warning: unrecognised UninstallString: $uninstallStr"
        }
    }
}

# Remove folder if still present (handles unregistered copies or failed uninstall)
if (Test-Path $Firefox86Dir) {
    Remove-Item $Firefox86Dir -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $Firefox86Dir) { Write-Output "[Fatal Error] Failed to remove x86 dir."; Stop-Transcript; exit 1 }
    Write-Output " -> Removed x86 installation directory."
}

# Get all real user accounts (NTUSER.DAT distinguishes users from system folders)
$UserProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { Test-Path (Join-Path $_.FullName "NTUSER.DAT") }

# --- 2b: Clean up HKCU AppData Firefox registrations ---
# AppData installs register in HKCU, not HKLM, so the scan above misses them.
# Their HKCU-registered scheduled tasks run after our script and overwrite installs.ini.
# We access each user's hive: via HKU\<SID> if logged in, or by loading NTUSER.DAT if not.
foreach ($WinUser in $UserProfiles) {
    $hiveFile = Join-Path $WinUser.FullName "NTUSER.DAT"
    $userSid  = try {
        (New-Object System.Security.Principal.NTAccount($WinUser.Name)).Translate(
            [System.Security.Principal.SecurityIdentifier]).Value
    } catch { $null }

    if (-not $userSid) {
        Write-Output "   Warning: Could not resolve SID for $($WinUser.Name) - HKCU cleanup skipped"
        continue
    }

    $alreadyMounted = Test-Path "Registry::HKEY_USERS\$userSid"
    $hiveMountKey   = $null
    $useTemporary   = $false

    if ($alreadyMounted) {
        $hiveMountKey = $userSid  # User is logged in - use existing mount
    } else {
        $hiveMountKey = "TempHive_$($WinUser.Name)"
        $null = & reg load "HKU\$hiveMountKey" $hiveFile 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Output "   Warning: Could not load hive for $($WinUser.Name) - skipping"
            continue
        }
        $useTemporary = $true
    }

    try {
        $hkcuPaths = @(
            "Registry::HKEY_USERS\$hiveMountKey\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "Registry::HKEY_USERS\$hiveMountKey\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        $appDataKey = Get-ChildItem $hkcuPaths -ErrorAction SilentlyContinue |
            Get-ItemProperty -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*Mozilla Firefox*" } |
            Select-Object -First 1

        if ($appDataKey) {
            Write-Output " -> Found HKCU Firefox for $($WinUser.Name): $($appDataKey.DisplayName)"
            $uStr = $appDataKey.UninstallString
            if ($uStr -and $uStr -match '(?i)helper\.exe') {
                $exePath = if ($uStr -match '^"([^"]+)"') { $matches[1] } else { $uStr.Split(' ')[0] }
                if (Test-Path $exePath) {
                    Start-Process $exePath -ArgumentList "/S" -Wait | Out-Null
                    Write-Output "   -> HKCU uninstaller completed"
                } else {
                    Write-Output "   -> Warning: HKCU uninstaller not found at $exePath"
                }
            }
        }

        # Disable per-user Firefox scheduled tasks - these overwrite installs.ini post-script
        $taskPath = "Registry::HKEY_USERS\$hiveMountKey\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\Mozilla"
        if (Test-Path $taskPath) {
            Get-ChildItem $taskPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Output "   -> Disabling per-user task: $($_.PSChildName)"
                try { Set-ItemProperty -Path $_.PSPath -Name "Index" -Value 0 -Type DWord -ErrorAction SilentlyContinue }
                catch { Write-Output "   -> Warning: Could not disable task - $_" }
            }
        }
    } catch {
        Write-Output "   Warning: HKCU cleanup failed for $($WinUser.Name) - $_"
    } finally {
        # Only unload if we loaded it - never unload a live user's hive
        if ($useTemporary) {
            [GC]::Collect(); Start-Sleep -Milliseconds 500
            $null = & reg unload "HKU\$hiveMountKey" 2>&1
        }
    }
}

# --- 2c: Remove rogue AppData Firefox folders ---
foreach ($WinUser in $UserProfiles) {
    foreach ($RelPath in $RoguePaths) {
        $AppDir = Join-Path $WinUser.FullName $RelPath
        if (Test-Path $AppDir) {
            Write-Output " -> Found rogue install for $($WinUser.Name): $AppDir"
            Remove-Item $AppDir -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path $AppDir) { Write-Output "   Warning: Could not remove $AppDir" }
            else { Write-Output "   Removed: $AppDir" }
        }
    }
}

# --- 2d: Remove existing Firefox shortcuts ---
# The MSI will recreate them at the correct 64-bit path
$shell = New-Object -ComObject WScript.Shell
$deletePaths = @(
    "C:\Users\*\Desktop\*.lnk",
    "C:\Users\Public\Desktop\*.lnk",
    "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\*.lnk",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\*.lnk"
)
try {
    Get-ChildItem -Path $deletePaths -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if ($shell.CreateShortcut($_.FullName).TargetPath -match "(?i)firefox\.exe$") {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch { Write-Output "   Warning: Could not inspect shortcut $($_.FullName) - $_" }
    }
} catch { Write-Output "   Warning: Shortcut enumeration failed - $_" }

# ---------------------------------------------------------------------------
# 3. PRE-INSTALL: Snapshot each user's active Firefox profile
#    Captured before the MSI runs - the upgrade creates new folders and rewrites
#    installs.ini, so we need the original state to restore correctly afterward.
#    Both Name and Path are saved: Path enables deterministic restore;
#    Name is the fallback if the original folder no longer exists post-upgrade.
# ---------------------------------------------------------------------------
Write-Output "[ProfileSave] Snapshotting Firefox profile configuration before upgrade..."
$ProfileSnapshots = @{}

foreach ($WinUser in $UserProfiles) {
    $FirefoxDir   = Join-Path $WinUser.FullName "AppData\Roaming\Mozilla\Firefox"
    $IniPath      = Join-Path $FirefoxDir "profiles.ini"
    $InstallsPath = Join-Path $FirefoxDir "installs.ini"
    $ProfilesDir  = Join-Path $FirefoxDir "Profiles"
    if (-not (Test-Path $IniPath)) { continue }

    try {
        $activeProfile = Get-FirefoxActiveProfile -IniContent (Get-Content $IniPath -Raw)
        if (-not $activeProfile) {
            Write-Output "   [ProfileSave] $($WinUser.Name): no active profile found - skipping."
            continue
        }

        Write-Output "   [ProfileSave] $($WinUser.Name): '$($activeProfile.Name)' ($($activeProfile.Path))"
        $ProfileSnapshots[$WinUser.FullName] = @{
            IniPath         = $IniPath
            InstallsPath    = $InstallsPath
            ProfilesDir     = $ProfilesDir
            TargetName      = $activeProfile.Name
            OriginalRelPath = $activeProfile.Path
        }
    } catch { Write-Output "   [ProfileSave] Warning: $($WinUser.Name) - $_" }
}

# ---------------------------------------------------------------------------
# 4. INSTALL MSI
#    Exit 3010 = success with reboot pending - acceptable for Intune
# ---------------------------------------------------------------------------
Write-Output "[Deployment] Executing Enterprise MSI Installation..."
$Process = Start-Process msiexec.exe -ArgumentList "/i `"$MsiPath`" ALLUSERS=1 /qn /norestart" -Wait -PassThru
if ($Process.ExitCode -ne 0 -and $Process.ExitCode -ne 3010) {
    Write-Output "[Error] MSI Exit Code: $($Process.ExitCode)"; Stop-Transcript; exit $Process.ExitCode
}

# ---------------------------------------------------------------------------
# 5. ENTERPRISE POLICIES
#    Registry only - sufficient for Intune-managed Windows devices.
#    policies.json is only needed for non-domain/non-Windows scenarios.
# ---------------------------------------------------------------------------
Write-Output "[Policies] Writing Firefox enterprise policies..."
$regBase = "HKLM:\SOFTWARE\Policies\Mozilla\Firefox"
$regDefs = @{
    "$regBase" = @{
        OverrideFirstRunPage       = [string]""  # No welcome page on first run
        OverridePostUpdatePage     = [string]""  # No "what's new" page after update
        DontCheckDefaultBrowser    = [int]1      # No "make Firefox default" prompt
        DisableFirefoxStudies      = [int]1      # No A/B experiments
        DisableTelemetry           = [int]1      # No usage data sent to Mozilla
        DisableDefaultBrowserAgent = [int]1      # Prevents agent overwriting installs.ini
        AppAutoUpdate              = [int]1      # Silent self-updates via built-in updater
    }
    "$regBase\UserMessaging" = @{
        WhatsNew                 = [int]0
        ExtensionRecommendations = [int]0
        FeatureRecommendations   = [int]0
        UrlbarInterventions      = [int]0
        MoreFromMozilla          = [int]0
    }
}

foreach ($path in $regDefs.Keys) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($name in $regDefs[$path].Keys) {
        $val      = $regDefs[$path][$name]
        $propType = if ($val -is [int]) { "DWord" } else { "String" }
        # $null -ne check required - empty string is falsy and would incorrectly trigger New-ItemProperty
        if ($null -ne (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue)) {
            Set-ItemProperty -Path $path -Name $name -Value $val -Force
        } else {
            New-ItemProperty -Path $path -Name $name -Value $val -PropertyType $propType -Force | Out-Null
        }
    }
}
Write-Output "   [Policies] Registry policies written"

# Disable machine-wide Default Browser Agent task only - leave Background Update enabled
Get-ScheduledTask -ErrorAction SilentlyContinue |
    Where-Object { $_.TaskName -like "*Firefox Default Browser Agent*" } |
    ForEach-Object {
        Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
        Write-Output "   [Policies] Disabled: $($_.TaskName)"
    }

# ---------------------------------------------------------------------------
# 6. PROFILE RESTORE
#    Restores each user's active profile in both profiles.ini and installs.ini.
#    Strategy: use original Path= captured pre-install (deterministic); fall back
#    to oldest folder matching *.ProfileName if the original no longer exists.
#    compatibility.ini is written to prevent Firefox treating this as a "profile
#    from a different install" - which would create default-release regardless of
#    what installs.ini says (critical for AppData -> Program Files migrations).
# ---------------------------------------------------------------------------
Write-Output "[ProfileRestore] Restoring Firefox profile configuration after upgrade..."

foreach ($UserPath in $ProfileSnapshots.Keys) {
    $Snap        = $ProfileSnapshots[$UserPath]
    $TargetName  = $Snap.TargetName
    $ProfilesDir = $Snap.ProfilesDir

    try {
        if (-not (Test-Path $Snap.IniPath)) {
            Write-Output "   [ProfileRestore] profiles.ini missing for $UserPath - skipping."
            continue
        }

        # Resolve profile folder: original path first, then name-match fallback
        $relPath = $null
        if ($Snap.OriginalRelPath) {
            $absPath = Join-Path (Split-Path $Snap.IniPath) $Snap.OriginalRelPath.Replace('/', '\')
            if (Test-Path $absPath) {
                $relPath = $Snap.OriginalRelPath
                Write-Output "   [ProfileRestore] Using original path: $relPath"
            } else {
                Write-Output "   [ProfileRestore] Original path not found - falling back to name match"
            }
        }
        if (-not $relPath -and (Test-Path $ProfilesDir)) {
            $tf = Get-ChildItem $ProfilesDir -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "*.$TargetName" } |
                Sort-Object CreationTime | Select-Object -First 1
            if ($tf) {
                $relPath = "Profiles/$($tf.Name)"
                Write-Output "   [ProfileRestore] Name-matched: $($tf.Name)"
            }
        }
        if (-not $relPath) {
            Write-Output "   [ProfileRestore] No folder found for '$TargetName' - skipping."
            continue
        }

        Write-Output "   [ProfileRestore] Target: '$TargetName' -> $relPath"

        Update-ProfilesIni -IniPath $Snap.IniPath -ProfileRelPath $relPath -Hashes $KnownHashes
        Write-Output "   [ProfileRestore] profiles.ini updated"

        Set-Content -Path $Snap.InstallsPath `
            -Value (Build-InstallHashBlock -Hashes $KnownHashes -ProfileRelPath $relPath) `
            -Encoding UTF8 -NoNewline
        Write-Output "   [ProfileRestore] installs.ini written"

        # Write compatibility.ini - prevents the "used by a different install" heuristic
        # from creating default-release when migrating from an AppData install path
        $profileAbsPath = Join-Path (Split-Path $Snap.IniPath) $relPath.Replace('/', '\')
        $compatContent  = "[Compatibility]`r`nLastVersion=148.0_20250401144603/148.0_20250401144603`r`nLastOSABI=winnt_x86_64-msvc`r`nLastPlatformDir=C:\Program Files\Mozilla Firefox`r`nLastAppDir=C:\Program Files\Mozilla Firefox\browser`r`n"
        Set-Content -Path (Join-Path $profileAbsPath "compatibility.ini") -Value $compatContent -Encoding UTF8 -NoNewline
        Write-Output "   [ProfileRestore] compatibility.ini written"

        # user.js is re-applied on every Firefox launch - survives self-updates
        $userJs = Join-Path $profileAbsPath "user.js"
@'
// Suppress Firefox first-run onboarding - managed by IT deployment
user_pref("browser.startup.homepage_override.mstone", "ignore");
user_pref("startup.homepage_override_url", "");
user_pref("startup.homepage_welcome_url", "");
user_pref("startup.homepage_welcome_url.additional", "");
user_pref("browser.shell.checkDefaultBrowser", false);
user_pref("browser.aboutwelcome.enabled", false);
user_pref("browser.aboutwelcome.didSeeFinalScreen", true);
user_pref("trailhead.firstrun.didSeeAboutWelcome", true);
user_pref("browser.laterrun.enabled", false);
user_pref("browser.migration.version", 999);
user_pref("datareporting.policy.firstRunURL", "");
user_pref("datareporting.policy.dataSubmissionPolicyBypassNotification", true);
'@ | Set-Content -Path $userJs -Encoding UTF8
        Write-Output "   [ProfileRestore] Done for $(Split-Path $UserPath -Leaf)"

    } catch { Write-Output "   [ProfileRestore] Warning: Failed for $UserPath - $_" }
}

# ---------------------------------------------------------------------------
# 7. SHORTCUT FIX: Redirect pinned taskbar shortcuts to the 64-bit path
#    Best-effort - pin behaviour can be temperamental depending on Windows version
# ---------------------------------------------------------------------------
Write-Output "[Sanitization] Redirecting Taskbar pins..."
$taskbarPaths = @("C:\Users\*\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*.lnk")
try {
    Get-ChildItem -Path $taskbarPaths -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $sc = $shell.CreateShortcut($_.FullName)
            if ($sc.TargetPath -match "(?i)firefox\.exe$") {
                $sc.TargetPath = $Firefox64Exe; $sc.WorkingDirectory = "C:\Program Files\Mozilla Firefox"; $sc.Save()
            }
        } catch { Write-Output "   Warning: Could not update shortcut $($_.FullName) - $_" }
    }
} catch { Write-Output "   Warning: Taskbar path enumeration failed - $_" }

# ---------------------------------------------------------------------------
# 8. RE-ASSERT installs.ini
#    MSI post-install activity can overwrite installs.ini after step 6.
#    Writing it again here as the last step before validation ensures it sticks.
# ---------------------------------------------------------------------------
Write-Output "[ProfileRestore] Re-asserting installs.ini..."
foreach ($UserPath in $ProfileSnapshots.Keys) {
    $Snap    = $ProfileSnapshots[$UserPath]
    $relPath = $null

    if ($Snap.OriginalRelPath) {
        $abs = Join-Path (Split-Path $Snap.IniPath) $Snap.OriginalRelPath.Replace('/', '\')
        if (Test-Path $abs) { $relPath = $Snap.OriginalRelPath }
    }
    if (-not $relPath -and (Test-Path $Snap.ProfilesDir)) {
        $tf = Get-ChildItem $Snap.ProfilesDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*.$($Snap.TargetName)" } |
            Sort-Object CreationTime | Select-Object -First 1
        if ($tf) { $relPath = "Profiles/$($tf.Name)" }
    }
    if ($relPath) {
        Set-Content -Path $Snap.InstallsPath `
            -Value (Build-InstallHashBlock -Hashes $KnownHashes -ProfileRelPath $relPath) `
            -Encoding UTF8 -NoNewline
        Write-Output "   [ProfileRestore] installs.ini re-asserted for $(Split-Path $UserPath -Leaf)"
    }
}

# ---------------------------------------------------------------------------
# 9. VALIDATION
#    Two tiers: critical checks exit 1 on failure; per-user checks log warnings
#    only so a single user's issue does not fail the whole deployment.
# ---------------------------------------------------------------------------
Write-Output "[Validation] Verifying installation state..."
$validationFailed = $false

# --- Critical ---
if (Test-Path $Firefox64Exe) {
    Write-Output "   [OK] firefox.exe version $((Get-Item $Firefox64Exe).VersionInfo.ProductVersion)"
} else {
    Write-Output "   [FAIL] firefox.exe not found at $Firefox64Exe"; $validationFailed = $true
}
if (Test-Path $Firefox86Dir) {
    Write-Output "   [FAIL] x86 directory still exists"; $validationFailed = $true
} else {
    Write-Output "   [OK] x86 directory absent"
}
if (Test-Path $regBase) {
    Write-Output "   [OK] Registry policies key present"
} else {
    Write-Output "   [FAIL] Registry policies key missing: $regBase"; $validationFailed = $true
}

# --- Per-user profile binding ---
foreach ($UserPath in $ProfileSnapshots.Keys) {
    $Snap       = $ProfileSnapshots[$UserPath]
    $userName   = Split-Path $UserPath -Leaf
    $userFailed = $false

    $relPath = $null
    if ($Snap.OriginalRelPath) {
        $abs = Join-Path (Split-Path $Snap.IniPath) $Snap.OriginalRelPath.Replace('/', '\')
        if (Test-Path $abs) { $relPath = $Snap.OriginalRelPath }
    }
    if (-not $relPath -and (Test-Path $Snap.ProfilesDir)) {
        $tf = Get-ChildItem $Snap.ProfilesDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*.$($Snap.TargetName)" } |
            Sort-Object CreationTime | Select-Object -First 1
        if ($tf) { $relPath = "Profiles/$($tf.Name)" }
    }
    if (-not $relPath) { Write-Output "   [WARN] $userName - could not resolve profile path"; continue }

    $profileAbsPath = Join-Path (Split-Path $Snap.IniPath) $relPath.Replace('/', '\')

    # Profile folder on disk
    if (Test-Path $profileAbsPath) {
        Write-Output "   [OK] $userName - profile folder exists: $relPath"
    } else {
        Write-Output "   [WARN] $userName - profile folder missing: $profileAbsPath"; $userFailed = $true
    }

    # installs.ini: both hashes pointing to the correct profile
    if (Test-Path $Snap.InstallsPath) {
        $installsContent = Get-Content $Snap.InstallsPath -Raw
        foreach ($hash in $KnownHashes) {
            if ($installsContent -match "(?ms)\[$hash\][^\[]*Default=$([regex]::Escape($relPath))") {
                Write-Output "   [OK] $userName - installs.ini $hash -> $relPath"
            } else {
                Write-Output "   [WARN] $userName - installs.ini hash $hash missing or wrong path"; $userFailed = $true
            }
        }
    } else {
        Write-Output "   [WARN] $userName - installs.ini not found"; $userFailed = $true
    }

    # profiles.ini: Default=1 on the correct [ProfileN] block
    if (Test-Path $Snap.IniPath) {
        $block = [regex]::Match((Get-Content $Snap.IniPath -Raw),
            "(?ms)^\[Profile\d+\][^\[]*Path=$([regex]::Escape($relPath))[^\[]*")
        if ($block.Success -and $block.Value -match '(?m)^Default=1') {
            Write-Output "   [OK] $userName - profiles.ini Default=1 for '$($Snap.TargetName)'"
        } elseif ($block.Success) {
            Write-Output "   [WARN] $userName - profiles.ini block found but Default=1 missing"; $userFailed = $true
        } else {
            Write-Output "   [WARN] $userName - profiles.ini has no block for $relPath"; $userFailed = $true
        }
    } else {
        Write-Output "   [WARN] $userName - profiles.ini not found"; $userFailed = $true
    }

    # compatibility.ini: LastPlatformDir must be Program Files (not AppData)
    $compatIni = Join-Path $profileAbsPath "compatibility.ini"
    if (Test-Path $compatIni) {
        $compatContent = Get-Content $compatIni -Raw
        if ($compatContent -match '(?m)^LastPlatformDir=C:\\Program Files\\Mozilla Firefox') {
            Write-Output "   [OK] $userName - compatibility.ini LastPlatformDir is Program Files"
        } else {
            $actual = if ($compatContent -match '(?m)^LastPlatformDir=(.+)') { $matches[1].Trim() } else { "(missing)" }
            Write-Output "   [WARN] $userName - compatibility.ini LastPlatformDir = '$actual'"; $userFailed = $true
        }
    } else {
        Write-Output "   [WARN] $userName - compatibility.ini not found"; $userFailed = $true
    }

    if (-not $userFailed) { Write-Output "   [OK] $userName - all profile checks passed" }
}

if ($validationFailed) {
    Write-Output "[Error] Critical validation failed - see above."
    Stop-Transcript; exit 1
}

Write-Output "=== Firefox Remediation Completed ==="
Stop-Transcript
exit 0
