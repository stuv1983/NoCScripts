<#
.SYNOPSIS
    Intune Win32 Install Command - Google Chrome Governance Remediation

.DESCRIPTION
    Use this as the install command for a dedicated Intune Win32 app.

    Purpose:
      - Keep the existing Chrome Enterprise deployment as the main installer
      - Use this separate Win32 app as the governance/remediation layer
      - Remove unsupported 32-bit Chrome and per-user AppData Chrome application binaries
      - Preserve user profile data/bookmarks/passwords under Chrome User Data
      - Repair Google Update policy, services and scheduled tasks
      - Force Chrome update channel policy to Stable
      - Trigger an immediate Google Update check

    Recommended package contents:
      - Install-Chrome-Governance-Remediation.ps1
      - Detect-Chrome-Governance-Remediation.ps1
      - googlechromestandaloneenterprise64.msi

    Recommended Intune install command:
      %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File .\Install-Chrome-Governance-Remediation.ps1

    Recommended Intune context:
      - Install behaviour: System
      - Run as 64-bit PowerShell
      - Restart behaviour: No specific action

.NOTES
    MSI handling strategy:
      - Chrome missing            -> MSI install
      - Chrome present, GoogleUpdate missing -> MSI repair (fvomus)
      - MSI repair fails (1603/1722) -> MSI install fallback
      - MSI install also fails    -> Exit 1 with clear message: manual cleanup required
        (Broken MSI registration can occur on test devices after aggressive component removal.
         This is NOT cleaned up automatically in production.)

    This script avoids force-closing Chrome. If Chrome is open and unsupported footprints
    need cleanup, it waits up to $MaxWaitMinutes and then exits 1618 so Intune can retry later.
#>

$ErrorActionPreference = 'Continue'

# ----------------------------
# Config
# ----------------------------
$MsiName        = 'googlechromestandaloneenterprise64.msi'
$ChromeSystem64 = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$ChromeX86Dir   = 'C:\Program Files (x86)\Google\Chrome'
$ChromeX86Exe   = Join-Path $ChromeX86Dir 'Application\chrome.exe'
$ChromeGUID     = '{8A69D345-D564-463C-AFF1-A69D9E530F96}'
$PolicyPath     = 'HKLM:\SOFTWARE\Policies\Google\Update'
$RequiredUpdateCheckMinutes = 1440
$RequiredTargetChannel = 'stable'
$ChromeUpdatePolicyName        = "Update$ChromeGUID"
$ChromeTargetChannelPolicyName = "TargetChannel$ChromeGUID"
$ChromeTargetVersionPolicyName = "TargetVersionPrefix$ChromeGUID"
$ChromeRollbackPolicyName      = "RollbackToTargetVersion$ChromeGUID"
$MaxWaitMinutes = 45

$LogRoot = 'C:\ProgramData\Kenstra\ChromeGovernance'
$LogPath = Join-Path $LogRoot 'ChromeGovernance-Install.log'

$ExcludedProfiles = @(
    'Public',
    'Default',
    'Default User',
    'All Users',
    'defaultuser0'
)

# ----------------------------
# Helpers
# ----------------------------
function Write-Step {
    param([string]$Message)

    $Line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Output $Line
    try {
        if (-not (Test-Path $LogRoot)) {
            New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $Line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

function Get-UserProfiles {
    return @(Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notin $ExcludedProfiles })
}

function Test-ChromeProcessRunning {
    return [bool](Get-Process -Name 'chrome' -ErrorAction SilentlyContinue)
}

function Wait-ForChromeToClose {
    param([int]$Minutes = 45)

    if (-not (Test-ChromeProcessRunning)) {
        return
    }

    $Timer = [Diagnostics.Stopwatch]::StartNew()
    while (Test-ChromeProcessRunning) {
        $Elapsed = [math]::Round($Timer.Elapsed.TotalMinutes, 1)
        Write-Step "Chrome is running. Waiting before cleanup. Elapsed: $Elapsed / $Minutes minutes."

        if ($Timer.Elapsed.TotalMinutes -ge $Minutes) {
            Write-Step 'Chrome remained open. Exiting 1618 so Intune can retry later without force-closing the user session.'
            exit 1618
        }

        Start-Sleep -Seconds 30
    }
}

function Set-ChromeGoogleUpdatePolicy {
    Write-Step 'Enforcing Google Update policy registry keys.'

    if (-not (Test-Path $PolicyPath)) {
        New-Item -Path $PolicyPath -Force | Out-Null
    }

    # Google Update policy, not Chrome browser policy.
    # UpdateDefault=1 and Update{ChromeGUID}=1 explicitly allow Chrome updates.
    # TargetChannel{ChromeGUID}=stable keeps the Enterprise install on the Stable channel.
    New-ItemProperty -Path $PolicyPath -Name 'UpdateDefault'                    -Value 1                          -PropertyType DWord  -Force | Out-Null
    New-ItemProperty -Path $PolicyPath -Name $ChromeUpdatePolicyName            -Value 1                          -PropertyType DWord  -Force | Out-Null
    New-ItemProperty -Path $PolicyPath -Name 'AutoUpdateCheckPeriodMinutes'     -Value $RequiredUpdateCheckMinutes -PropertyType DWord  -Force | Out-Null
    New-ItemProperty -Path $PolicyPath -Name $ChromeTargetChannelPolicyName     -Value $RequiredTargetChannel      -PropertyType String -Force | Out-Null

    # Remove version-pinning policies that would prevent auto-update.
    Remove-ItemProperty -Path $PolicyPath -Name $ChromeTargetVersionPolicyName -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $PolicyPath -Name $ChromeRollbackPolicyName      -ErrorAction SilentlyContinue

    Write-Step "UpdateDetail: Updates: Enabled | Channel: $RequiredTargetChannel | CheckPeriodMinutes: $RequiredUpdateCheckMinutes"
}

function Find-ChromeEnterpriseMsi {
    # Prefer the MSI packaged beside the Win32 install script.
    $CandidateNames = @(
        $MsiName,
        'ChromeEnterprise64.msi',
        'googlechromestandaloneenterprise64.msi'
    ) | Select-Object -Unique

    foreach ($Name in $CandidateNames) {
        $Candidate = Join-Path $PSScriptRoot $Name
        if (Test-Path $Candidate) {
            return (Resolve-Path $Candidate).Path
        }
    }

    return $null
}

function Test-GoogleUpdaterExecutablePresent {
    $UpdaterPatterns = @(
        "$env:ProgramData\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "$env:ProgramFiles\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "${env:ProgramFiles(x86)}\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "$env:ProgramFiles\Google\Update\GoogleUpdate.exe",
        "${env:ProgramFiles(x86)}\Google\Update\GoogleUpdate.exe"
    )

    foreach ($Pattern in $UpdaterPatterns) {
        if (Get-ChildItem -Path $Pattern -ErrorAction SilentlyContinue | Select-Object -First 1) {
            return $true
        }
    }

    return $false
}

function Test-GoogleUpdaterServicePresent {
    $Services = @(Get-Service -Name 'gupdate','gupdatem','GoogleUpdater*' -ErrorAction SilentlyContinue)
    return ($Services.Count -gt 0)
}

function Invoke-ChromeMsiInstall {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MsiPath
    )

    $MsiFullPath = (Resolve-Path $MsiPath).Path
    $TimeStamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
    $MsiLog      = Join-Path $LogRoot "ChromeMSI-Install-$TimeStamp.log"

    if (-not (Test-Path $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    $Arguments = "/i `"$MsiFullPath`" /qn /norestart /L*v `"$MsiLog`""

    Write-Step "Chrome MSI install starting: $MsiFullPath"
    Write-Step "Chrome MSI log: $MsiLog"

    $Process  = Start-Process -FilePath 'msiexec.exe' -ArgumentList $Arguments -Wait -PassThru -ErrorAction SilentlyContinue
    $ExitCode = $Process.ExitCode

    Write-Step "Chrome MSI install exited with code: $ExitCode"
    return $ExitCode
}

function Invoke-ChromeMsiRepair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MsiPath
    )

    $MsiFullPath = (Resolve-Path $MsiPath).Path
    $TimeStamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
    $MsiLog      = Join-Path $LogRoot "ChromeMSI-Repair-$TimeStamp.log"

    if (-not (Test-Path $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    # /fvomus: full repair - files, registry, shortcuts, services, reinstall if missing
    $Arguments = "/fvomus `"$MsiFullPath`" /qn /norestart /L*v `"$MsiLog`""

    Write-Step "Chrome MSI repair starting: $MsiFullPath"
    Write-Step "Chrome MSI log: $MsiLog"

    $Process  = Start-Process -FilePath 'msiexec.exe' -ArgumentList $Arguments -Wait -PassThru -ErrorAction SilentlyContinue
    $ExitCode = $Process.ExitCode

    Write-Step "Chrome MSI repair exited with code: $ExitCode"
    return $ExitCode
}

function Repair-ChromeMsiIfRequired {
    $MsiPath              = Find-ChromeEnterpriseMsi
    $ChromePresent        = Test-Path $ChromeSystem64
    $UpdaterExePresent    = Test-GoogleUpdaterExecutablePresent
    $UpdaterServicePresent = Test-GoogleUpdaterServicePresent

    # -------------------------------------------------------
    # Case 1: Chrome x64 is missing entirely — fresh install
    # -------------------------------------------------------
    if (-not $ChromePresent) {
        if (-not $MsiPath) {
            Write-Step "Chrome x64 is missing and MSI was not packaged beside the script. Expected: $MsiName"
            Write-Step 'Cannot rebuild Chrome or Google Update components without the Enterprise x64 MSI.'
            exit 1
        }

        Write-Step 'Chrome x64 is missing. Attempting MSI install.'
        $ExitCode = Invoke-ChromeMsiInstall -MsiPath $MsiPath

        if ($ExitCode -notin @(0, 3010)) {
            Write-Step "Chrome MSI install failed with exit code $ExitCode."
            Write-Step 'Chrome MSI registration may be orphaned. Manual cleanup required on this device.'
            Write-Step 'Suggestion: Remove orphaned MSI registration with msiexec /x or use a dedicated recovery script on test devices.'
            exit 1
        }

        Write-Step 'Chrome MSI install succeeded.'
        return
    }

    # -------------------------------------------------------
    # Case 2: Chrome present but Google Update components missing — repair first, install as fallback
    # -------------------------------------------------------
    if (-not $UpdaterExePresent -or -not $UpdaterServicePresent) {
        if (-not $MsiPath) {
            Write-Step 'Chrome x64 exists, but Google Update binaries/services are missing.'
            Write-Step "MSI was not packaged beside the script. Expected: $MsiName"
            Write-Step 'Cannot recreate Google Update components without the Enterprise x64 MSI.'
            exit 1
        }

        Write-Step 'Chrome x64 exists but Google Update components are missing. Attempting MSI repair.'
        $RepairExitCode = Invoke-ChromeMsiRepair -MsiPath $MsiPath

        if ($RepairExitCode -in @(0, 3010)) {
            Write-Step 'Chrome MSI repair succeeded.'
            return
        }

        # Repair failed — could be orphaned MSI registration (e.g. error 1722/1603 after aggressive component removal)
        Write-Step "Chrome MSI repair failed with exit code $RepairExitCode. Attempting MSI install as fallback."
        $InstallExitCode = Invoke-ChromeMsiInstall -MsiPath $MsiPath

        if ($InstallExitCode -in @(0, 3010)) {
            Write-Step 'Chrome MSI install (fallback after failed repair) succeeded.'
            return
        }

        # Both repair and install failed
        Write-Step "Chrome MSI install fallback also failed with exit code $InstallExitCode."
        Write-Step 'Chrome MSI registration appears broken. Manual cleanup required on this device.'
        Write-Step 'Likely cause: Chrome files/Google Update components were removed while Windows Installer still has Chrome registered.'
        Write-Step 'Suggestion: Remove orphaned MSI registration on this device before re-running remediation.'
        exit 1
    }

    # -------------------------------------------------------
    # Case 3: Everything present — no MSI action needed
    # -------------------------------------------------------
    Write-Step 'Chrome x64 and Google Update components are present. MSI repair/install not required.'
}

function Get-UnsupportedChromeFootprint {
    $Findings = New-Object System.Collections.Generic.List[string]

    if (Test-Path $ChromeX86Exe) {
        $Findings.Add("32-bit Chrome executable found: $ChromeX86Exe")
    } elseif (Test-Path $ChromeX86Dir) {
        $X86AppContent = Get-ChildItem -Path $ChromeX86Dir -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '\\Application\\' } |
            Select-Object -First 1
        if ($X86AppContent) {
            $Findings.Add("32-bit Chrome application footprint found: $ChromeX86Dir")
        }
    }

    foreach ($Profile in (Get-UserProfiles)) {
        $UserChromeExe = Join-Path $Profile.FullName 'AppData\Local\Google\Chrome\Application\chrome.exe'
        if (Test-Path $UserChromeExe) {
            $Findings.Add("Per-user AppData Chrome binary found for profile '$($Profile.Name)': $UserChromeExe")
        }
    }

    return $Findings
}

function Invoke-GoogleChromeUninstallString {
    param(
        [Parameter(Mandatory = $true)][string]$UninstallString,
        [Parameter(Mandatory = $true)][string]$ScopeDescription
    )

    try {
        $Command = $UninstallString.Trim()

        if ($Command -match '(?i)msiexec') {
            $ProductCode = [regex]::Match($Command, '\{[0-9A-Fa-f\-]{36}\}').Value
            if ($ProductCode) {
                Write-Step "Attempting MSI uninstall for ${ScopeDescription} product ${ProductCode}"
                $Process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $ProductCode /qn /norestart" -Wait -PassThru -ErrorAction SilentlyContinue
                Write-Step "MSI uninstall for ${ScopeDescription} exited with code: $($Process.ExitCode)"
            }
        } elseif ($Command -match '(?i)setup\.exe') {
            $Exe = [regex]::Match($Command, '"([^"]*setup\.exe)"').Groups[1].Value
            if (-not $Exe) {
                $Exe = ($Command -split '\s+')[0]
            }

            if (Test-Path $Exe) {
                Write-Step "Attempting setup.exe uninstall for ${ScopeDescription} using: $Exe"
                $Args = '--uninstall --system-level --multi-install --chrome --force-uninstall'
                $Process = Start-Process -FilePath $Exe -ArgumentList $Args -Wait -PassThru -ErrorAction SilentlyContinue
                Write-Step "setup.exe uninstall for ${ScopeDescription} exited with code: $($Process.ExitCode)"
            }
        }
    } catch {
        Write-Step "Uninstall attempt warning for ${ScopeDescription}: $($_.Exception.Message)"
    }
}

function Remove-X86Chrome {
    Write-Step 'Checking/removing 32-bit Chrome footprint.'

    $UninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $X86Entries = @(Get-ItemProperty -Path $UninstallRoots -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -eq 'Google Chrome' -and (
                ($_.InstallLocation -like 'C:\Program Files (x86)\Google\Chrome*') -or
                ($_.DisplayIcon     -like 'C:\Program Files (x86)\Google\Chrome*') -or
                ($_.UninstallString -like '*Program Files (x86)*Google*Chrome*')
            )
        })

    foreach ($Entry in $X86Entries) {
        if ($Entry.UninstallString) {
            Invoke-GoogleChromeUninstallString -UninstallString $Entry.UninstallString -ScopeDescription '32-bit Chrome'
        }
    }

    if (Test-Path $ChromeX86Dir) {
        try {
            Remove-Item -Path $ChromeX86Dir -Recurse -Force -ErrorAction Stop
            Write-Step "Removed 32-bit Chrome directory: $ChromeX86Dir"
        } catch {
            Write-Step "Warning: Could not remove 32-bit Chrome directory: $($_.Exception.Message)"
        }
    }
}

function Remove-AppDataChromeApplicationBinaries {
    Write-Step 'Removing per-user AppData Chrome Application folders only. User Data is preserved.'

    foreach ($Profile in (Get-UserProfiles)) {
        $UserChromeRoot = Join-Path $Profile.FullName 'AppData\Local\Google\Chrome'
        $ApplicationDir = Join-Path $UserChromeRoot 'Application'
        $UserDataDir    = Join-Path $UserChromeRoot 'User Data'

        if (Test-Path $ApplicationDir) {
            try {
                Remove-Item -Path $ApplicationDir -Recurse -Force -ErrorAction Stop
                Write-Step "Removed AppData Chrome Application folder for profile: $($Profile.Name)"
            } catch {
                Write-Step "Warning: Could not remove AppData Chrome Application for profile '$($Profile.Name)': $($_.Exception.Message)"
            }
        }

        # Do not remove User Data. Log it so the behaviour is clear.
        if (Test-Path $UserDataDir) {
            Write-Step "Preserved Chrome User Data for profile: $($Profile.Name)"
        }
    }
}

function Remove-StalePerUserChromeUninstallEntries {
    Write-Step 'Removing stale per-user Chrome uninstall entries where they point to AppData Chrome.'

    foreach ($ProfileKey in (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue)) {
        $Sid         = Split-Path $ProfileKey.PSPath -Leaf
        $ProfilePath = $ProfileKey.GetValue('ProfileImagePath')

        if (-not $ProfilePath -or $ProfilePath -notlike 'C:\Users\*') {
            continue
        }

        $UserHiveUninstall = "Registry::HKEY_USERS\$Sid\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        if (-not (Test-Path $UserHiveUninstall)) {
            continue
        }

        Get-ChildItem $UserHiveUninstall -ErrorAction SilentlyContinue | ForEach-Object {
            $Entry = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($Entry.DisplayName -like 'Google Chrome*' -and (
                $Entry.InstallLocation -like "$ProfilePath\AppData\Local\Google\Chrome\Application*" -or
                $Entry.UninstallString -like '*AppData\Local\Google\Chrome*' -or
                $Entry.DisplayIcon     -like '*AppData\Local\Google\Chrome*'
            )) {
                Write-Step "Removing stale per-user Chrome uninstall entry for ${ProfilePath}: $($_.PSPath)"
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Repair-GoogleUpdateServices {
    Write-Step 'Repairing Google Update services.'

    foreach ($Name in @('gupdate', 'gupdatem')) {
        $Svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
        if ($Svc) {
            try {
                if ($Name -eq 'gupdate') {
                    Set-Service -Name $Name -StartupType Automatic -ErrorAction SilentlyContinue
                } else {
                    Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue
                }
                Start-Service -Name $Name -ErrorAction SilentlyContinue
                Write-Step "Service repaired: ${Name}"
            } catch {
                Write-Step "Service repair warning for ${Name}: $($_.Exception.Message)"
            }
        } else {
            Write-Step "Service not found (may require MSI repair to recreate): ${Name}"
        }
    }

    $GoogleUpdaterServices = @(Get-Service -Name 'GoogleUpdater*' -ErrorAction SilentlyContinue)
    foreach ($Svc in $GoogleUpdaterServices) {
        try {
            if ($Svc.StartType -eq 'Disabled') {
                Set-Service -Name $Svc.Name -StartupType Manual -ErrorAction SilentlyContinue
            }
            Start-Service -Name $Svc.Name -ErrorAction SilentlyContinue
            Write-Step "Service repaired: $($Svc.Name)"
        } catch {
            Write-Step "Service repair warning for $($Svc.Name): $($_.Exception.Message)"
        }
    }
}

function Repair-GoogleUpdateTasks {
    Write-Step 'Enabling Google Update scheduled tasks.'

    $Tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        $_.TaskName -match '^GoogleUpdateTaskMachine' -or
        $_.TaskName -match '^GoogleUpdaterTaskSystem'
    })

    if ($Tasks.Count -eq 0) {
        Write-Step 'No Google Update scheduled tasks found (may require MSI repair to recreate).'
        return
    }

    foreach ($Task in $Tasks) {
        try {
            Enable-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction SilentlyContinue | Out-Null
            Write-Step "Task enabled: $($Task.TaskPath)$($Task.TaskName)"
        } catch {
            Write-Step "Task repair warning for $($Task.TaskName): $($_.Exception.Message)"
        }
    }
}

function Repair-ChromeShortcuts {
    if (-not (Test-Path $ChromeSystem64)) {
        Write-Step 'Skipping shortcut repair because x64 Chrome is not present.'
        return
    }

    Write-Step 'Rewiring Chrome shortcuts to x64 system install.'

    try {
        $Shell          = New-Object -ComObject WScript.Shell
        $NewTarget      = $ChromeSystem64
        $NewWorkingDir  = Split-Path -Path $NewTarget -Parent

        # Remove accidental desktop EXE stubs (not .lnk shortcuts).
        foreach ($Profile in (Get-UserProfiles)) {
            $UserDesktop = Join-Path $Profile.FullName 'Desktop'
            if (Test-Path $UserDesktop) {
                foreach ($Stub in @('chrome.exe', 'Google Chrome.exe')) {
                    $StubPath = Join-Path $UserDesktop $Stub
                    if (Test-Path $StubPath) {
                        Remove-Item -Path $StubPath -Force -ErrorAction SilentlyContinue
                        Write-Step "Removed desktop executable stub: $StubPath"
                    }
                }
            }
        }

        $SearchPaths = @(
            'C:\Users\Public\Desktop\*.lnk',
            'C:\Users\*\Desktop\*.lnk',
            'C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\*.lnk',
            'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\*.lnk',
            'C:\Users\*\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*.lnk'
        )

        foreach ($Link in (Get-ChildItem -Path $SearchPaths -ErrorAction SilentlyContinue)) {
            try {
                $Shortcut = $Shell.CreateShortcut($Link.FullName)
                if ($Shortcut.TargetPath -match '(?i)Google\\Chrome\\Application\\chrome\.exe') {
                    $Shortcut.TargetPath     = $NewTarget
                    $Shortcut.WorkingDirectory = $NewWorkingDir
                    $Shortcut.Save()
                    Write-Step "Rewired shortcut: $($Link.FullName)"
                }
            } catch {}
        }
    } catch {
        Write-Step "Shortcut repair warning: $($_.Exception.Message)"
    }
}

function Invoke-GoogleUpdateCheck {
    Write-Step 'Triggering Google Update check.'
    $Triggered = $false

    $ModernUpdater = @(
        "$env:ProgramData\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "$env:ProgramFiles\Google\GoogleUpdater\*\GoogleUpdater.exe"
    ) | ForEach-Object {
        Get-ChildItem -Path $_ -ErrorAction SilentlyContinue
    } | Sort-Object FullName -Descending | Select-Object -First 1

    if ($ModernUpdater) {
        Write-Step "Found GoogleUpdater.exe: $($ModernUpdater.FullName)"
        Start-Process -FilePath $ModernUpdater.FullName -ArgumentList '--update-apps' -WindowStyle Hidden -ErrorAction SilentlyContinue
        $Triggered = $true
    }

    if (-not $Triggered) {
        $LegacyPaths = @(
            "$env:ProgramFiles\Google\Update\GoogleUpdate.exe",
            "${env:ProgramFiles(x86)}\Google\Update\GoogleUpdate.exe"
        )

        foreach ($Path in $LegacyPaths) {
            if (Test-Path $Path) {
                Write-Step "Found GoogleUpdate.exe: $Path"
                Start-Process -FilePath $Path -ArgumentList '/ua /installsource scheduler' -WindowStyle Hidden -ErrorAction SilentlyContinue
                $Triggered = $true
                break
            }
        }
    }

    if (-not $Triggered) {
        Write-Step 'Warning: No Google Update executable found to trigger update check.'
    }
}

function Test-FinalCompliance {
    $Failures = New-Object System.Collections.Generic.List[string]

    if (-not (Test-Path $ChromeSystem64)) {
        $Failures.Add('Chrome x64 system install missing.')
    }

    foreach ($Finding in (Get-UnsupportedChromeFootprint)) {
        $Failures.Add($Finding)
    }

    if (-not (Test-Path $PolicyPath)) {
        $Failures.Add('Google Update policy key missing.')
    } else {
        $Policy = Get-ItemProperty -Path $PolicyPath -ErrorAction SilentlyContinue

        if ($Policy.UpdateDefault -ne 1) {
            $Failures.Add("UpdateDefault is not 1. Current: $($Policy.UpdateDefault)")
        }
        if ($Policy.$ChromeUpdatePolicyName -ne 1) {
            $Failures.Add("${ChromeUpdatePolicyName} is not 1. Current: $($Policy.$ChromeUpdatePolicyName)")
        }
        if ($Policy.AutoUpdateCheckPeriodMinutes -ne $RequiredUpdateCheckMinutes) {
            $Failures.Add("AutoUpdateCheckPeriodMinutes is not $RequiredUpdateCheckMinutes. Current: $($Policy.AutoUpdateCheckPeriodMinutes)")
        }
        if ([string]$Policy.$ChromeTargetChannelPolicyName -ne $RequiredTargetChannel) {
            $Failures.Add("${ChromeTargetChannelPolicyName} is not '$RequiredTargetChannel'. Current: $($Policy.$ChromeTargetChannelPolicyName)")
        }
        if ($null -ne $Policy.$ChromeTargetVersionPolicyName -and [string]$Policy.$ChromeTargetVersionPolicyName -ne '') {
            $Failures.Add("${ChromeTargetVersionPolicyName} is set (version pin). Current: $($Policy.$ChromeTargetVersionPolicyName)")
        }
        if ($null -ne $Policy.$ChromeRollbackPolicyName -and [string]$Policy.$ChromeRollbackPolicyName -ne '') {
            $Failures.Add("${ChromeRollbackPolicyName} is set (rollback policy). Current: $($Policy.$ChromeRollbackPolicyName)")
        }
    }

    $UpdateServices = @(Get-Service -Name 'gupdate','gupdatem','GoogleUpdater*' -ErrorAction SilentlyContinue)
    if ($UpdateServices.Count -eq 0) {
        $Failures.Add('No Google Update services found.')
    } elseif (($UpdateServices | Where-Object { $_.StartType -ne 'Disabled' } | Measure-Object).Count -eq 0) {
        $Failures.Add('Google Update services are all disabled.')
    }

    $Tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        $_.TaskName -match '^GoogleUpdateTaskMachine' -or
        $_.TaskName -match '^GoogleUpdaterTaskSystem'
    })
    if ($Tasks.Count -eq 0) {
        $Failures.Add('No Google Update scheduled tasks found.')
    } elseif (($Tasks | Where-Object { $_.State -ne 'Disabled' } | Measure-Object).Count -eq 0) {
        $Failures.Add('Google Update scheduled tasks are all disabled.')
    }

    $UpdaterExecutables = @(
        "$env:ProgramData\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "$env:ProgramFiles\Google\GoogleUpdater\*\GoogleUpdater.exe",
        "$env:ProgramFiles\Google\Update\GoogleUpdate.exe",
        "${env:ProgramFiles(x86)}\Google\Update\GoogleUpdate.exe"
    )

    $UpdaterFound = $false
    foreach ($Pattern in $UpdaterExecutables) {
        if (Get-ChildItem -Path $Pattern -ErrorAction SilentlyContinue | Select-Object -First 1) {
            $UpdaterFound = $true
            break
        }
    }

    if (-not $UpdaterFound) {
        $Failures.Add('No Google Update executable found.')
    }

    return $Failures
}

# ----------------------------
# Main
# ----------------------------
Write-Step '=== Chrome Governance Remediation starting ==='

# Step 1: MSI install/repair if Chrome or Google Update components are missing.
# This must run first — services/tasks/shortcuts depend on binaries being present.
Repair-ChromeMsiIfRequired

# Step 2: Log Chrome version after any MSI action.
if (Test-Path $ChromeSystem64) {
    try {
        $ChromeVersion = (Get-Item $ChromeSystem64).VersionInfo.ProductVersion
        Write-Step "Chrome x64 detected: ${ChromeVersion}"
    } catch {
        Write-Step "Unable to read Chrome x64 version: $($_.Exception.Message)"
    }
}

# Step 3: Enforce Google Update policy.
Set-ChromeGoogleUpdatePolicy

# Step 4: Repair services and tasks (re-enable disabled ones; cannot recreate missing without MSI).
Repair-GoogleUpdateServices
Repair-GoogleUpdateTasks

# Step 5: Remove unsupported x86/AppData Chrome footprint.
$UnsupportedFindings = @(Get-UnsupportedChromeFootprint)
if ($UnsupportedFindings.Count -gt 0) {
    Write-Step 'Unsupported Chrome footprint found:'
    foreach ($Finding in $UnsupportedFindings) {
        Write-Step " - ${Finding}"
    }

    Wait-ForChromeToClose -Minutes $MaxWaitMinutes
    Remove-AppDataChromeApplicationBinaries
    Remove-X86Chrome
} else {
    Write-Step 'No unsupported x86/AppData Chrome footprint found.'
}

# Step 6: Remove stale per-user ARP entries.
Remove-StalePerUserChromeUninstallEntries

# Step 7: Rewire shortcuts to x64 system install.
Repair-ChromeShortcuts

# Step 8: Trigger Google Update check.
Invoke-GoogleUpdateCheck

Start-Sleep -Seconds 3

# Step 9: Final compliance check to confirm all checks pass.
$FinalFailures = @(Test-FinalCompliance)
if ($FinalFailures.Count -gt 0) {
    Write-Step 'Final compliance check failed:'
    foreach ($Failure in $FinalFailures) {
        Write-Step " - ${Failure}"
    }
    Write-Step '=== Chrome Governance Remediation completed with failures ==='
    exit 1
}

Write-Step "UpdateDetail: Updates: Enabled | Channel: $RequiredTargetChannel | CheckPeriodMinutes: $RequiredUpdateCheckMinutes"
Write-Step 'Final compliance check passed.'
Write-Step '=== Chrome Governance Remediation completed successfully ==='
exit 0
