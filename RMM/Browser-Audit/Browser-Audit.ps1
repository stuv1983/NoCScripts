# ============================================================
#  Browser Audit Script 
#  Checks Firefox, Chrome, and Edge installations
#  (System-wide C:\ paths AND per-user AppData paths)
# ============================================================

# Stores all browser audit findings before final output
$Results = @()

# ── Helper: get file version from an exe ──────────────────────────────────────
function Get-ExeVersion {
    param([string]$Path)

    # If the executable exists, return its ProductVersion from file metadata
    if (Test-Path $Path) {
        return (Get-Item $Path).VersionInfo.ProductVersion
    }

    # Return null if the file does not exist so callers can skip it cleanly
    return $null
}

# ── Helper: check Windows service start type ─────────────────────────────────
function Get-ServiceStatus {
    param([string[]]$ServiceNames)

    # Holds a formatted status string for each requested service
    $parts = @()

    foreach ($name in $ServiceNames) {
        # Try to get the service object; suppress errors if the service is missing
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue

        if ($svc) {
            # Query WMI to retrieve the startup mode (Auto / Manual / Disabled)
            # because Get-Service does not expose StartMode directly
            $startType = (Get-WmiObject Win32_Service -Filter "Name='$name'" -ErrorAction SilentlyContinue).StartMode

            # Build a readable status string for reporting
            $parts += "$name`: $startType ($($svc.Status))"
        } else {
            # Record that the service was not found on the device
            $parts += "$name`: Not Found"
        }
    }

    # Return all service states as a single pipe-delimited string
    return ($parts -join " | ")
}

# ── Helper: check scheduled task state and last run result ───────────────────
function Get-ScheduledTaskStatus {
    param([string[]]$TaskNames)
    $parts = @()
    foreach ($name in $TaskNames) {
        $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
        if ($task) {
            $info = Get-ScheduledTaskInfo -TaskName $name -ErrorAction SilentlyContinue
            $lastResult = if ($info.LastTaskResult -eq 0) { "OK" } else { "0x{0:X}" -f $info.LastTaskResult }
            $parts += "$name`: $($task.State) (last: $lastResult)"
        } else {
            $parts += "$name`: Not Found"
        }
    }
    return ($parts -join " | ")
}

# ── Helper: find browser in uninstall registry (both 64/32-bit hives) ────────
function Get-UninstallInfo {
    param([string]$DisplayNamePattern)
    $hives = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($hive in $hives) {
        if (-not (Test-Path $hive)) { continue }
        $match = Get-ChildItem $hive -ErrorAction SilentlyContinue |
            Get-ItemProperty -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like $DisplayNamePattern } |
            Select-Object -First 1
        if ($match) {
            return [PSCustomObject]@{
                InstallMethod = if ($match.WindowsInstaller -eq 1) { "MSI" } else { "EXE" }
                Publisher     = if ($match.Publisher) { $match.Publisher } else { "Unknown" }
                InstallSource = if ($match.InstallSource) { $match.InstallSource } else { "Unknown" }
            }
        }
    }
    return [PSCustomObject]@{ InstallMethod = "Not in Registry"; Publisher = "N/A"; InstallSource = "N/A" }
}

# ── Helper: check if a policy registry key exists and has enforced values ─────
function Get-PolicyState {
    param([string]$PolicyPath)
    if (Test-Path $PolicyPath) {
        $vals = (Get-Item $PolicyPath -ErrorAction SilentlyContinue).Property
        if ($vals -and $vals.Count -gt 0) { return "Managed ($($vals.Count) policies)" }
        return "Key exists (no values)"
    }
    return "Not Managed"
}

# ── Helper: expand AppData paths for ALL local user profiles ─────────────────
function Get-UserPaths {
    param([string]$RelativePath)

    # Stores all matching full paths found under user profiles
    $paths = @()

    # Root folder containing local Windows user profiles
    $profileRoot = "$env:SystemDrive\Users"

    if (Test-Path $profileRoot) {
        # Enumerate each profile folder under C:\Users
        Get-ChildItem $profileRoot -Directory | ForEach-Object {
            # Combine the user profile path with the relative browser executable path
            $full = Join-Path $_.FullName $RelativePath

            # Only return paths that actually exist
            if (Test-Path $full) { $paths += $full }
        }
    }

    # Return all discovered valid paths
    return $paths
}


# ════════════════════════════════════════════════════════════
#  FIREFOX
# ════════════════════════════════════════════════════════════
$firefoxExePaths = @(
    # System-wide installs
    "$env:SystemDrive\Program Files\Mozilla Firefox\firefox.exe",
    "$env:SystemDrive\Program Files (x86)\Mozilla Firefox\firefox.exe"
)

# Per-user AppData installs (Firefox rarely installs here, but possible)
$firefoxExePaths += Get-UserPaths "AppData\Local\Mozilla Firefox\firefox.exe"

$firefoxFound = $false
$firefoxUpdateSvc = Get-ServiceStatus @("MozillaMaintenance")
$firefoxUninstall = Get-UninstallInfo "*Firefox*"
$firefoxPolicy    = Get-PolicyState "HKLM:\SOFTWARE\Policies\Mozilla\Firefox"

foreach ($path in $firefoxExePaths) {
    $ver = Get-ExeVersion $path
    if ($ver) {
        $firefoxFound = $true
        $Results += [PSCustomObject]@{
            Browser          = "Firefox"
            Location         = $path
            Version          = $ver
            InstallType      = if ($path -like "*Users*") { "Per-User (AppData)" } else { "System-Wide" }
            InstallMethod    = $firefoxUninstall.InstallMethod
            Publisher        = $firefoxUninstall.Publisher
            InstallSource    = $firefoxUninstall.InstallSource
            PolicyManaged    = $firefoxPolicy
            UpdatePolicyHive = "N/A"
            UpdateService    = $firefoxUpdateSvc
            UpdateTasks      = "N/A"
        }
    }
}

if (-not $firefoxFound) {
    $Results += [PSCustomObject]@{
        Browser          = "Firefox"
        Location         = "Not Found"
        Version          = "N/A"
        InstallType      = "N/A"
        InstallMethod    = $firefoxUninstall.InstallMethod
        Publisher        = $firefoxUninstall.Publisher
        InstallSource    = $firefoxUninstall.InstallSource
        PolicyManaged    = $firefoxPolicy
        UpdatePolicyHive = "N/A"
        UpdateService    = $firefoxUpdateSvc
        UpdateTasks      = "N/A"
    }
}


# ════════════════════════════════════════════════════════════
#  GOOGLE CHROME
# ════════════════════════════════════════════════════════════
$chromeExePaths = @(
    # System-wide installs
    "$env:SystemDrive\Program Files\Google\Chrome\Application\chrome.exe",
    "$env:SystemDrive\Program Files (x86)\Google\Chrome\Application\chrome.exe"
)

# Per-user AppData installs (very common for Chrome)
$chromeExePaths += Get-UserPaths "AppData\Local\Google\Chrome\Application\chrome.exe"

$chromeFound = $false
$chromeUpdateSvc    = Get-ServiceStatus @("gupdate", "gupdatem", "GoogleChromeElevationService", "GoogleUpdaterInternalService", "GoogleUpdaterService")
$chromeUninstall    = Get-UninstallInfo "*Google Chrome*"
$chromePolicy       = Get-PolicyState "HKLM:\SOFTWARE\Policies\Google\Chrome"
$chromeUpdatePolicy = Get-PolicyState "HKLM:\SOFTWARE\Policies\Google\Update"
$chromeUpdateTasks  = Get-ScheduledTaskStatus @("GoogleUpdateTaskMachineUA", "GoogleUpdateTaskMachineCore")

foreach ($path in $chromeExePaths) {
    $ver = Get-ExeVersion $path
    if ($ver) {
        $chromeFound = $true
        $Results += [PSCustomObject]@{
            Browser           = "Chrome"
            Location          = $path
            Version           = $ver
            InstallType       = if ($path -like "*Users*") { "Per-User (AppData)" } else { "System-Wide" }
            InstallMethod     = $chromeUninstall.InstallMethod
            Publisher         = $chromeUninstall.Publisher
            InstallSource     = $chromeUninstall.InstallSource
            PolicyManaged     = $chromePolicy
            UpdatePolicyHive  = $chromeUpdatePolicy
            UpdateService     = $chromeUpdateSvc
            UpdateTasks       = $chromeUpdateTasks
        }
    }
}

if (-not $chromeFound) {
    $Results += [PSCustomObject]@{
        Browser           = "Chrome"
        Location          = "Not Found"
        Version           = "N/A"
        InstallType       = "N/A"
        InstallMethod     = $chromeUninstall.InstallMethod
        Publisher         = $chromeUninstall.Publisher
        InstallSource     = $chromeUninstall.InstallSource
        PolicyManaged     = $chromePolicy
        UpdatePolicyHive  = $chromeUpdatePolicy
        UpdateService     = $chromeUpdateSvc
        UpdateTasks       = $chromeUpdateTasks
    }
}


# ════════════════════════════════════════════════════════════
#  MICROSOFT EDGE
# ════════════════════════════════════════════════════════════
$edgeExePaths = @(
    # System-wide installs
    "$env:SystemDrive\Program Files\Microsoft\Edge\Application\msedge.exe",
    "$env:SystemDrive\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
)

# Per-user AppData installs
$edgeExePaths += Get-UserPaths "AppData\Local\Microsoft\Edge\Application\msedge.exe"

$edgeFound = $false
$edgeUpdateSvc   = Get-ServiceStatus @("edgeupdate", "edgeupdatem")
$edgeUninstall   = Get-UninstallInfo "*Microsoft Edge*"
$edgePolicy      = Get-PolicyState "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$edgeUpdateTasks = Get-ScheduledTaskStatus @("MicrosoftEdgeUpdateTaskMachineUA", "MicrosoftEdgeUpdateTaskMachineCore")

foreach ($path in $edgeExePaths) {
    $ver = Get-ExeVersion $path
    if ($ver) {
        $edgeFound = $true
        $Results += [PSCustomObject]@{
            Browser          = "Edge"
            Location         = $path
            Version          = $ver
            InstallType      = if ($path -like "*Users*") { "Per-User (AppData)" } else { "System-Wide" }
            InstallMethod    = $edgeUninstall.InstallMethod
            Publisher        = $edgeUninstall.Publisher
            InstallSource    = $edgeUninstall.InstallSource
            PolicyManaged    = $edgePolicy
            UpdatePolicyHive = "N/A"
            UpdateService    = $edgeUpdateSvc
            UpdateTasks      = $edgeUpdateTasks
        }
    }
}

if (-not $edgeFound) {
    $Results += [PSCustomObject]@{
        Browser          = "Edge"
        Location         = "Not Found"
        Version          = "N/A"
        InstallType      = "N/A"
        InstallMethod    = $edgeUninstall.InstallMethod
        Publisher        = $edgeUninstall.Publisher
        InstallSource    = $edgeUninstall.InstallSource
        PolicyManaged    = $edgePolicy
        UpdatePolicyHive = "N/A"
        UpdateService    = $edgeUpdateSvc
        UpdateTasks      = $edgeUpdateTasks
    }
}


# ════════════════════════════════════════════════════════════
#  OUTPUT
# ════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "==========================="
Write-Output "  BROWSER AUDIT RESULTS - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Output "  Device: $env:COMPUTERNAME"
Write-Output "==========================="

# Split into focused tables so columns stay readable
Write-Output "-- Install Footprint --"
$Results | Format-Table -AutoSize -Property Browser, InstallType, InstallMethod, Publisher, Version, Location

Write-Output "-- Governance & Update State --"
$Results | Format-Table -AutoSize -Property Browser, PolicyManaged, UpdatePolicyHive, UpdateService, UpdateTasks

# ── N-able friendly single-line summary (stdout) ─────────────────────────────
Write-Output "--- Summary ---"
foreach ($r in $Results) {
    if ($r.Location -ne "Not Found") {
        Write-Output "$($r.Browser) INSTALLED | Type: $($r.InstallType) | Method: $($r.InstallMethod) | Publisher: $($r.Publisher) | Version: $($r.Version) | Policy: $($r.PolicyManaged) | UpdatePolicy: $($r.UpdatePolicyHive) | UpdateSvc: $($r.UpdateService) | UpdateTasks: $($r.UpdateTasks) | Path: $($r.Location)"
    } else {
        Write-Output "$($r.Browser) NOT INSTALLED | Method: $($r.InstallMethod) | Policy: $($r.PolicyManaged) | UpdatePolicy: $($r.UpdatePolicyHive) | UpdateSvc: $($r.UpdateService) | UpdateTasks: $($r.UpdateTasks)"
    }
}

# ── Exit 0 = success (N-able expects this) ───────────────────────────────────
exit 0