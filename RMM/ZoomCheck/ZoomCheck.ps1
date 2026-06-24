#Requires -Version 5.1
<#
.SYNOPSIS
    Finds all Zoom installations on this machine and reports location + version.
.DESCRIPTION
    Searches Program Files, Program Files (x86), and every user's AppData
    (both Roaming and Local) for Zoom.exe, Zoom.msi, and other Zoom-related files,
    then reads the file version from each hit.
#>

[CmdletBinding()]
param()

# --- Get device name for audit purposes ------------------------------------
$deviceName = $env:COMPUTERNAME
$auditTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "=== Zoom Audit Report ===" -ForegroundColor Green
Write-Host "Device Name: $deviceName" -ForegroundColor Green
Write-Host "Timestamp: $auditTimestamp`n" -ForegroundColor Green

# --- Search roots ----------------------------------------------------------
$searchPaths = [System.Collections.Generic.List[string]]::new()

# System-wide installs
$searchPaths.Add("$env:ProgramFiles\Zoom")
$searchPaths.Add("${env:ProgramFiles(x86)}\Zoom")

# Per-user installs — enumerate every profile under C:\Users
$profileRoot = "$env:SystemDrive\Users"
if (Test-Path $profileRoot) {
    foreach ($profile in Get-ChildItem -Path $profileRoot -Directory -ErrorAction SilentlyContinue) {
        $searchPaths.Add("$($profile.FullName)\AppData\Roaming\Zoom")
        $searchPaths.Add("$($profile.FullName)\AppData\Local\Zoom")
    }
}

# Also check the currently-running user explicitly (covers cases where
# $env:SystemDrive\Users differs from actual profile location)
$searchPaths.Add("$env:APPDATA\Zoom")
$searchPaths.Add("$env:LOCALAPPDATA\Zoom")

# --- Find Zoom files (exe, msi) in every root --------------------------------
$found = [System.Collections.Generic.List[PSCustomObject]]::new()
$seen  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($root in ($searchPaths | Sort-Object -Unique)) {
    if (-not (Test-Path $root -ErrorAction SilentlyContinue)) { continue }

    # Search for Zoom.exe (case-insensitive)
    $exes = Get-ChildItem -Path $root -Filter 'Zoom.exe' -Recurse -ErrorAction SilentlyContinue
    foreach ($exe in $exes) {
        if (-not $seen.Add($exe.FullName)) { continue }   # skip duplicates

        try {
            $ver = $exe.VersionInfo.FileVersion
        } catch {
            $ver = 'Unknown'
        }

        $found.Add([PSCustomObject]@{
            DeviceName = $deviceName
            Type       = 'Executable'
            FileName   = $exe.Name
            Version    = if ($ver) { $ver } else { 'Unknown' }
            Path       = $exe.FullName
            CreatedTime = $exe.CreationTime
        })
    }

    # Search for Zoom.msi or zoom.msi installers
    $msis = Get-ChildItem -Path $root -Filter '*.msi' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)zoom' }
    foreach ($msi in $msis) {
        if (-not $seen.Add($msi.FullName)) { continue }

        $found.Add([PSCustomObject]@{
            DeviceName = $deviceName
            Type       = 'MSI Installer'
            FileName   = $msi.Name
            Version    = $msi.VersionInfo.FileVersion
            Path       = $msi.FullName
            CreatedTime = $msi.CreationTime
        })
    }
}

# --- Report: file-system hits -----------------------------------------------
if ($found.Count -eq 0) {
    Write-Host "No Zoom installations found." -ForegroundColor Yellow
} else {
    Write-Host "`n>> File System Findings: $($found.Count) Zoom installation(s) found`n" -ForegroundColor Cyan
    $found | Sort-Object Path | Format-Table -AutoSize @(
        'DeviceName',
        'Type',
        'FileName',
        'Version',
        'Path',
        'CreatedTime'
    )
}

# --- Registry uninstall keys ------------------------------------------------
Write-Host ">> Registry Uninstall Keys:`n" -ForegroundColor Cyan

$regRoots = @(
    # Machine-wide (64-bit)
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    # Machine-wide (32-bit on 64-bit OS)
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
    # Current user
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
)

# Also load every other user's hive via HKU if we can
$hkuKeys = @()
try {
    $hkuKeys = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21' -and $_.PSChildName -notmatch '_Classes$' } |
        ForEach-Object { "Registry::HKEY_USERS\$($_.PSChildName)\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" }
} catch {}

# HKU SID entries are the same physical key as HKCU for the current user.
# Deduplicate by normalising each key's leaf name + uninstall exe path.
$allRegRoots = $regRoots + $hkuKeys

$seenKeys  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$regResults = foreach ($root in $allRegRoots) {
    if (-not (Test-Path $root -ErrorAction SilentlyContinue)) { continue }
    Get-ChildItem -Path $root -ErrorAction SilentlyContinue | ForEach-Object {
        $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        if ($props.DisplayName -match 'Zoom') {
            # Dedup key: subkey name + uninstall exe path
            $dedupeId = "$($_.PSChildName)|$($props.UninstallString)"
            if (-not $seenKeys.Add($dedupeId)) { return }

            $uninstallCmd = $props.UninstallString
            # Zoom doesn't populate QuietUninstallString; /silent is the supported flag
            $quietCmd = if ($props.QuietUninstallString) {
                $props.QuietUninstallString
            } elseif ($uninstallCmd) {
                "$uninstallCmd /silent"
            } else { '' }

            [PSCustomObject]@{
                DeviceName     = $deviceName
                DisplayName    = $props.DisplayName
                DisplayVersion = $props.DisplayVersion
                RegistryKey    = $_.PSPath -replace 'Microsoft\.PowerShell\.Core\\Registry::', ''
                UninstallString = $uninstallCmd
                SilentUninstall = $quietCmd
            }
        }
    }
}

if (-not $regResults) {
    Write-Host "No Zoom uninstall keys found in registry." -ForegroundColor Yellow
} else {
    Write-Host "$($regResults | Measure-Object | Select-Object -ExpandProperty Count) registry entry(ies) found:`n" -ForegroundColor Cyan
    $regResults | Sort-Object DisplayName | Format-List
}

Write-Host "`n=== End of Audit Report ===" -ForegroundColor Green
