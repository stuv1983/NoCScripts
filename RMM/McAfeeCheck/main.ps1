<#
.SYNOPSIS
    Read-only McAfee check for N-able. Answers:
      1. Is McAfee installed?
      2. What are its uninstall keys/commands?
      3. What other AV is installed?

    Distinguishes a real install from a stale Security Center registration
    (registered in WSC but nothing on the machine) - the case that makes
    AVCheck warn while nothing is actually installed.

    Makes NO CHANGES. Never runs an uninstall command.

    Exit codes:
      0 = no McAfee at all
      1 = McAfee installed
      3 = McAfee NOT installed, but stale WSC registration present (AVCheck false positive)
      2 = check failed
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
$mc = 'mcafee|webadvisor'

# --- Real install evidence -------------------------------------------------
$products = @(
    Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                     'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match $mc } |
    Select-Object DisplayName, DisplayVersion, PSChildName, UninstallString
)

try { $mcSvc = @(Get-CimInstance Win32_Service -ErrorAction Stop |
        Where-Object { $_.Name -match '^mfe|^mc-fw|^mc-wps' -or $_.DisplayName -match 'McAfee' }) }
catch { $mcSvc = @() }

try { $mcAppx = @(Get-AppxPackage -AllUsers -Name '*McAfee*' -ErrorAction SilentlyContinue) }
catch { $mcAppx = @() }

# --- Security Center view --------------------------------------------------
$wscOk = $true
try { $wscAv = @(Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop) }
catch { $wscOk = $false; $wscAv = @() }

$wscMcAfee = @($wscAv | Where-Object { $_.displayName -match $mc })
$otherAv   = @($wscAv | Where-Object { $_.displayName -notmatch $mc } | ForEach-Object {
    $active = if ($_.productState -band 0x1000) { 'active' } else { 'inactive' }
    "$($_.displayName) ($active)"
})

# --- Verdict ---------------------------------------------------------------
$installed = ($products.Count -gt 0 -or $mcSvc.Count -gt 0 -or $mcAppx.Count -gt 0)
$staleOnly = (-not $installed -and $wscMcAfee.Count -gt 0)

# --- Output ----------------------------------------------------------------
Write-Output "Device: $env:COMPUTERNAME"
Write-Output ''
Write-Output "McAfee installed: $(if ($installed) { 'YES' } elseif ($staleOnly) { 'NO - but stale Security Center entry remains' } else { 'NO' })"

if ($products.Count -gt 0) {
    Write-Output ''
    Write-Output 'Uninstall keys:'
    foreach ($p in $products) {
        Write-Output "- $($p.DisplayName) v$($p.DisplayVersion)"
        Write-Output "  Key: $($p.PSChildName)"
        Write-Output "  Cmd: $($p.UninstallString)"
    }
}
if ($mcAppx.Count -gt 0) {
    Write-Output ''
    Write-Output 'Store app (no uninstall key - remove via Remove-AppxPackage):'
    $mcAppx | Sort-Object -Property PackageFullName -Unique | ForEach-Object {
        Write-Output "- $($_.Name) v$($_.Version)"
    }
}
if ($installed -and $products.Count -eq 0 -and $mcAppx.Count -eq 0) {
    Write-Output '(services present but no uninstall entries - vendor removal tool needed)'
}

if ($staleOnly) {
    Write-Output ''
    Write-Output 'Stale entries (this is why AVCheck warns - nothing to uninstall):'
    foreach ($w in $wscMcAfee) {
        Write-Output "- $($w.displayName) -> fix by clearing its key under HKLM:\SOFTWARE\Microsoft\Security Center\Provider\Av"
    }
}

Write-Output ''
if ($wscOk) {
    Write-Output "Other AV: $(if ($otherAv) { $otherAv -join ', ' } else { 'none' })"
}
else {
    Write-Output 'Other AV: could not query Security Center'
}

if ($installed) { exit 1 }
elseif ($staleOnly) { exit 3 }
elseif (-not $wscOk) { exit 2 }
else { exit 0 }