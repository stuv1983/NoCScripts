<#
.SYNOPSIS
  Production Detection Script for Firefox Enterprise Migration.
  Logic Flow:
  1. If any AppData Firefox exists -> exit 1 (Trigger Migration)
  2. Else if Program Files Firefox exists:
     a. If version < target -> exit 1 (Trigger Update)
     b. If version >= target -> exit 0 (Compliant)
  3. Else (no Firefox anywhere) -> exit 0 (Skip Net-New Install)
#>

$TargetVersion = [version]"147.0.4"
$ErrorActionPreference = "SilentlyContinue"

# ==============================================================================
# 1. If any AppData Firefox exists -> exit 1
# ==============================================================================
if (Test-Path "C:\Users") {
    $PerUserFirefox = Get-ChildItem -Path "C:\Users\*\AppData\Local\Mozilla Firefox\firefox.exe"
    if ($PerUserFirefox) {
        Write-Output "Non-Compliant: Consumer Firefox detected in AppData. Migration required."
        exit 1
    }
}

# ==============================================================================
# 2. Else if Program Files Firefox exists
# ==============================================================================
$System64 = "$env:ProgramFiles\Mozilla Firefox\firefox.exe"
$System32 = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"

$SystemFirefoxFound = $false
$NeedsUpdate = $false

foreach ($path in @($System64, $System32)) {
    if (Test-Path $path) {
        $SystemFirefoxFound = $true
        # Firefox version strings sometimes contain extra info, grabbing just the number
        $LocalVer = [version]((Get-Item $path).VersionInfo.ProductVersion -split '\s+')[0]
        
        if ($LocalVer -lt $TargetVersion) {
            Write-Output "Non-Compliant: Enterprise MSI is outdated ($LocalVer < $TargetVersion)."
            $NeedsUpdate = $true
        } else {
            Write-Output "Found compliant Enterprise MSI version ($LocalVer)."
        }
    }
}

if ($SystemFirefoxFound) {
    if ($NeedsUpdate) {
        exit 1 # version < target
    } else {
        exit 0 # version >= target
    }
}

# ==============================================================================
# 3. Else (no Firefox anywhere) -> exit 0 (skip)
# ==============================================================================
Write-Output "Firefox is not installed anywhere. Compliant (Skip Install)."
exit 0