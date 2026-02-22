<#
.SYNOPSIS
  Production Detection Script for Chrome Enterprise Migration.
  Logic Flow:
  1. If any AppData Chrome exists -> exit 1 (Trigger Migration)
  2. Else if Program Files Chrome exists:
     a. If version < target -> exit 1 (Trigger Update)
     b. If version >= target -> exit 0 (Compliant)
  3. Else (no Chrome anywhere) -> exit 0 (Skip Net-New Install)
#>

$TargetVersion = [version]"145.0.7632.110"
$ErrorActionPreference = "SilentlyContinue"

# ==============================================================================
# 1. If any AppData Chrome exists -> exit 1
# ==============================================================================
if (Test-Path "C:\Users") {
    $PerUserChrome = Get-ChildItem -Path "C:\Users\*\AppData\Local\Google\Chrome\Application\chrome.exe"
    if ($PerUserChrome) {
        Write-Output "Non-Compliant: Consumer Chrome detected in AppData. Migration required."
        exit 1
    }
}

# ==============================================================================
# 2. Else if Program Files Chrome exists
# ==============================================================================
$System64 = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$System32 = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"

$SystemChromeFound = $false
$NeedsUpdate = $false

foreach ($path in @($System64, $System32)) {
    if (Test-Path $path) {
        $SystemChromeFound = $true
        $LocalVer = [version]((Get-Item $path).VersionInfo.ProductVersion -split '\s+')[0]
        
        if ($LocalVer -lt $TargetVersion) {
            Write-Output "Non-Compliant: Enterprise MSI is outdated ($LocalVer < $TargetVersion)."
            $NeedsUpdate = $true
        } else {
            Write-Output "Found compliant Enterprise MSI version ($LocalVer)."
        }
    }
}

if ($SystemChromeFound) {
    if ($NeedsUpdate) {
        exit 1 # version < target
    } else {
        exit 0 # version >= target
    }
}

# ==============================================================================
# 3. Else (no Chrome anywhere) -> exit 0 (skip)
# ==============================================================================
Write-Output "Chrome is not installed anywhere. Compliant (Skip Install)."
exit 0