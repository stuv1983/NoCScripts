<#
.SYNOPSIS
  Auto-Pilot Detection for Google Chrome (Hybrid Policy).
  Queries Google's official API for the latest Stable version.
#>
$ErrorActionPreference = "SilentlyContinue"

# 1. GET LOCAL VERSION (Update Only - If missing, Exit 0)
$Chrome64 = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$Chrome32 = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"

$LocalPath = if (Test-Path $Chrome64) { $Chrome64 } elseif (Test-Path $Chrome32) { $Chrome32 } else { $null }
if (-not $LocalPath) { Write-Output "Not Installed - Compliant"; exit 0 }

$LocalVer = (Get-Item $LocalPath).VersionInfo.ProductVersion
# Clean version string (remove build metadata if present)
$CleanLocal = [version]($LocalVer -split '\s+')[0]

# 2. GET ONLINE VERSION (Google API)
try {
    # Official endpoint for latest stable versions
    $Uri = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json"
    $Response = Invoke-RestMethod -Uri $Uri -Method Get -ErrorAction Stop
    $OnlineVer = [version]$Response.channels.Stable.version
} catch {
    Write-Warning "API Unreachable. Assuming Compliant to prevent errors."
    exit 0
}

# 3. COMPARE
if ($CleanLocal -ge $OnlineVer) {
    Write-Output "Compliant ($CleanLocal >= $OnlineVer)"
    exit 0
} else {
    Write-Output "Update Needed ($CleanLocal < $OnlineVer)"
    exit 1
}