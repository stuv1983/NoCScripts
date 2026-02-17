<#
.SYNOPSIS
  Auto-Pilot Detection for VLC Media Player.
#>
$ErrorActionPreference = "SilentlyContinue"

# 1. Get Local Version & Arch
$VLC64 = "$env:ProgramFiles\VideoLAN\VLC\vlc.exe"
$VLC32 = "${env:ProgramFiles(x86)}\VideoLAN\VLC\vlc.exe"

if (Test-Path $VLC64) {
    $LocalPath = $VLC64
    $StatusUrl = "http://update.videolan.org/vlc/status-win-x64"
} elseif (Test-Path $VLC32) {
    $LocalPath = $VLC32
    $StatusUrl = "http://update.videolan.org/vlc/status-win-x86"
} else {
    Write-Output "Not Installed"; exit 0
}

$LocalVer = (Get-Item $LocalPath).VersionInfo.ProductVersion

# 2. Get Online Version
try {
    # VideoLAN returns raw text like "3.0.20"
    $OnlineVer = [version](Invoke-RestMethod -Uri $StatusUrl -Method Get -ErrorAction Stop).Trim()
} catch {
    Write-Output "Offline - Skipping"; exit 0
}

# 3. Compare
$CleanLocal = [version]($LocalVer -split '\s+')[0]

if ($CleanLocal -ge $OnlineVer) {
    Write-Output "Compliant"
    exit 0
} else {
    Write-Output "Update Needed"
    exit 1
}