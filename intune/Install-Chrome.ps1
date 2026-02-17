<#
.SYNOPSIS
  Install-Chrome.ps1 (Auto-Pilot, Fail-Closed with Cached Fallback)

.DESCRIPTION
  Update-only installer for Google Chrome that pairs with Detect-Chrome.ps1.

  CURRENT MODE:
    ✔ Stable channel — ACTIVE
    ✖ Extended Stable — COMMENTED OUT

  Install flow:
    1) If Chrome is running -> exit 0 (avoid disrupting user; Intune retries later).
    2) If Chrome is NOT installed -> exit 0 (update-only policy).
    3) Determine required version (vendor -> cache -> fail-closed).
    4) If installed version < required -> download latest installer and update silently.

  Vendor API:
    https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/<channel>/versions
    Channel identifiers include 'stable' and 'extended'. citeturn0search1

  Registry cache:
    HKLM:\Software\Kenstra\Evergreen\Chrome
      LatestVendorVersion (REG_SZ)
      LastVendorCheckUtc  (REG_SZ)

#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Chrome64  = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$Chrome32  = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
$Installer = "$env:TEMP\ChromeSetup.exe"

$RegPath = "HKLM:\Software\Kenstra\Evergreen\Chrome"

function Get-SafeVersionFromString {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    # Extract numeric dotted version portion only
    $m = [regex]::Match($Value, '(\d+(\.\d+){1,3})')
    if (-not $m.Success) { return $null }

    try { return [version]$m.Groups[1].Value } catch { return $null }
}

function Get-SafeVersionFromExe {
    param([string]$Path)

    # Missing file => not installed
    if (-not (Test-Path $Path)) { return $null }

    try {
        $raw = (Get-Item $Path).VersionInfo.ProductVersion
        return Get-SafeVersionFromString $raw
    }
    catch { return $null }
}

function Invoke-JsonGet {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [int]$TimeoutSec = 15
    )

    # Force TLS 1.2 for older systems / .NET defaults
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    # VersionHistory API is unauthenticated for these endpoints
    return Invoke-RestMethod -Uri $Uri -Method Get -TimeoutSec $TimeoutSec -ErrorAction Stop
}

function Get-CachedVendorVersion {
    param([Parameter(Mandatory=$true)][string]$RegPath)

    try {
        $v = (Get-ItemProperty -Path $RegPath -Name "LatestVendorVersion" -ErrorAction SilentlyContinue)."LatestVendorVersion"
        return Get-SafeVersionFromString $v
    } catch { return $null }
}

function Set-CachedVendorVersion {
    param(
        [Parameter(Mandatory=$true)][string]$RegPath,
        [Parameter(Mandatory=$true)][string]$VersionString
    )

    # Create path if missing
    New-Item -Path $RegPath -Force | Out-Null

    # Store latest vendor version observed
    New-ItemProperty -Path $RegPath -Name "LatestVendorVersion" -Value $VersionString -PropertyType String -Force | Out-Null

    # Store when we last successfully checked the vendor (UTC)
    New-ItemProperty -Path $RegPath -Name "LastVendorCheckUtc" -Value ((Get-Date).ToUniversalTime().ToString("o")) -PropertyType String -Force | Out-Null
}

function Download-File {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$OutFile
    )

    # Prefer BITS (more reliable through some proxies)
    try {
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
            Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
            return
        }
    } catch { }

    # Fallback to Invoke-WebRequest
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
}

function Assert-DownloadOk {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [int64]$MinBytes = 1048576
    )

    if (-not (Test-Path $FilePath)) { throw "Installer download missing: $FilePath" }
    if ((Get-Item $FilePath).Length -lt $MinBytes) { throw "Installer too small; download likely failed." }
}

function Get-LatestChromeChannelVersion {
    param([Parameter(Mandatory=$true)][string]$Channel)

    $uri = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/$Channel/versions"
    $data = Invoke-JsonGet -Uri $uri
    if (-not $data.versions) { return $null }

    $versions = @()
    foreach ($v in $data.versions) {
        $pv = Get-SafeVersionFromString $v.version
        if ($pv) { $versions += $pv }
    }
    if ($versions.Count -eq 0) { return $null }

    return ($versions | Sort-Object -Descending | Select-Object -First 1)
}

function Get-LatestChromeVersion {
    # ===== OPTION A — STABLE (ACTIVE) =====
    $channel = "stable"

    # ===== OPTION B — EXTENDED STABLE (Commented Out) =====
    # $channel = "extended"

    return Get-LatestChromeChannelVersion -Channel $channel
}

try {
    # Avoid disrupting users
    if (Get-Process "chrome" -ErrorAction SilentlyContinue) { exit 0 }

    $installed = Get-SafeVersionFromExe $Chrome64
    if (-not $installed) { $installed = Get-SafeVersionFromExe $Chrome32 }
    if (-not $installed) { exit 0 }

    $latest = $null

    try {
        $latest = Get-LatestChromeVersion
        if ($latest) {
            Set-CachedVendorVersion -RegPath $RegPath -VersionString $latest.ToString()
        }
    } catch {
        $latest = $null
    }

    if (-not $latest) {
        $latest = Get-CachedVendorVersion -RegPath $RegPath
    }

    # Fail-closed: if we can't determine "latest", fail so Intune retries.
    if (-not $latest) { exit 1 }

    if ($installed -ge $latest) { exit 0 }

    # Download and run Google's latest installer silently
    $url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
    Download-File -Url $url -OutFile $Installer
    Assert-DownloadOk -FilePath $Installer

    $p = Start-Process -FilePath $Installer -ArgumentList "/silent /install" -Wait -PassThru -WindowStyle Hidden
    $code = $p.ExitCode
    if ($code -ne 0 -and $code -ne 3010) { throw "Chrome installer exit code: $code" }

    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    exit 0
}
catch {
    exit 1
}
