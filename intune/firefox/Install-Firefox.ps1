<#
.SYNOPSIS
    Install-Firefox.ps1
    Auto-Pilot update-only installer with:
      - Fail-Closed logic
      - Cached fallback
      - ESR tracking (Release available but commented)

.DESCRIPTION
    Install flow:

    1) If Firefox running → exit 0 (no disruption).
    2) If not installed → exit 0 (update-only policy).
    3) Determine required version:
         - Vendor API
         - Cached version fallback
    4) If both unavailable → Fail-Closed (exit 1).
    5) If installed version < required → download and update silently.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Firefox64 = "$env:ProgramFiles\Mozilla Firefox\firefox.exe"
$Firefox32 = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
$Installer = "$env:TEMP\FirefoxSetup.exe"

$RegPath = "HKLM:\Software\Kenstra\Evergreen\Firefox"

# ---------- Shared Helpers ----------

function Get-SafeVersionFromString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $m = [regex]::Match($Value, '(\d+(\.\d+){1,3})')
    if (-not $m.Success) { return $null }
    return [version]$m.Groups[1].Value
}

function Get-SafeVersionFromExe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        return Get-SafeVersionFromString ((Get-Item $Path).VersionInfo.ProductVersion)
    } catch { return $null }
}

function Invoke-JsonGet {
    param([string]$Uri)
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    Invoke-RestMethod -Uri $Uri -TimeoutSec 15 -ErrorAction Stop
}

function Get-CachedVendorVersion {
    try {
        return Get-SafeVersionFromString (
            (Get-ItemProperty -Path $RegPath -Name "LatestVendorVersion" -ErrorAction SilentlyContinue)."LatestVendorVersion"
        )
    } catch { return $null }
}

function Set-CachedVendorVersion {
    param([string]$VersionString)
    New-Item -Path $RegPath -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "LatestVendorVersion" -Value $VersionString -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name "LastVendorCheckUtc" -Value ((Get-Date).ToUniversalTime().ToString("o")) -PropertyType String -Force | Out-Null
}

function Get-LatestFirefoxVersion {
    $uri = "https://product-details.mozilla.org/1.0/firefox_versions.json"
    $data = Invoke-JsonGet -Uri $uri

    # ===== OPTION A — RELEASE (Commented Out) =====
    # return Get-SafeVersionFromString $data.LATEST_FIREFOX_VERSION

    # ===== OPTION B — ESR (ACTIVE) =====
    return Get-SafeVersionFromString $data.FIREFOX_ESR
}

function Download-File {
    param([string]$Url, [string]$OutFile)

    try {
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
            Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
            return
        }
    } catch {}

    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
}

# ===============================
# MAIN INSTALL LOGIC
# ===============================

try {

    # Do not interrupt active users
    if (Get-Process "firefox" -ErrorAction SilentlyContinue) { exit 0 }

    # Determine installed version
    $installed = Get-SafeVersionFromExe $Firefox64
    if (-not $installed) { $installed = Get-SafeVersionFromExe $Firefox32 }

    # Update-only: if not installed, do nothing
    if (-not $installed) { exit 0 }

    $latest = $null

    # Try vendor
    try {
        $latest = Get-LatestFirefoxVersion
        if ($latest) {
            Set-CachedVendorVersion -VersionString $latest.ToString()
        }
    } catch {
        $latest = $null
    }

    # Fallback to cache
    if (-not $latest) {
        $latest = Get-CachedVendorVersion
    }

    # Fail-Closed if nothing available
    if (-not $latest) { exit 1 }

    # Already compliant
    if ($installed -ge $latest) { exit 0 }

    # Choose ESR installer architecture
    $url = "https://download.mozilla.org/?product=firefox-esr-latest&os=win64&lang=en-US"
    if (Test-Path $Firefox32 -and -not (Test-Path $Firefox64)) {
        $url = "https://download.mozilla.org/?product=firefox-esr-latest&os=win&lang=en-US"
    }

    Download-File -Url $url -OutFile $Installer

    # Validate download size (>1MB sanity check)
    if ((Get-Item $Installer).Length -lt 1048576) {
        throw "Installer download invalid."
    }

    # Silent install
    $p = Start-Process -FilePath $Installer -ArgumentList "/S /MaintenanceService=true" `
        -Wait -PassThru -WindowStyle Hidden

    # Accept 0 (success) and 3010 (reboot required)
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) {
        throw "Installer failed."
    }

    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    exit 0
}
catch {
    # Fail-Closed on unexpected errors
    exit 1
}
