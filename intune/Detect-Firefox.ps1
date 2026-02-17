<#
.SYNOPSIS
    Detect-Firefox.ps1
    Auto-Pilot detection with:
      - Fail-Closed logic
      - Cached vendor fallback
      - ESR tracking (Release available but commented)

.DESCRIPTION
    Detection flow:

    1) Query Mozilla vendor API for latest version.
       - If successful → cache version in registry.
    2) If vendor API fails → use cached version.
    3) If BOTH vendor + cache unavailable → Fail-Closed (exit 1).

    Update-only behaviour:
      - If Firefox is NOT installed → compliant (exit 0).
      - If installed and version < required → non-compliant (exit 1).
#>

# Stop immediately on unexpected errors
$ErrorActionPreference = "Stop"

# Enforce strict PowerShell syntax
Set-StrictMode -Version Latest

# Paths for 64-bit and 32-bit Firefox
$Firefox64 = "$env:ProgramFiles\Mozilla Firefox\firefox.exe"
$Firefox32 = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"

# Registry location for cached vendor version
$RegPath = "HKLM:\Software\Kenstra\Evergreen\Firefox"

# -------------------------------
# Convert string to safe [version]
# -------------------------------
function Get-SafeVersionFromString {
    param([string]$Value)

    # If null or empty → no version
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    # Extract numeric version portion (ignore "esr" suffix)
    $m = [regex]::Match($Value, '(\d+(\.\d+){1,3})')
    if (-not $m.Success) { return $null }

    # Convert safely to [version] type
    return [version]$m.Groups[1].Value
}

# -------------------------------
# Get installed version from EXE
# -------------------------------
function Get-SafeVersionFromExe {
    param([string]$Path)

    # If file missing → not installed
    if (-not (Test-Path $Path)) { return $null }

    try {
        $raw = (Get-Item $Path).VersionInfo.ProductVersion
        return Get-SafeVersionFromString $raw
    }
    catch {
        return $null
    }
}

# -------------------------------
# Invoke REST call safely (TLS 1.2)
# -------------------------------
function Invoke-JsonGet {
    param([string]$Uri)

    try {
        # Force TLS 1.2 (required on older systems)
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}

    Invoke-RestMethod -Uri $Uri -TimeoutSec 15 -ErrorAction Stop
}

# -------------------------------
# Get cached vendor version
# -------------------------------
function Get-CachedVendorVersion {
    try {
        $v = (Get-ItemProperty -Path $RegPath -Name "LatestVendorVersion" -ErrorAction SilentlyContinue)."LatestVendorVersion"
        return Get-SafeVersionFromString $v
    }
    catch { return $null }
}

# -------------------------------
# Cache vendor version to registry
# -------------------------------
function Set-CachedVendorVersion {
    param([string]$VersionString)

    # Ensure registry path exists
    New-Item -Path $RegPath -Force | Out-Null

    # Store latest vendor version
    New-ItemProperty -Path $RegPath -Name "LatestVendorVersion" `
        -Value $VersionString -PropertyType String -Force | Out-Null

    # Store timestamp of last successful check
    New-ItemProperty -Path $RegPath -Name "LastVendorCheckUtc" `
        -Value ((Get-Date).ToUniversalTime().ToString("o")) `
        -PropertyType String -Force | Out-Null
}

# -------------------------------
# Get latest Firefox version
# -------------------------------
function Get-LatestFirefoxVersion {

    $uri = "https://product-details.mozilla.org/1.0/firefox_versions.json"
    $data = Invoke-JsonGet -Uri $uri

    # ===== OPTION A — RELEASE (Commented Out) =====
    # return Get-SafeVersionFromString $data.LATEST_FIREFOX_VERSION

    # ===== OPTION B — ESR (ACTIVE) =====
    return Get-SafeVersionFromString $data.FIREFOX_ESR
}

# ===============================
# MAIN LOGIC
# ===============================

try {

    # Determine installed version (64-bit preferred)
    $installed = Get-SafeVersionFromExe $Firefox64
    if (-not $installed) { $installed = Get-SafeVersionFromExe $Firefox32 }

    # Update-only: Firefox missing = compliant
    if (-not $installed) { exit 0 }

    $latest = $null

    # Attempt live vendor query
    try {
        $latest = Get-LatestFirefoxVersion
        if ($latest) {
            # Cache successful result
            Set-CachedVendorVersion -VersionString $latest.ToString()
        }
    }
    catch {
        $latest = $null
    }

    # If vendor failed, use cache
    if (-not $latest) {
        $latest = Get-CachedVendorVersion
    }

    # Fail-Closed: if no vendor AND no cache → non-compliant
    if (-not $latest) { exit 1 }

    # Compare versions
    if ($installed -ge $latest) {
        exit 0
    }
    else {
        exit 1
    }
}
catch {
    # Any unexpected failure → Fail-Closed
    exit 1
}
