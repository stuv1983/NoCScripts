<#
.SYNOPSIS
  Update-only Intune Win32 app scripts for enforcing a minimum secure version.

.DESCRIPTION
  These scripts are designed for "update-only" deployments:
    - If the app is NOT installed, the device is considered compliant (no install is performed).
    - If installed and version is below TargetVersion, the install script downloads and performs a silent in-place update.

  Important Intune behaviour:
    - Intune runs the Install command ONLY when the Detection script returns a non-zero exit code.
    - Therefore, whenever you change TargetVersion in the Install command, you MUST also update TargetVersion in the Detection script
      (or centralise the version in a shared config source).

.NOTES
  Intended execution context: SYSTEM (Intune Win32 app), 64-bit PowerShell host.
  Exit codes:
    - Detection scripts: 0 = compliant, 1 = non-compliant
    - Install scripts: 0 = success / no action, 1 = failure

  "App running" handling:
    - If the app is running, install scripts exit 0 (skip). This avoids noisy Intune "failed" installs.
      Detection will remain non-compliant until the app closes, at which point the update will apply on a later run.

#>

# =============================
# Install-VLC.ps1 (Update-only)
# =============================
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$TargetVersion
)

$ErrorActionPreference = "Stop"

$VLC64 = "$env:ProgramFiles\VideoLAN\VLC\vlc.exe"
$VLC32 = "${env:ProgramFiles(x86)}\VideoLAN\VLC\vlc.exe"
$Installer = "$env:TEMP\VLCSetup.exe"

# -----------------------------
# Helper: Safe version parsing
# -----------------------------
function Get-SafeVersionFromExe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) { return $null }

    try {
        $raw = (Get-Item $Path).VersionInfo.ProductVersion
        if (-not $raw) { return $null }

        # Common formats:
        #   "122.0.1"
        #   "122.0.1 (build 20260201)"
        #   "122.0.1.0"
        $firstToken = ($raw -split '\s+')[0]

        # Extract numeric x.y(.z)(.w) portion
        $m = [regex]::Match($firstToken, '(\d+(\.\d+){1,3})')
        if (-not $m.Success) { return $null }

        return [version]$m.Groups[1].Value
    }
    catch {
        return $null
    }
}

# -----------------------------
# Helper: Download with BITS fallback
# -----------------------------
function Download-File {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$OutFile
    )

    # Prefer BITS (generally more proxy/TLS friendly)
    try {
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
            Start-BitsTransfer -Source $Url -Destination $OutFile -ErrorAction Stop
            return
        }
    } catch { }

    # Fallback: Invoke-WebRequest
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch { }

    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
}

# -----------------------------
# Helper: Basic download sanity check
# -----------------------------
function Assert-ValidDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [int64]$MinBytes = 1048576 # 1MB default
    )

    if (-not (Test-Path $FilePath)) { throw "Installer download missing: $FilePath" }

    $len = (Get-Item $FilePath).Length
    if ($len -lt $MinBytes) { throw "Download verification failed (file too small: $len bytes)." }
}

try {
    if (Get-Process "vlc" -ErrorAction SilentlyContinue) {
        Write-Output "VLC is running. Skipping update for now."
        exit 0
    }

    $target = [version]$TargetVersion
    $v64 = Get-SafeVersionFromExe -Path $VLC64
    $v32 = Get-SafeVersionFromExe -Path $VLC32

    if (-not $v64 -and -not $v32) {
        Write-Output "VLC not installed. (Update-only) No action."
        exit 0
    }

    if (($v64 -and $v64 -ge $target) -or ($v32 -and $v32 -ge $target)) {
        Write-Output "VLC meets minimum version ($target). No action."
        exit 0
    }

    # VideoLAN "last" endpoints track latest stable for each architecture
    if ($v32) {
        $Url = "https://get.videolan.org/vlc/last/win32/vlc-win32.exe"
        Write-Output "Downloading VLC (32-bit) update..."
    } else {
        $Url = "https://get.videolan.org/vlc/last/win64/vlc-win64.exe"
        Write-Output "Downloading VLC (64-bit) update..."
    }

    Download-File -Url $Url -OutFile $Installer
    Assert-ValidDownload -FilePath $Installer -MinBytes 1048576

    $p = Start-Process -FilePath $Installer -ArgumentList "/S" -Wait -PassThru
    $code = $p.ExitCode

    if ($code -ne 0 -and $code -ne 3010) {
        throw "VLC installer exit code: $code"
    }

    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    Write-Output "VLC updated successfully (exit code $code)."
    exit 0
}
catch {
    Write-Error "VLC update failed: $($_.Exception.Message)"
    exit 1
}
