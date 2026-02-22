<#
.SYNOPSIS
  Update-Firefox.ps1 - Silently updates Mozilla Firefox (x64 or x86) using Mozilla's official installer redirect.

.DESCRIPTION
  PowerShell 5.1 compatible (hardened).
  Fixes:
    - Correct detection for x86 installs and per-user installs
    - Works when run from 32-bit PowerShell
    - Avoids Start-Process -ArgumentList validation edge-cases by always passing a string[]

  Behaviour:
    - Detect Firefox via: common paths + HKLM/HKCU uninstall registry (both 64/32 views)
    - Download latest installer from https://download.mozilla.org/
    - Validate Authenticode signature (Mozilla Corporation)
    - Run silent upgrade (/S).
    - Does NOT force-close Firefox unless -ForceClose.

.PARAMETER ForceClose
  Stops firefox.exe before update attempt.

.PARAMETER Language
  Download language (default en-US)

.PARAMETER Product
  firefox-latest-ssl (default) or firefox-esr-latest-ssl

.PARAMETER TimeoutSeconds
  Installer timeout (default 1800 seconds)
#>

[CmdletBinding()]
param(
  [switch]$ForceClose,
  [string]$Language = "en-US",
  [string]$Product  = "firefox-latest-ssl",
  [ValidateRange(60,7200)]
  [int]$TimeoutSeconds = 1800
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-KV {
  param([string]$Key,[string]$Value)
  if ($null -eq $Value) { $Value = "" }
  Write-Output ("{0} : {1}" -f $Key, $Value)
}

function Test-IsElevated {
  try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}

function Get-EnvPathSafe {
  param([string]$Name)
  try {
    $val = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($val)) { return $null }
    return $val
  } catch { return $null }
}

function Get-FirefoxFromCommonPaths {
  $pfW6432 = Get-EnvPathSafe "ProgramW6432"
  $pf      = Get-EnvPathSafe "ProgramFiles"
  $pfX86   = Get-EnvPathSafe "ProgramFiles(x86)"

  $candidates = @()
  if ($pfW6432) { $candidates += (Join-Path $pfW6432 "Mozilla Firefox\firefox.exe") }
  if ($pf)      { $candidates += (Join-Path $pf "Mozilla Firefox\firefox.exe") }
  if ($pfX86)   { $candidates += (Join-Path $pfX86 "Mozilla Firefox\firefox.exe") }

  $candidates += (Join-Path $env:LOCALAPPDATA "Programs\Mozilla Firefox\firefox.exe")
  $candidates += (Join-Path $env:LOCALAPPDATA "Mozilla Firefox\firefox.exe")

  $candidates = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

  foreach ($exe in $candidates) {
    try {
      $v = (Get-Item $exe).VersionInfo.FileVersion
      return @{ Path=$exe; Version=$v; Scope="Path"; RegKey="" }
    } catch {}
  }
  return $null
}

function Get-UninstallEntries {
  $regPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )

  $items = @()
  foreach ($p in $regPaths) {
    try {
      $items += Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match '^Mozilla Firefox($|\s)' }
    } catch {}
  }
  return $items
}

function Get-FirefoxFromRegistry {
  $entries = @(Get-UninstallEntries)
  if ($entries.Count -eq 0) { return $null }

  foreach ($e in $entries) {
    $exe = $null

    if ($e.InstallLocation) {
      $candidate = Join-Path $e.InstallLocation "firefox.exe"
      if (Test-Path $candidate) { $exe = $candidate }
    }

    if (-not $exe -and $e.DisplayIcon) {
      $icon = ($e.DisplayIcon -replace ',\s*\d+\s*$','').Trim('"')
      if (Test-Path $icon) { $exe = $icon }
    }

    $v = $null
    if ($exe -and (Test-Path $exe)) {
      try { $v = (Get-Item $exe).VersionInfo.FileVersion } catch {}
    }
    if (-not $v) { $v = $e.DisplayVersion }

    $scope = if ($e.PSPath -match '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE') { "HKLM" } else { "HKCU" }

    # --- FIX START ---
    # Replaced ternary operator ($var ? a : b) with PS 5.1 compatible if/else
    $finalPath = if ($exe) { $exe } else { "" }
    return @{ Path=$finalPath; Version=$v; Scope=$scope; RegKey=$e.PSChildName }
    # --- FIX END ---
  }

  $e0 = $entries | Select-Object -First 1
  $scope0 = if ($e0.PSPath -match '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE') { "HKLM" } else { "HKCU" }
  return @{ Path=""; Version=$e0.DisplayVersion; Scope=$scope0; RegKey=$e0.PSChildName }
}

function Get-FirefoxInfo {
  $byPath = Get-FirefoxFromCommonPaths
  if ($byPath) { return $byPath }
  $byReg = Get-FirefoxFromRegistry
  if ($byReg) { return $byReg }
  return $null
}

function Get-InstallArch {
  param([string]$ExePath)
  if ([string]::IsNullOrWhiteSpace($ExePath)) { return "" }
  if ($ExePath -match '(?i)Program Files \(x86\)') { return "x86" }
  if ($ExePath -match '(?i)Program Files\\')      { return "x64" }
  return "Unknown"
}

function Stop-FirefoxIfRequested {
  if (-not $ForceClose) { return @() }
  $stopped = @()
  $procs = Get-Process firefox -ErrorAction SilentlyContinue
  foreach ($p in $procs) {
    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; $stopped += $p.Id } catch {}
  }
  return $stopped
}

function Download-File {
  param([string]$Url,[string]$Out)
  try {
    Start-BitsTransfer -Source $Url -Destination $Out -ErrorAction Stop
    return $true
  } catch {
    try {
      Invoke-WebRequest -Uri $Url -OutFile $Out -UseBasicParsing -ErrorAction Stop
      return $true
    } catch { return $false }
  }
}

function Test-MozillaSignature {
  param([string]$Path)
  try {
    $sig = Get-AuthenticodeSignature -FilePath $Path
    return ($sig.Status -eq "Valid" -and $sig.SignerCertificate.Subject -match "Mozilla Corporation")
  } catch { return $false }
}

function Run-Installer {
  param(
    [Parameter(Mandatory)][string]$Exe,
    [Parameter()][string]$Args
  )

  # HARDENING:
  # -ArgumentList in PS 5.1 validates not-null/not-empty.
  # Also, passing a single string can be treated oddly in some hosts; pass string[].
  $argArray = @()
  if (-not [string]::IsNullOrWhiteSpace($Args)) {
    $argArray = @($Args)
  }

  if ($argArray.Count -gt 0) {
    $p = Start-Process -FilePath $Exe -ArgumentList $argArray -PassThru -WindowStyle Hidden
  } else {
    $p = Start-Process -FilePath $Exe -PassThru -WindowStyle Hidden
  }

  $ok = $p.WaitForExit($TimeoutSeconds * 1000)
  if (-not $ok) {
    try { $p.Kill() } catch {}
    return @{ Done=$false; Exit=$null; TimedOut=$true }
  }
  return @{ Done=$true; Exit=$p.ExitCode; TimedOut=$false }
}

# --------------------
# Output header
# --------------------
Write-KV "Script_Name" "Update-Firefox"
Write-KV "Script_Version" "2.6"
Write-KV "Timestamp" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Write-KV "Elevation_Status" ($(if (Test-IsElevated) { "Elevated" } else { "NotElevated" }))

$infoBefore = Get-FirefoxInfo
if (-not $infoBefore) {
  Write-KV "Firefox_Status" "NotInstalled"
  Write-KV "Result" "OK"
  Write-KV "Result_Detail" "Firefox not detected by path or registry."
  exit 0
}

Write-KV "Firefox_Status" "Installed"
Write-KV "Firefox_Detect_Scope" $infoBefore.Scope
Write-KV "Firefox_Detect_RegKey" $infoBefore.RegKey
Write-KV "Firefox_ExePath_Before" $infoBefore.Path
Write-KV "Firefox_Version_Before" $infoBefore.Version

$running = (Get-Process firefox -ErrorAction SilentlyContinue) -ne $null
Write-KV "Firefox_IsRunning" ([bool]$running)
Write-KV "ForceClose_Enabled" ([bool]$ForceClose)

$stopped = Stop-FirefoxIfRequested
Write-KV "Firefox_StoppedPIDs" (($stopped -join ","))

$installArch = Get-InstallArch -ExePath $infoBefore.Path
Write-KV "Firefox_InstallArch" $installArch

$osToken = "win"
if ([Environment]::Is64BitOperatingSystem -and $installArch -ne "x86") { $osToken = "win64" }
Write-KV "Firefox_Download_OsToken" $osToken
Write-KV "Firefox_Download_Product" $Product
Write-KV "Firefox_Download_Language" $Language

$url = "https://download.mozilla.org/?product=$Product&os=$osToken&lang=$Language"
Write-KV "Download_Url" $url

$work = Join-Path $env:TEMP "KCCC_FirefoxUpdate"
New-Item $work -ItemType Directory -Force | Out-Null
$installer = Join-Path $work ("FirefoxInstaller_{0}_{1}.exe" -f $Product, $osToken)
Write-KV "Installer_Path" $installer

if (-not (Download-File -Url $url -Out $installer)) {
  Write-KV "Result" "FAIL"
  Write-KV "Result_Detail" "Download failed"
  exit 1
}

if (-not (Test-MozillaSignature -Path $installer)) {
  Write-KV "Result" "FAIL"
  Write-KV "Result_Detail" "Invalid installer signature (expected Mozilla Corporation)."
  exit 1
}
Write-KV "Installer_Signature_Valid" "True"

$installerArgs = "/S /MaintenanceService=true"
Write-KV "Installer_Args" $installerArgs

$run = Run-Installer -Exe $installer -Args $installerArgs
Write-KV "Installer_TimedOut" ([bool]$run.TimedOut)
Write-KV "Installer_ExitCode" ($run.Exit)

Start-Sleep -Seconds 2
$infoAfter = Get-FirefoxInfo
if ($infoAfter) {
  Write-KV "Firefox_ExePath_After" $infoAfter.Path
  Write-KV "Firefox_Version_After" $infoAfter.Version
}

if ($run.Done -and ($run.Exit -eq 0 -or $run.Exit -eq 3010)) {
  Write-KV "Result" "OK"
  if ($running -and -not $ForceClose) {
    Write-KV "Result_Detail" "Installer ran successfully. If Firefox was open, update may finalise after close/reopen."
  } else {
    Write-KV "Result_Detail" "Installer ran successfully."
  }
  exit 0
}

Write-KV "Result" "WARN"
Write-KV "Result_Detail" "Installer did not report success. Review exit code."
exit 2
