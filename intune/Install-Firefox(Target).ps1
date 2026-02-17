# Install-Firefox_fixedTarget.ps1

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetVersion
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Firefox64 = "$env:ProgramFiles\Mozilla Firefox\firefox.exe"
$Firefox32 = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
$Installer = "$env:TEMP\FirefoxSetup.exe"

function Get-SafeVersionFromExe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $raw = (Get-Item $Path).VersionInfo.ProductVersion
        $token = ($raw -split '\s+')[0]
        $m = [regex]::Match($token, '(\d+(\.\d+){1,3})')
        if (-not $m.Success) { return $null }
        return [version]$m.Groups[1].Value
    } catch { return $null }
}

function Download-File {
    param([string]$Url,[string]$OutFile)
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
}

try {
    if (Get-Process "firefox" -ErrorAction SilentlyContinue) { exit 0 }

    $target = [version]$TargetVersion
    $installed = Get-SafeVersionFromExe $Firefox64
    if (-not $installed) { $installed = Get-SafeVersionFromExe $Firefox32 }
    if (-not $installed) { exit 0 }

    if ($installed -ge $target) { exit 0 }

    $url = "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US"
    Download-File -Url $url -OutFile $Installer

    $p = Start-Process -FilePath $Installer -ArgumentList "/S /MaintenanceService=true" -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { exit 1 }

    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    exit 0
}
catch { exit 1 }
