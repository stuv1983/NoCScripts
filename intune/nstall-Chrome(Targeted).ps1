# Install-Chrome_fixedTarget.ps1

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetVersion
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Chrome64  = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$Chrome32  = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
$Installer = "$env:TEMP\ChromeSetup.exe"

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
    if (Get-Process "chrome" -ErrorAction SilentlyContinue) { exit 0 }

    $target = [version]$TargetVersion
    $installed = Get-SafeVersionFromExe $Chrome64
    if (-not $installed) { $installed = Get-SafeVersionFromExe $Chrome32 }
    if (-not $installed) { exit 0 }

    if ($installed -ge $target) { exit 0 }

    $url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
    Download-File -Url $url -OutFile $Installer

    $p = Start-Process -FilePath $Installer -ArgumentList "/silent /install" -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { exit 1 }

    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    exit 0
}
catch { exit 1 }
