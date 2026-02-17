# Detect-Chrome_fixedTarget.ps1
# MUST match Intune Install command -TargetVersion value

$TargetVersion = "145.0.7632.76"

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Exe64 = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$Exe32 = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"

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

try {
    $target = [version]$TargetVersion
    $v64 = Get-SafeVersionFromExe $Exe64
    $v32 = Get-SafeVersionFromExe $Exe32

    if (-not $v64 -and -not $v32) { exit 0 }

    if (($v64 -and $v64 -ge $target) -or ($v32 -and $v32 -ge $target)) { exit 0 }

    exit 1
}
catch { exit 1 }
