<#
.SYNOPSIS
    Detect-Chrome.ps1
    Intune Win32 detection script (Strict MSI Option B - Update Only Mode)
#>

$TargetVersion = [version]"145.0.7632.110"

# ==============================================================================
# PHASE 1: Gather Evidence (No logic, just data collection)
# ==============================================================================
$perUserDetected = $false
$system32BitDetected = $false
$system64BitPath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$system64BitDetected = Test-Path -LiteralPath $system64BitPath

# 1) Check AppData
$skip = @("All Users","Default","Default User","Public","WDAGUtilityAccount")
$users = Get-ChildItem -LiteralPath "C:\Users" -Directory -ErrorAction SilentlyContinue | Where-Object { $skip -notcontains $_.Name }
foreach ($u in $users) {
    if (Test-Path -LiteralPath (Join-Path -Path $u.FullName -ChildPath "AppData\Local\Google\Chrome\Application\chrome.exe")) {
        $perUserDetected = $true
        break
    }
}

# 2) Check 32-bit
if (Test-Path -LiteralPath "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe") {
    $system32BitDetected = $true
}

# 3) Check Registry
$keys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$chromeExes = Get-ItemProperty -Path $keys -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName -match "^Google Chrome$" -and $_.WindowsInstaller -ne 1 }

$chromeMsi = Get-ItemProperty -Path $keys -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName -match "^Google Chrome$" -and $_.WindowsInstaller -eq 1 }

# ==============================================================================
# PHASE 2: Waterfall Evaluation
# ==============================================================================
$anyChromeFootprint = ($perUserDetected -or $system32BitDetected -or $system64BitDetected -or $chromeExes -or $chromeMsi)

# Gatekeeper: Is Chrome even installed?
if (-not $anyChromeFootprint) {
    Write-Output "Chrome not installed. No action required (Update Only Mode)."
    exit 0
}

# The Waterfall
if ($perUserDetected) {
    Write-Output "NON-COMPLIANT: Per-user Chrome (AppData) detected."
    exit 1
}

if ($system32BitDetected) {
    Write-Output "NON-COMPLIANT: Legacy 32-bit Chrome detected."
    exit 1
}

if ($chromeExes) {
    Write-Output "NON-COMPLIANT: System Chrome EXE registry key detected."
    exit 1
}

if (-not $chromeMsi) {
    Write-Output "NON-COMPLIANT: Enterprise MSI registry registration is missing."
    exit 1
}

# If we survive the waterfall, it must be the 64-bit MSI. Verify version & physical files.
if ($system64BitDetected) {
    $v = (Get-Item -LiteralPath $system64BitPath).VersionInfo.ProductVersion
    if ([version]$v -lt $TargetVersion) {
        Write-Output "NON-COMPLIANT: Enterprise MSI is outdated ($v < $TargetVersion)."
        exit 1
    }
} else {
    Write-Output "NON-COMPLIANT: Ghost State - MSI registry exists but physical files are missing."
    exit 1
}

# Victory
Write-Output "COMPLIANT: Meets Strict Enterprise MSI Chrome standard."
exit 0