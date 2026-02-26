<#
.SYNOPSIS
    Detect-Firefox.ps1
    Intune Custom Detection Script for Firefox Enterprise (Strict 64-Bit MSI).
    
.DESCRIPTION
    STRICT WATERFALL LOGIC (Standardise-If-Present Mode):
    0. PASS if Firefox is completely absent (No physical files OR Registry keys).
    1. FAIL if any per-user AppData installation exists (forces remediation).
    2. FAIL if Legacy 32-bit Firefox exists (forces remediation to 64-bit).
    3. PASS only if 64-bit System-level binary exists AND meets/exceeds target version.
    4. FAIL if missing, outdated, or ghosted.
#>

$ErrorActionPreference = "SilentlyContinue"

$TargetVersion = [version]"147.0.4"

# ==============================================================================
# PHASE 0: GATEKEEPER (IS IT EVEN INSTALLED?)
# ==============================================================================
$ffInstalled = $false
$ffSystemPaths = @(
    "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
    "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
)

# 1. Quick check for physical System footprints
if ((Test-Path -LiteralPath $ffSystemPaths[0]) -or (Test-Path -LiteralPath $ffSystemPaths[1])) { 
    $ffInstalled = $true 
}

# 2. Quick check for Registry footprints (Uninstall Keys & Native Mozilla Hives)
if (-not $ffInstalled) {
    # Check standard Uninstall keys (tightened regex to avoid unrelated Mozilla products)
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $ffReg = Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -match "(?i)^Mozilla Firefox" }
             
    # Check native vendor hives (catches deeply corrupted/half-uninstalled states)
    $mozillaHives = @(
        "HKLM:\SOFTWARE\Mozilla\Mozilla Firefox",
        "HKLM:\SOFTWARE\WOW6432Node\Mozilla\Mozilla Firefox"
    )
    $ffHive = Get-Item -Path $mozillaHives -ErrorAction SilentlyContinue

    if ($ffReg -or $ffHive) { $ffInstalled = $true }
}

# 3. Quick check for AppData footprints
if (-not $ffInstalled) {
    $users = Get-ChildItem -LiteralPath "C:\Users" -Directory -Force -ErrorAction SilentlyContinue
    foreach ($u in $users) {
        if ($u.Name -match "^(All Users|Default|Default User|Public|WDAGUtilityAccount|Administrator)$") { continue }
        
        $roguePaths = @(
            "$($u.FullName)\AppData\Local\Mozilla Firefox\firefox.exe",
            "$($u.FullName)\AppData\Local\Mozilla\Firefox\firefox.exe",
            "$($u.FullName)\AppData\Local\Firefox\firefox.exe"
        )
        
        foreach ($exe in $roguePaths) {
            if (Test-Path -LiteralPath $exe) {
                $ffInstalled = $true
                break
            }
        }
        if ($ffInstalled) { break }
    }
}

if (-not $ffInstalled) {
    Write-Output "Compliant: Firefox not installed. No action required (Standardise-If-Present Mode)."
    exit 0
}

# ==============================================================================
# PHASE 1: ROGUE / PER-USER INSTALLATION CHECK (IMMEDIATE FAIL)
# ==============================================================================
$users = Get-ChildItem -LiteralPath "C:\Users" -Directory -Force

foreach ($u in $users) {
    if ($u.Name -match "^(All Users|Default|Default User|Public|WDAGUtilityAccount|Administrator)$") { continue }

    $roguePaths = @(
        "$($u.FullName)\AppData\Local\Mozilla Firefox\firefox.exe",
        "$($u.FullName)\AppData\Local\Mozilla\Firefox\firefox.exe",
        "$($u.FullName)\AppData\Local\Firefox\firefox.exe"
    )

    foreach ($exe in $roguePaths) {
        if (Test-Path -LiteralPath $exe) {
            Write-Warning "Non-compliant: Rogue per-user Firefox found at $exe. Requires remediation."
            exit 1 
        }
    }
}

# ==============================================================================
# PHASE 2: LEGACY 32-BIT CHECK (IMMEDIATE FAIL)
# ==============================================================================
$legacy32BitPath = "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"

if (Test-Path -LiteralPath $legacy32BitPath) {
    Write-Warning "Non-compliant: Legacy 32-bit Firefox found at $legacy32BitPath. Requires remediation."
    exit 1
}

# ==============================================================================
# PHASE 3: 64-BIT SYSTEM-LEVEL COMPLIANCE CHECK (PASS CONDITION)
# ==============================================================================
$system64BitPath = "$env:ProgramFiles\Mozilla Firefox\firefox.exe"

if (Test-Path -LiteralPath $system64BitPath) {
    $fileVersionString = (Get-Item -LiteralPath $system64BitPath).VersionInfo.ProductVersion
    
    if (-not [string]::IsNullOrWhiteSpace($fileVersionString)) {
        $cleanVersion = $fileVersionString -replace '[a-zA-Z\-].*',''
        $installedVersion = [version]$cleanVersion

        if ($installedVersion -ge $TargetVersion) {
            Write-Output "Compliant: 64-bit System Firefox found at $system64BitPath (Version: $installedVersion)"
            exit 0
        } else {
            Write-Warning "Non-compliant: 64-bit System Firefox is outdated ($installedVersion < $TargetVersion)."
            exit 1
        }
    }
}

# ==============================================================================
# PHASE 4: GHOST FOOTPRINT / MISSING (CATCH-ALL FAIL)
# ==============================================================================
Write-Warning "Non-compliant: 64-bit System Firefox binary missing or broken."
exit 1