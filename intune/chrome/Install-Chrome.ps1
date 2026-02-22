<#
.SYNOPSIS
  Backs up user data, updates System Chrome using locally packaged MSIs, 
  and safely removes dormant AppData installations to clear CVE alerts.
#>
$TargetVersion = "145.0.7632.110" 

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue" 

# Point to the MSIs located in the exact same folder as this script
$Installer64 = Join-Path $PSScriptRoot "googlechromestandaloneenterprise64.msi"
$Installer32 = Join-Path $PSScriptRoot "googlechromestandaloneenterprise.msi"
$GlobalExitCode = 0

Write-Output "Executing Packaged Chrome patch sequence..."

# ==============================================================================
# PHASE 1: BACKUP USER DATA (Non-Fatal with Explicit Logging)
# ==============================================================================
$BackupRoot = "C:\temp\ChromeBackup"
if (-not (Test-Path $BackupRoot)) { New-Item -Path $BackupRoot -ItemType Directory -Force | Out-Null }

Write-Output "Initiating pre-update backup of Bookmarks, History, and Passwords..."

if (Test-Path "C:\Users") {
    $Profiles = Get-ChildItem -Path "C:\Users" -Directory
    foreach ($profile in $Profiles) {
        $WinUserName = $profile.Name
        $UserDataPath = "$($profile.FullName)\AppData\Local\Google\Chrome\User Data"

        if (Test-Path $UserDataPath) {
            $UserBackupDir = "$BackupRoot\$WinUserName"
            if (-not (Test-Path $UserBackupDir)) { New-Item -Path $UserBackupDir -ItemType Directory -Force | Out-Null }
            
            # Back up Local State
            $LocalState = "$UserDataPath\Local State"
            if (Test-Path $LocalState) {
                try { Copy-Item -Path $LocalState -Destination $UserBackupDir -Force -ErrorAction Stop } catch { }
            }

            # Back up individual Chrome Profiles
            $ChromeProfiles = Get-ChildItem -Path $UserDataPath -Directory | Where-Object { $_.Name -match "^Default$|^Profile \d+$" }
            
            foreach ($cp in $ChromeProfiles) {
                $CPName = $cp.Name
                $CPBackupDir = "$UserBackupDir\$CPName"
                if (-not (Test-Path $CPBackupDir)) { New-Item -Path $CPBackupDir -ItemType Directory -Force | Out-Null }

                $FilesToBackup = @("Bookmarks", "History", "Login Data")
                foreach ($file in $FilesToBackup) {
                    $SourceFile = "$($cp.FullName)\$file"
                    if (Test-Path $SourceFile) {
                        try { Copy-Item -Path $SourceFile -Destination $CPBackupDir -Force -ErrorAction Stop } catch { }
                    }
                }
            }
        }
    }
}
Write-Output "Backup phase complete."

# ==============================================================================
# PHASE 2: SYSTEM MSI INSTALL/UPDATE (From Package)
# ==============================================================================
function Update-Instance($exePath, $arch) {
    if (-not (Test-Path $exePath)) { return }
    $LocalVer = (Get-Item $exePath).VersionInfo.ProductVersion
    $CleanLocal = [version]($LocalVer -split '\s+')[0]
    
    if ($CleanLocal -lt [version]$TargetVersion) {
        Write-Output "Updating System Installation at $exePath..."
        
        $setupExe = if ($arch -eq "win64") { $Installer64 } else { $Installer32 }
        
        # Failsafe: Ensure the MSI was actually packaged and extracted
        if (-not (Test-Path $setupExe)) {
            Write-Error "CRITICAL: $setupExe is missing from the Intune package payload!"
            $script:GlobalExitCode = 1
            return
        }
        
        $p = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$setupExe`" /qn /norestart" -Wait -PassThru
        if ($p.ExitCode -eq 3010) { $script:GlobalExitCode = 3010 } elseif ($p.ExitCode -ne 0) { $script:GlobalExitCode = $p.ExitCode }
    }
}

try {
    # Update existing System installs
    Update-Instance -exePath "$env:ProgramFiles\Google\Chrome\Application\chrome.exe" -arch "win64"
    Update-Instance -exePath "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" -arch "win32"

    # If NO System installs exist, lay down the 64-bit baseline
    if (-not (Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") -and 
        -not (Test-Path "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe")) {
        
        Write-Output "Deploying 64-bit System MSI baseline..."
        
        if (-not (Test-Path $Installer64)) {
            Write-Error "CRITICAL: $Installer64 is missing from the Intune package payload!"
            $GlobalExitCode = 1
        } else {
            $p = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$Installer64`" /qn /norestart" -Wait -PassThru
            if ($p.ExitCode -eq 3010) { $GlobalExitCode = 3010 } elseif ($p.ExitCode -ne 0) { $GlobalExitCode = $p.ExitCode }
        }
    }

# ==============================================================================
# PHASE 3: SAFE CVE CLEANUP (AppData Remediation)
# ==============================================================================
    Write-Output "Initiating Safe CVE Cleanup for AppData installations..."
    
    $RunningChromePaths = Get-Process chrome -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue

    if (Test-Path "C:\Users") {
        $Profiles = Get-ChildItem -Path "C:\Users" -Directory
        foreach ($profile in $Profiles) {
            $AppDataAppFolder = "$($profile.FullName)\AppData\Local\Google\Chrome\Application"
            $AppDataExe = "$AppDataAppFolder\chrome.exe"
            
            if (Test-Path $AppDataExe) {
                $IsRunning = $RunningChromePaths -contains $AppDataExe
                
                if ($IsRunning) {
                    Write-Warning "Skipping CVE Cleanup for $($profile.Name): Browser is actively running from AppData."
                } else {
                    Write-Output "Removing dormant AppData binaries for $($profile.Name) to clear CVEs..."
                    try {
                        Remove-Item -Path $AppDataAppFolder -Recurse -Force -ErrorAction Stop
                    } catch {
                        Write-Warning "Failed to remove $AppDataAppFolder - $($_.Exception.Message)"
                    }
                }
            }
        }
    }
} 
finally {
    # No MSI cleanup block needed anymore. 
    # Intune automatically wipes the C:\Windows\IMECache temp folder when the deployment finishes.
    Write-Output "Execution complete."
}

Write-Output "Exiting script with Intune Code: $GlobalExitCode"
exit $GlobalExitCode