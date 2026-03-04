<#
.SYNOPSIS
    Hardened Detection for TeamViewer Eradication via Intune.
.DESCRIPTION
    Scans HKLM registry, background services, and per-user AppData for TeamViewer footprints.
    
    INTUNE DETECTION LOGIC:
    - Exit 0 (Success): TeamViewer IS detected. This tells Intune the app is present, 
      which triggers the Uninstall assignment to execute the removal script.
    - Exit 1 (Fail): TeamViewer is NOT detected. This tells Intune the device is clean 
      (Compliant) and the uninstaller does not need to run.
#>

# --- CONFIGURATION TOGGLES ---
$AllowOnServers = $false    # Set to $true to allow detection on Servers and Domain Controllers

# --- STEP 1: WORKSTATION GUARDRAIL ---
# Intune primarily manages workstations. We skip Servers (3) and Domain Controllers (2) 
# by default to prevent accidental removal from critical support infrastructure.
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

if ((-not $os -or $os.ProductType -ne 1) -and -not $AllowOnServers) { 
    Write-Output "COMPLIANT: Target is not a workstation. Skipping detection."
    exit 1 
} elseif ($os -and $os.ProductType -ne 1 -and $AllowOnServers) {
    Write-Output "Target is a server, but `$AllowOnServers is enabled. Proceeding with detection..."
}

$targetFound = $false

# --- STEP 2: REGISTRY DETECTION (HKLM) ---
# We use an anchored regex '^TeamViewer\b' to ensure we strictly match "TeamViewer" 
# or "TeamViewer Host", while avoiding false positives like "TeamViewerMeeting" or other plugins.
$pattern = "^TeamViewer\b" 
$registryPaths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$installedApps = Get-ItemProperty -Path $registryPaths -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName -match $pattern }

if ($installedApps) {
    Write-Output "NON-COMPLIANT: Found TeamViewer registry entry: $($installedApps.DisplayName)"
    $targetFound = $true
}

# --- STEP 3: SERVICE DETECTION ---
# Catch "Ghost" installations. Sometimes a botched uninstallation removes the registry keys 
# and program files, but leaves the background service running, keeping the vulnerability alive.
if (Get-Service -Name "*TeamViewer*" -ErrorAction SilentlyContinue) {
    Write-Output "NON-COMPLIANT: TeamViewer background service detected."
    $targetFound = $true
}

# --- STEP 4: PER-USER APPDATA DETECTION ---
# TeamViewer offers a "Consumer" install that does not require Admin rights. It bypasses 
# Program Files and HKLM entirely, installing directly into the user's AppData folder. 
# We must scan all user profiles to catch these stealth deployments.
$skipList = @("All Users","Default","Default User","Public","WDAGUtilityAccount","DefaultAppPool","Administrator")
$userProfiles = Get-ChildItem -LiteralPath "C:\Users" -Directory -ErrorAction SilentlyContinue | 
                Where-Object { $skipList -notcontains $_.Name }

foreach ($u in $userProfiles) {
    $checkPaths = @(
        (Join-Path $u.FullName "AppData\Local\TeamViewer\TeamViewer.exe"),
        (Join-Path $u.FullName "AppData\Roaming\TeamViewer\TeamViewer.exe")
    )
    foreach ($cp in $checkPaths) {
        if (Test-Path -LiteralPath $cp) { 
            Write-Output "NON-COMPLIANT: Per-user binary found in $($u.Name)'s AppData."
            $targetFound = $true
            break # Break inner loop for this specific user, but continue scanning other users
        }
    }
}

# --- STEP 5: INTUNE EXIT LOGIC ---
if ($targetFound) { 
    # Exit 0 tells Intune: "Yes, it's here. Run the uninstaller."
    exit 0 
} else { 
    # Exit non-zero tells Intune: "Nothing found. The machine is clean."
    Write-Output "COMPLIANT: TeamViewer is not detected on this system."
    exit 1 
}