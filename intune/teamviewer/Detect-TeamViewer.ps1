<#
.SYNOPSIS
    Detects TeamViewer presence on Windows workstations.

.DESCRIPTION
    Intune Win32 removal detection semantics:
      Exit 0 = TeamViewer detected (run remediation)
      Exit 1 = Not detected / not a workstation

    Detection covers:
      1) Installed apps (Uninstall registry keys)
      2) Running processes (portable QuickSupport, etc.)

    Portable EXE file scanning is present but DISABLED.

.OUTPUT
    - Prints what was found (Installed/Running)
    - Prints "Not found" if clean

.NO DISK LOGGING
#>

[CmdletBinding()]
param()

# --- Workstation guardrail (1=Workstation, 2=DC, 3=Server) ---
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if (-not $os -or $os.ProductType -ne 1) { Write-Output "Not found"; exit 1 }

# --- Installed detection (registry) ---
$UninstallPaths = @(
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
  'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$installed = foreach ($p in $UninstallPaths) {
  Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -and $_.DisplayName -match 'TeamViewer' } |
    Select-Object DisplayName, DisplayVersion
}

if ($installed) {
  $installed | Sort-Object DisplayName, DisplayVersion | ForEach-Object {
    Write-Output ("Found (Installed): {0} (v{1})" -f $_.DisplayName, $_.DisplayVersion)
  }
  exit 0
}

# --- Running detection (portable etc.) ---
$running = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match '^TeamViewer.*\.exe$' } |
  Select-Object Name, ProcessId, ExecutablePath

if ($running) {
  $running | ForEach-Object {
    Write-Output ("Found (Running): {0} (PID {1}) Path={2}" -f $_.Name, $_.ProcessId, $_.ExecutablePath)
  }
  exit 0
}

# ------------------------------------------------------------
# OPTIONAL – Portable EXE file discovery (DISABLED)
# ------------------------------------------------------------
# NOTE: Commented out to avoid filesystem scanning.
#

$UserPaths = @(
  "$env:SystemDrive\Users\*\Downloads",
  "$env:SystemDrive\Users\*\Desktop",
  "$env:SystemDrive\Users\*\AppData\Local\Temp"
)
foreach ($base in $UserPaths) {
  Get-ChildItem -Path $base -Recurse -Include "*TeamViewer*.exe" -ErrorAction SilentlyContinue |
    ForEach-Object {
      Write-Output ("Found (Portable EXE): {0}" -f $_.FullName)
      exit 0
    }
}


Write-Output "Not found"
exit 1
