<#
.SYNOPSIS
  Intune Win32 Detection Script - TeamViewer Removal

.DESCRIPTION
  Exit 0 = TeamViewer NOT found (compliant)
  Exit 1 = TeamViewer found (non-compliant -> run uninstall)

#>

$paths = @(
  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
  Where-Object { $_.DisplayName -match '^TeamViewer' -and ($_.UninstallString -or $_.QuietUninstallString) }

if ($apps) { exit 1 } else { exit 0 }
