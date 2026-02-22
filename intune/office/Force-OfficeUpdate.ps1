<#
.SYNOPSIS
  Forces Office Click-to-Run to check for and stage the latest updates silently.
#>
$ErrorActionPreference = "SilentlyContinue"

$RegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$C2RClient = "$env:ProgramFiles\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"

# If Office isn't installed, exit silently
if (-not (Test-Path $RegPath) -or -not (Test-Path $C2RClient)) {
    Write-Output "Office Click-to-Run is not present. Exiting."
    exit 0
}

Write-Output "Office detected. Triggering background update check..."

# Fire the native update engine asynchronously 
Start-Process -FilePath $C2RClient -ArgumentList "/update user displaylevel=false forceappshutdown=false" -Wait:$false

Write-Output "Update triggered successfully."
exit 0