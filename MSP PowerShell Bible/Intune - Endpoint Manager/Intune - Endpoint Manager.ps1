<#
.SYNOPSIS
    MSP PowerShell Bible - Intune / Endpoint Manager

.PURPOSE
    Device inventory, compliance, app management, Autopilot, and remote actions.

.REQUIRED MODULE
    Microsoft.Graph
#>

# ============================================================
# CONNECT TO INTUNE / ENDPOINT MANAGER
# ============================================================

Install-Module Microsoft.Graph -Scope CurrentUser -Force
Import-Module Microsoft.Graph

Connect-MgGraph -Scopes `
    "DeviceManagementManagedDevices.ReadWrite.All", `
    "DeviceManagementApps.Read.All", `
    "DeviceManagementConfiguration.Read.All", `
    "DeviceManagementServiceConfig.Read.All"

Select-MgProfile -Name "v1.0"

# Confirm connection
Get-MgContext


# ============================================================
# Get-MgDeviceManagementManagedDevice
# What it does:
#   Lists enrolled Intune managed devices.
# When to use:
#   Device inventory, compliance checks, user/device lookup.
# ============================================================

Get-MgDeviceManagementManagedDevice -All |
    Select-Object DeviceName, UserDisplayName, OperatingSystem, ComplianceState, Id


# ============================================================
# Invoke-MgDeviceManagementManagedDeviceWipe
# What it does:
#   Remotely wipes an Intune managed device.
# When to use:
#   Lost/stolen device or secure decommissioning.
# ============================================================

Invoke-MgDeviceManagementManagedDeviceWipe -ManagedDeviceId "<GUID>" `
    -KeepEnrollmentData:$false `
    -KeepUserData:$false


# ============================================================
# Get-MgDeviceManagementDeviceCompliancePolicy
# What it does:
#   Lists Intune device compliance policies.
# When to use:
#   Compliance policy review and troubleshooting.
# ============================================================

Get-MgDeviceManagementDeviceCompliancePolicy |
    Select-Object DisplayName, Id


# ============================================================
# Get-MgDeviceAppManagementMobileApp
# What it does:
#   Lists managed apps in Intune.
# When to use:
#   App inventory and assignment reviews.
# ============================================================

Get-MgDeviceAppManagementMobileApp |
    Select-Object DisplayName, Publisher, Id


# ============================================================
# Get-MgDeviceManagementWindowsAutopilotDeviceIdentity
# What it does:
#   Lists Windows Autopilot device identities.
# When to use:
#   Autopilot inventory and enrolment checks.
# ============================================================

Get-MgDeviceManagementWindowsAutopilotDeviceIdentity |
    Select-Object DisplayName, SerialNumber, GroupTag, Id


# ============================================================
# Invoke-MgDeviceManagementManagedDeviceRemoteLock
# What it does:
#   Sends a remote lock action to a managed device.
# When to use:
#   Lost device response.
# ============================================================

Invoke-MgDeviceManagementManagedDeviceRemoteLock -ManagedDeviceId "<GUID>"
