<#
.SYNOPSIS
    MSP PowerShell Bible - Windows 365 / Azure Virtual Desktop

.PURPOSE
    Cloud PC inventory, host pools, session hosts, workspaces, and lifecycle checks.

.REQUIRED MODULES
    Microsoft.Graph
    Az.DesktopVirtualization
#>

# ============================================================
# CONNECT TO WINDOWS 365 AND AZURE VIRTUAL DESKTOP
# ============================================================

Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module Az.DesktopVirtualization -Scope CurrentUser -Force
Install-Module Az.Accounts -Scope CurrentUser -Force

Import-Module Microsoft.Graph
Import-Module Az.Accounts
Import-Module Az.DesktopVirtualization

Connect-MgGraph -Scopes `
    "CloudPC.Read.All", `
    "DeviceManagementManagedDevices.Read.All"

Connect-AzAccount

# Optional: set subscription context
Set-AzContext -Subscription "<SUBSCRIPTION-ID-OR-NAME>"


# ============================================================
# Get-MgDeviceManagementVirtualEndpointCloudPC
# What it does:
#   Lists Windows 365 Cloud PCs.
# When to use:
#   Cloud PC inventory and assignment review.
# ============================================================

Get-MgDeviceManagementVirtualEndpointCloudPC


# ============================================================
# Get-AzWvdHostPool
# What it does:
#   Lists Azure Virtual Desktop host pools.
# When to use:
#   AVD environment inventory.
# ============================================================

Get-AzWvdHostPool


# ============================================================
# Get-AzWvdSessionHost
# What it does:
#   Lists session hosts in an AVD host pool.
# When to use:
#   Session host troubleshooting, drain mode checks, capacity review.
# ============================================================

Get-AzWvdSessionHost -HostPoolName "<PoolName>" -ResourceGroupName "<RG>"


# ============================================================
# Get-AzWvdWorkspace
# What it does:
#   Lists AVD workspaces.
# When to use:
#   Workspace/application group mapping checks.
# ============================================================

Get-AzWvdWorkspace


# ============================================================
# Stop-AzWvdSessionHost
# What it does:
#   Stops a session host where supported by the installed module/provider.
# When to use:
#   Maintenance or controlled host lifecycle actions.
# ============================================================

Stop-AzWvdSessionHost -ResourceGroupName "<RG>" -HostPoolName "<Pool>" -Name "<HostName>"
