<#
.SYNOPSIS
    MSP PowerShell Bible - Azure Core Services

.PURPOSE
    Resource groups, virtual machines, networking, RBAC, and automation basics.

.REQUIRED MODULE
    Az
#>

# ============================================================
# CONNECT TO AZURE
# ============================================================

Install-Module Az -Scope CurrentUser -Force
Import-Module Az

Connect-AzAccount

# Optional: set subscription context
Set-AzContext -Subscription "<SUBSCRIPTION-ID-OR-NAME>"


# ============================================================
# Get-AzResourceGroup
# What it does:
#   Lists Azure resource groups.
# When to use:
#   Subscription inventory and environment review.
# ============================================================

Get-AzResourceGroup


# ============================================================
# New-AzResourceGroup
# What it does:
#   Creates a resource group.
# When to use:
#   New project/environment provisioning.
# ============================================================

New-AzResourceGroup -Name "RG-IT" -Location "AustraliaEast"


# ============================================================
# Get-AzVM
# What it does:
#   Lists Azure virtual machines.
# When to use:
#   VM inventory, support, and cost review.
# ============================================================

Get-AzVM |
    Select-Object Name, ResourceGroupName, Location


# ============================================================
# Start-AzVM
# What it does:
#   Starts an Azure VM.
# When to use:
#   Bring a stopped VM online.
# ============================================================

Start-AzVM -Name "VM1" -ResourceGroupName "RG-Prod"


# ============================================================
# Stop-AzVM
# What it does:
#   Stops/deallocates an Azure VM.
# When to use:
#   Maintenance, cost saving, or incident containment.
# ============================================================

Stop-AzVM -Name "VM1" -ResourceGroupName "RG-Prod" -Force


# ============================================================
# Get-AzVirtualNetwork
# What it does:
#   Lists virtual networks.
# When to use:
#   Network inventory and troubleshooting.
# ============================================================

Get-AzVirtualNetwork


# ============================================================
# Get-AzRoleAssignment
# What it does:
#   Lists Azure RBAC assignments.
# When to use:
#   Permission audits and access troubleshooting.
# ============================================================

Get-AzRoleAssignment -SignInName "admin@contoso.com"

# ============================================================
# New-AzRoleAssignment
# What it does:
#   Assigns an Azure RBAC role to a user or group.
# When to use:
#   Granting access to resources.
# ============================================================
New-AzRoleAssignment -SignInName "user@contoso.com" -RoleDefinitionName "Contributor" -ResourceGroupName "RG-IT"

# ============================================================
# Get-AzAutomationAccount
# What it does:
#   Lists Azure Automation accounts.
# When to use:
#   Automation inventory and management.
# ============================================================
Get-AzAutomationAccount

# ============================================================
# New-AzAutomationAccount
# What it does:
#   Creates an Azure Automation account.
# When to use:
#   Setting up automation for a project or environment.
# ============================================================
New-AzAutomationAccount -Name "Auto-IT" -ResourceGroupName "RG-IT" -Location "AustraliaEast"