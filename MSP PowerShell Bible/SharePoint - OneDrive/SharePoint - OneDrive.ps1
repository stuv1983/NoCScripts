<#
.SYNOPSIS
    MSP PowerShell Bible - SharePoint / OneDrive

.PURPOSE
    Site lifecycle, permissions, OneDrive ownership, external users, and deleted site recovery.

.REQUIRED MODULE
    Microsoft.Online.SharePoint.PowerShell
#>

# ============================================================
# CONNECT TO SHAREPOINT ONLINE
# ============================================================

Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
Import-Module Microsoft.Online.SharePoint.PowerShell

Connect-SPOService -Url "https://tenant-admin.sharepoint.com"


# ============================================================
# Get-SPOSite
# What it does:
#   Lists SharePoint and OneDrive sites.
# When to use:
#   Site inventory, ownership checks, storage reviews.
# ============================================================

Get-SPOSite -Limit All |
    Select-Object Url, Owner, Template, StorageUsageCurrent


# ============================================================
# New-SPOSite
# What it does:
#   Creates a classic site collection.
# When to use:
#   Legacy site provisioning where classic site collection creation is required.
# ============================================================

New-SPOSite -Url "https://tenant.sharepoint.com/sites/IT" `
    -Owner "admin@contoso.com" `
    -StorageQuota 2048 `
    -Title "IT"


# ============================================================
# Set-SPOSite
# What it does:
#   Modifies site settings, including ownership.
# When to use:
#   OneDrive ownership transfer, site lock/settings changes.
# ============================================================

Set-SPOSite -Identity "https://tenant-my.sharepoint.com/personal/john_contoso_com" `
    -Owner "admin@contoso.com"


# ============================================================
# Add-SPOUser
# What it does:
#   Adds a user to a SharePoint site group.
# When to use:
#   Grant site access without using the SharePoint UI.
# ============================================================

Add-SPOUser -Site "https://tenant.sharepoint.com/sites/IT" `
    -Group "Members" `
    -LoginName "john@contoso.com"


# ============================================================
# Get-SPOExternalUser
# What it does:
#   Lists guest/external users.
# When to use:
#   External sharing reviews and tenant guest audits.
# ============================================================

Get-SPOExternalUser -PageSize 200


# ============================================================
# Remove-SPOExternalUser
# What it does:
#   Removes an external user by UniqueId.
# When to use:
#   External access cleanup.
# ============================================================

Remove-SPOExternalUser -UniqueIds "<GUID>"


# ============================================================
# Get-SPODeletedSite
# What it does:
#   Lists deleted SharePoint/OneDrive sites.
# When to use:
#   Restore investigations after accidental deletion.
# ============================================================

Get-SPODeletedSite


# ============================================================
# Restore-SPODeletedSite
# What it does:
#   Restores a deleted SharePoint or OneDrive site.
# When to use:
#   Recover deleted site collections or OneDrive sites.
# ============================================================

Restore-SPODeletedSite -Identity "https://tenant-my.sharepoint.com/personal/john_contoso_com"
