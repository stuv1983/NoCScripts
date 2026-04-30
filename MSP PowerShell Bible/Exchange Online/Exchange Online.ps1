<#
.SYNOPSIS
    MSP PowerShell Bible - Exchange Online

.PURPOSE
    Mailbox provisioning, permissions, distribution groups, and mail flow troubleshooting.

.REQUIRED MODULE
    ExchangeOnlineManagement
#>

# ============================================================
# CONNECT TO EXCHANGE ONLINE
# ============================================================

Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline

# Confirm connection
Get-ConnectionInformation


# ============================================================
# Get-Mailbox
# What it does:
#   Lists or queries Exchange Online mailboxes.
# When to use:
#   Mailbox audits, recipient checks, shared mailbox reviews.
# ============================================================

Get-Mailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails


# ============================================================
# Set-Mailbox
# What it does:
#   Changes mailbox settings such as quotas, forwarding, holds, and policies.
# When to use:
#   Mailbox configuration changes.
# ============================================================

Set-Mailbox -Identity "john@contoso.com" -ProhibitSendQuota 4GB


# ============================================================
# Add-MailboxPermission
# What it does:
#   Grants mailbox permissions such as FullAccess.
# When to use:
#   Shared mailbox access, manager access, delegated mailbox access.
# ============================================================

$mailbox = "shared@contoso.com"
$user    = "manager@contoso.com"

if (-not (Get-MailboxPermission -Identity $mailbox | Where-Object { $_.User -like $user -and $_.AccessRights -contains "FullAccess" })) {
    Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false
}


# ============================================================
# Add-RecipientPermission
# What it does:
#   Grants SendAs permission.
# When to use:
#   Shared mailbox or mailbox SendAs delegation.
# ============================================================

Add-RecipientPermission -Identity "shared@contoso.com" `
    -Trustee "manager@contoso.com" `
    -AccessRights SendAs `
    -Confirm:$false


# ============================================================
# New-DistributionGroup
# What it does:
#   Creates a distribution group.
# When to use:
#   Mail distribution list creation.
# ============================================================

New-DistributionGroup -Name "IT Team" -PrimarySmtpAddress "it@contoso.com"


# ============================================================
# Add-DistributionGroupMember
# What it does:
#   Adds a member to a distribution group.
# When to use:
#   User onboarding or group membership updates.
# ============================================================

Add-DistributionGroupMember -Identity "IT Team" -Member "john@contoso.com"


# ============================================================
# Get-MessageTrace
# What it does:
#   Traces message delivery in Exchange Online.
# When to use:
#   Mail delivery troubleshooting, missing email investigation.
# ============================================================

Get-MessageTrace -RecipientAddress "user@contoso.com" `
    -StartDate (Get-Date).AddDays(-2) `
    -EndDate (Get-Date)


# ============================================================
# New-TransportRule
# What it does:
#   Creates a mail flow rule.
# When to use:
#   Mail security, disclaimers, attachment blocking, routing exceptions.
# ============================================================

New-TransportRule -Name "Block EXE" `
    -AttachmentTypeMatchesWords "exe" `
    -RejectMessageReasonText "Executable attachments blocked"


# ============================================================
# ARCHIVE MAILBOX VALIDATION
# NOTE:
# ArchiveStatus may show "None" even when ArchiveState is "Local"
# and AutoExpandingArchiveEnabled is True.
# Always validate using statistics and folder checks.
# ============================================================

$User = "user@contoso.com"


# ============================================================
# Get-Mailbox (Archive State Check)
# What it does:
#   Validates mailbox and archive configuration state.
# When to use:
#   First step in archive troubleshooting.
# ============================================================

Get-Mailbox $User |
    Format-List DisplayName,PrimarySmtpAddress,RecipientTypeDetails,
                ArchiveStatus,ArchiveState,AutoExpandingArchiveEnabled


# ============================================================
# Get-MailboxStatistics (Primary Mailbox)
# What it does:
#   Shows primary mailbox size, item count, and activity.
# When to use:
#   Validate mailbox health and size vs archive usage.
# ============================================================

Get-MailboxStatistics $User |
    Format-List TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime


# ============================================================
# Get-MailboxStatistics -Archive
# What it does:
#   Shows archive mailbox size, item count, and last access.
# When to use:
#   Confirm archive exists, is populated, and active.
# ============================================================

Get-MailboxStatistics $User -Archive |
    Format-List TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime


# ============================================================
# Get-MailboxFolderStatistics -Archive (Targeted Folders)
# What it does:
#   Lists key archive folders such as Inbox and Conflicts.
# When to use:
#   Validate folder-level data for backup or migration issues.
# ============================================================

Get-MailboxFolderStatistics $User -Archive |
    Where-Object {
        $_.Name -like "Inbox*" -or
        $_.Name -like "Conflicts"
    } |
    Select-Object Name,ItemsInFolder,FolderSize


# ============================================================
# Get-MailboxFolderStatistics -Archive (Hierarchy Count)
# What it does:
#   Counts total folders in archive.
# When to use:
#   Confirm full archive hierarchy is accessible.
# ============================================================

Get-MailboxFolderStatistics $User -Archive |
    Measure-Object -Property Name -Count

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder List)
# What it does:
#   Lists all folders in the archive.
# When to use:
#   Validate archive folder structure and content.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Search for Specific Folder)
# What it does:
#   Searches for a specific folder by name in the archive.
# When to use:
#   Validate presence of specific folders like "Sent Items" or "Deleted Items".
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Where-Object { $_.Name -like "*Sent Items*" } |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Last Accessed)
# What it does:
#   Shows last accessed time for archive folders.
# When to use:
#   Validate recent activity in the archive.
# ============================================================  
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,LastAccessTime

# ============================================================
# Get-MailboxFolderStatistics -Archive (Deleted Items Check)
# What it does:
#   Checks for items in the archive's Deleted Items folder.
# When to use:
#   Validate if items are being deleted from the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Where-Object { $_.Name -like "*Deleted Items*" } |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Conflict Folder Check)
# What it does:
#   Checks for items in the archive's Conflicts folder.
# When to use:
#   Validate if there are synchronization conflicts in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Where-Object { $_.Name -like "*Conflicts*" } |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Large Folders)
# What it does:
#   Identifies large folders in the archive.
# When to use:
#   Validate if specific folders are consuming excessive space.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Where-Object { $_.FolderSize -gt 100MB } |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Empty Folders)
# What it does:
#   Identifies empty folders in the archive.
# When to use:
#   Validate if there are unused folders in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Where-Object { $_.ItemsInFolder -eq 0 } |
    Select-Object Name,ItemsInFolder,FolderSize

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size Summary)
# What it does:
#   Provides a summary of folder sizes in the archive.
# When to use:
#   Validate overall space usage in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Group-Object -Property Name |
    Select-Object Name,@{Name="TotalSize";Expression={($_.Group | Measure-Object -Property FolderSize -Sum).Sum}},@{Name="TotalItems";Expression={($_.Group | Measure-Object -Property ItemsInFolder -Sum).Sum}}

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size Trend)
# What it does:
#   Shows folder size trends over time in the archive.
# When to use:
#   Validate if specific folders are growing rapidly in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Type)
# What it does:
#   Categorizes folder sizes by type (e.g., Inbox, Sent Items).
# When to use:
#   Validate if specific types of folders are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Group-Object -Property Name |
    Select-Object Name,@{Name="TotalSize";Expression={($_.Group | Measure-Object -Property FolderSize -Sum).Sum}},@{Name="TotalItems";Expression={($_.Group | Measure-Object -Property ItemsInFolder -Sum).Sum}}

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Last Access)
# What it does:
#   Shows folder sizes based on last access time in the archive.
# When to use:
#   Validate if recently accessed folders are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Item Count)
# What it does:
#   Shows folder sizes based on item count in the archive.
# When to use:
#   Validate if folders with more items are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property ItemsInFolder -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Name)
# What it does:
#   Shows folder sizes based on folder name in the archive.
# When to use:
#   Validate if specific named folders are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property Name

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Size)
# What it does:
#   Shows folder sizes based on size in the archive.
# When to use:
#   Validate if larger folders are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Size and Item Count)
# What it does:
#   Shows folder sizes based on size and item count in the archive.
# When to use:
#   Validate if folders with more items are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property FolderSize,ItemsInFolder -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Last Access and Item Count)
# What it does:
#   Shows folder sizes based on last access time and item count in the archive.
# When to use:
#   Validate if recently accessed folders with more items are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime,ItemsInFolder -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Last Access and Size)
# What it does:
#   Shows folder sizes based on last access time and size in the archive.
# When to use:
#   Validate if recently accessed larger folders are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Name and Item Count)
# What it does:
#   Shows folder sizes based on name and item count in the archive.
# When to use:
#   Validate if specific named folders with more items are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property Name,ItemsInFolder -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Name and Last Access)
# What it does:
#   Shows folder sizes based on name and last access time in the archive.
# When to use:
#   Validate if specific named folders that were recently accessed are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property Name,LastAccessTime -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Name and Size)
# What it does:
#   Shows folder sizes based on name and size in the archive.
# When to use:
#   Validate if specific named folders that are larger are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property Name,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Last Access, Item Count, and Size)
# What it does:
#   Shows folder sizes based on last access time, item count, and size in the archive.
# When to use:
#   Validate if recently accessed folders with more items and larger size are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime,ItemsInFolder,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Name, Last Access, Item Count, and Size)
# What it does:
#   Shows folder sizes based on name, last access time, item count, and size in the archive.
# When to use:
#   Validate if specific named folders that were recently accessed with more items and larger size are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property Name,LastAccessTime,ItemsInFolder,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Last Access, Name, Item Count, and Size)
# What it does:
#   Shows folder sizes based on last access time, name, item count, and size in the archive.
# When to use:
#   Validate if recently accessed folders with specific names, more items, and larger size are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property LastAccessTime,Name,ItemsInFolder,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Item Count, Last Access, and Size)
# What it does:
#   Shows folder sizes based on item count, last access time, and size in the archive.
# When to use:
#   Validate if folders with more items that were recently accessed and larger size are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property ItemsInFolder,LastAccessTime,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Item Count, Name, and Size)
# What it does:
#   Shows folder sizes based on item count, name, and size in the archive.
# When to use:
#   Validate if folders with more items that have specific names and larger size are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize |
    Sort-Object -Property ItemsInFolder,Name,FolderSize -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Size, Last Access, and Item Count)
# What it does:
#   Shows folder sizes based on size, last access time, and item count in the archive.
# When to use:
#   Validate if larger folders that were recently accessed with more items are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property FolderSize,LastAccessTime,ItemsInFolder -Descending

# ============================================================
# Get-MailboxFolderStatistics -Archive (Folder Size by Size, Item Count, and Last Access)
# What it does:
#   Shows folder sizes based on size, item count, and last access time in the archive.
# When to use:
#   Validate if larger folders with more items that were recently accessed are consuming more space in the archive.
# ============================================================
Get-MailboxFolderStatistics $User -Archive |
    Select-Object Name,ItemsInFolder,FolderSize,LastAccessTime |
    Sort-Object -Property FolderSize,ItemsInFolder,LastAccessTime -Descending
    