<#
.SYNOPSIS
    MSP PowerShell Bible - Microsoft Teams

.PURPOSE
    Teams lifecycle, membership, policy assignment, and governance.

.REQUIRED MODULE
    MicrosoftTeams
#>

# ============================================================
# CONNECT TO MICROSOFT TEAMS
# ============================================================

Install-Module MicrosoftTeams -Scope CurrentUser -Force
Import-Module MicrosoftTeams

Connect-MicrosoftTeams


# ============================================================
# Get-Team
# What it does:
#   Lists Teams in the tenant.
# When to use:
#   Teams inventory, ownership and visibility review.
# ============================================================

Get-Team |
    Select-Object DisplayName, Visibility, GroupId


# ============================================================
# New-Team
# What it does:
#   Creates a new Microsoft Team.
# When to use:
#   Team provisioning.
# ============================================================

New-Team -DisplayName "IT Team" -Visibility Private


# ============================================================
# Add-TeamUser
# What it does:
#   Adds a member or owner to a Team.
# When to use:
#   Access requests, onboarding, owner remediation.
# ============================================================

Add-TeamUser -GroupId "<GUID>" -User "john@contoso.com" -Role Member


# ============================================================
# Remove-TeamUser
# What it does:
#   Removes a user from a Team.
# When to use:
#   Offboarding or access removal.
# ============================================================

Remove-TeamUser -GroupId "<GUID>" -User "john@contoso.com"


# ============================================================
# Get-CsTeamsMessagingPolicy
# What it does:
#   Lists Teams messaging policies.
# When to use:
#   Policy reviews and governance checks.
# ============================================================

Get-CsTeamsMessagingPolicy


# ============================================================
# Grant-CsTeamsMessagingPolicy
# What it does:
#   Assigns a messaging policy to a user.
# When to use:
#   Restrict or customise user messaging behaviour.
# ============================================================

Grant-CsTeamsMessagingPolicy -Identity "john@contoso.com" -PolicyName "RestrictedMessaging"


# ============================================================
# Get-CsOnlineUser
# What it does:
#   Shows Teams-enabled user settings including voice/hybrid attributes.
# When to use:
#   Teams voice troubleshooting and user policy validation.
# ============================================================

Get-CsOnlineUser -Identity "john@contoso.com" |
    Select-Object DisplayName, TeamsCallingPolicy, OnlineVoiceRoutingPolicy

# ============================================================
# Set-CsOnlineUser
# What it does:
#   Modifies Teams-related user settings, such as enabling/disabling Teams or assigning voice policies.
# When to use:
#   User provisioning, troubleshooting, or policy enforcement.
# ============================================================
Set-CsOnlineUser -Identity "john@contoso.com" -TeamsCallingPolicy "MyTeamsCallingPolicy"

# ============================================================
# Get-TeamChannel
# What it does:
#   Lists channels within a specific Team.
# When to use:
#   Channel inventory, ownership, and visibility review.
# ============================================================
Get-TeamChannel -GroupId "<GUID>" |
    Select-Object DisplayName, MembershipType, Id

# ============================================================
# New-TeamChannel
# What it does:
#   Creates a new channel within a Team.
# When to use:
#   Channel provisioning for specific projects or topics.
# ============================================================
New-TeamChannel -GroupId "<GUID>" -DisplayName "Project Alpha" -MembershipType Standard 

# ============================================================
# Remove-TeamChannel
# What it does:
#   Deletes a channel from a Team.
# When to use:
#   Channel cleanup or restructuring.
# ============================================================
Remove-TeamChannel -GroupId "<GUID>" -Id "<ChannelId>"

# ============================================================
# Get-TeamUser
# What it does:
#   Lists members and owners of a Team.
# When to use:
#   Membership reviews and audits.
# ============================================================
Get-TeamUser -GroupId "<GUID>" |
    Select-Object User, Role

# ============================================================
# Set-Team
# What it does:
#   Modifies Team properties such as display name, description, or visibility.
# When to use:
#   Team updates or rebranding.
# ============================================================
Set-Team -GroupId "<GUID>" -DisplayName "New IT Team Name" -Description "Updated description for IT Team" -Visibility Private   

# ============================================================
# Get-TeamFunSettings
# What it does:
#   Retrieves the fun settings for a Team, such as allowing Giphy, stickers, and        memes.
# When to use:
#   Reviewing or auditing Team fun settings for compliance.
# ============================================================
Get-TeamFunSettings -GroupId "<GUID>"

# ============================================================  
# Set-TeamFunSettings
# What it does:
#   Modifies the fun settings for a Team, such as enabling or disabling Giphy, stickers, and memes.
# When to use:
#   Enforcing fun settings for compliance or user experience.
# ============================================================
Set-TeamFunSettings -GroupId "<GUID>" -AllowGiphy $false -AllowStickersAndMemes $false -AllowCustomMemes $false

# ============================================================
# Get-TeamGuestSettings
# What it does:
#   Retrieves the guest settings for a Team, such as allowing guests to create or delete channels.
# When to use:
#   Reviewing or auditing Team guest settings for compliance.
# ============================================================
Get-TeamGuestSettings -GroupId "<GUID>"

# ============================================================
# Set-TeamGuestSettings
# What it does:
#   Modifies the guest settings for a Team, such as allowing or disallowing guests to create or delete channels.
# When to use:
#   Enforcing guest settings for compliance or user experience.
# ============================================================
Set-TeamGuestSettings -GroupId "<GUID>" -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false  

# ============================================================
# Get-TeamMemberSettings
# What it does:
#   Retrieves the member settings for a Team, such as allowing members to create or delete channels.
# When to use:
#   Reviewing or auditing Team member settings for compliance.
# ============================================================
Get-TeamMemberSettings -GroupId "<GUID>"

# ============================================================
# Set-TeamMemberSettings
# What it does:
#   Modifies the member settings for a Team, such as allowing or disallowing members to create or delete channels.
# When to use:
#   Enforcing member settings for compliance or user experience.
# ============================================================
Set-TeamMemberSettings -GroupId "<GUID>" -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false

# ============================================================
# Get-TeamMessagingSettings
# What it does:
#   Retrieves the messaging settings for a Team, such as allowing users to edit or delete messages          and allowing priority notifications.
# When to use:
#   Reviewing or auditing Team messaging settings for compliance.
# ============================================================
Get-TeamMessagingSettings -GroupId "<GUID>"

# ============================================================
# Set-TeamMessagingSettings
# What it does:
#   Modifies the messaging settings for a Team, such as allowing or disallowing users       to edit or delete messages and allowing priority notifications.
# When to use:
#   Enforcing messaging settings for compliance or user experience.
# ============================================================
Set-TeamMessagingSettings -GroupId "<GUID>" -AllowUserEditMessages $false -AllowUserDeleteMessages $false -AllowOwnerDeleteMessages $false -AllowTeamMentions $false -AllowChannelMentions $false -AllowPriorityNotifications $false

# ============================================================
# Get-TeamFunSettings
# What it does:
#   Retrieves the fun settings for a Team, such as allowing Giphy, stickers, and memes.
# When to use:
#   Reviewing or auditing Team fun settings for compliance.
# ============================================================
Get-TeamFunSettings -GroupId "<GUID>"

# ============================================================
# Set-TeamFunSettings
# What it does:
#   Modifies the fun settings for a Team, such as enabling or disabling Giphy, stickers, and memes.
# When to use:
#   Enforcing fun settings for compliance or user experience.
# ============================================================
Set-TeamFunSettings -GroupId "<GUID>" -AllowGiphy $false -AllowStickersAndMemes $false -AllowCustomMemes $false

# ============================================================
# Get-TeamArchivedState
# What it does:
#   Retrieves the archived state of a Team.
# When to use:
#   Checking if a Team is archived before performing actions or for inventory purposes.
# ============================================================
Get-TeamArchivedState -GroupId "<GUID>"

# ============================================================
# Set-TeamArchivedState
# What it does:
#   Archives or unarchives a Team.
# When to use:
#   Archiving inactive Teams or unarchiving Teams that need to be reactivated.
# ============================================================
Set-TeamArchivedState -GroupId "<GUID>" -Archived $true

# ============================================================
# Get-TeamVisibility
# What it does:
#   Retrieves the visibility setting of a Team (Public or Private).
# When to use:
#   Checking the visibility of a Team for inventory or compliance purposes.
# ============================================================
Get-TeamVisibility -GroupId "<GUID>"

# ============================================================
# Set-TeamVisibility
# What it does:
#   Changes the visibility of a Team to Public or Private.
# When to use:
#   Modifying the visibility of a Team for compliance or user experience reasons.
# ============================================================
Set-TeamVisibility -GroupId "<GUID>" -Visibility Public

# ============================================================
# Get-TeamArchivedState
# What it does:
#   Retrieves the archived state of a Team.
# When to use:
#   Checking if a Team is archived before performing actions or for inventory purposes.
# ============================================================
Get-TeamArchivedState -GroupId "<GUID>" 

# ============================================================
# Set-TeamArchivedState
# What it does:
#   Archives or unarchives a Team.
# When to use:
#   Archiving inactive Teams or unarchiving Teams that need to be reactivated.
# ============================================================
Set-TeamArchivedState -GroupId "<GUID>" -Archived $true

# ============================================================
# Get-TeamTemplate
# What it does:
#   Lists available Team templates in the tenant.
# When to use:
#   Reviewing available templates for Team provisioning.
# ============================================================
Get-TeamTemplate

# ============================================================
# New-TeamFromTemplate
# What it does:
#   Creates a new Team based on a specified template.
# When to use:
#   Provisioning Teams with predefined settings and channels.
# ============================================================
New-TeamFromTemplate -DisplayName "HR Team" -Template "HRTemplate" -Visibility Private

# ============================================================
# Get-TeamApp
# What it does:
#   Lists apps installed in a Team.
# When to use:
#   Reviewing app usage and compliance within a Team.
# ============================================================
Get-TeamApp -GroupId "<GUID>"   

# ============================================================
# Install-TeamApp
# What it does:
#   Installs an app into a Team.
# When to use:
#   Adding functionality to a Team through apps.
# ============================================================
Install-TeamApp -GroupId "<GUID>" -AppId "<AppId>"

# ============================================================
# Uninstall-TeamApp
# What it does:
#   Removes an app from a Team.
# When to use:
#   Removing unused or non-compliant apps from a Team.
# ============================================================
Uninstall-TeamApp -GroupId "<GUID>" -AppId "<AppId>"

# ============================================================
# Get-TeamAppPermissionPolicy
# What it does:
#   Lists Teams app permission policies in the tenant.
# When to use:
#   Reviewing app permission policies for compliance and governance.
# ============================================================
Get-TeamAppPermissionPolicy

# ============================================================
# Grant-TeamAppPermissionPolicy
# What it does:
#   Assigns an app permission policy to a user.
# When to use:
#   Enforcing app permission policies for specific users.
# ============================================================
Grant-TeamAppPermissionPolicy -Identity "<UserPrincipalName>" -PolicyId "<PolicyId>"

# ============================================================
# Get-TeamAppSetupPolicy
# What it does:
#   Lists Teams app setup policies in the tenant.
# When to use:
#   Reviewing app setup policies for compliance and governance.
# ============================================================
Get-TeamAppSetupPolicy

# ============================================================
# Grant-TeamAppSetupPolicy
# What it does:
#   Assigns an app setup policy to a user.
# When to use:
#   Enforcing app setup policies for specific users.
# ============================================================
Grant-TeamAppSetupPolicy -Identity "<UserPrincipalName>" -PolicyId "<PolicyId>"

# ============================================================
# Get-TeamAppPermissionPolicy
# What it does:
#   Lists Teams app permission policies in the tenant.
# When to use:
#   Reviewing app permission policies for compliance and governance.
# ============================================================
Get-TeamAppPermissionPolicy

# ============================================================
# Grant-TeamAppPermissionPolicy
# What it does:
#   Assigns an app permission policy to a user.
# When to use:
#   Enforcing app permission policies for specific users.
# ============================================================
Grant-TeamAppPermissionPolicy -Identity "<UserPrincipalName>" -PolicyId "<PolicyId>"

# ============================================================
# Get-TeamAppSetupPolicy
# What it does:
#   Lists Teams app setup policies in the tenant.
# When to use:
#   Reviewing app setup policies for compliance and governance.
# ============================================================
Get-TeamAppSetupPolicy


