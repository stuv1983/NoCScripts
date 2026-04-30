<#
.SYNOPSIS
    MSP PowerShell Bible - Azure AD / Entra ID (Identity)

.PURPOSE
    User lifecycle, groups, licensing, roles, and authentication.

.REQUIRED MODULE
    Microsoft.Graph
#>

# ============================================================
# CONNECT TO AZURE AD / ENTRA ID
# ============================================================

Install-Module Microsoft.Graph -Scope CurrentUser -Force
Import-Module Microsoft.Graph

Connect-MgGraph -Scopes `
    "User.ReadWrite.All", `
    "Group.ReadWrite.All", `
    "Directory.ReadWrite.All", `
    "Organization.Read.All"

Select-MgProfile -Name "v1.0"

# Confirm connection
Get-MgContext


# ============================================================
# Get-MgUser
# What it does:
#   Retrieves users from Microsoft Entra ID.
# When to use:
#   User audits, account status checks, licence reviews, onboarding/offboarding validation.
# ============================================================

Get-MgUser -All |
    Select-Object DisplayName, UserPrincipalName, AccountEnabled


# ============================================================
# New-MgUser
# What it does:
#   Creates a new Entra ID user.
# When to use:
#   New starter onboarding or test account creation.
# ============================================================

New-MgUser -DisplayName "John Doe" `
    -UserPrincipalName "john@contoso.com" `
    -MailNickname "john" `
    -PasswordProfile @{ Password = "TempP@ss123!"; ForceChangePasswordNextSignIn = $true } `
    -AccountEnabled:$true


# ============================================================
# Update-MgUser
# What it does:
#   Modifies user properties.
# When to use:
#   Disable accounts, update department/title, change user metadata.
# ============================================================

Update-MgUser -UserId "john@contoso.com" -AccountEnabled:$false


# ============================================================
# Remove-MgUser
# What it does:
#   Deletes a user.
# When to use:
#   Final offboarding after retention/export requirements are complete.
# ============================================================

Remove-MgUser -UserId "john@contoso.com" -Confirm:$true


# ============================================================
# Get-MgGroup
# What it does:
#   Lists Entra ID groups and group types.
# When to use:
#   Group audits, membership reviews, security/M365 group identification.
# ============================================================

Get-MgGroup -All |
    Select-Object DisplayName, GroupTypes, MailEnabled, SecurityEnabled, Id


# ============================================================
# New-MgGroup
# What it does:
#   Creates a security or Microsoft 365 group.
# When to use:
#   Access control, application assignment groups, policy targeting groups.
# ============================================================

New-MgGroup -DisplayName "IT Staff" `
    -MailEnabled:$false `
    -SecurityEnabled:$true `
    -MailNickname "ITStaff"


# ============================================================
# Add-MgGroupMember
# What it does:
#   Adds a user or directory object to a group.
# When to use:
#   Access requests, onboarding, group-based licensing, Intune assignments.
# ============================================================

$groupId = "<GROUP-GUID>"
$userId  = "<USER-GUID>"

$existingMember = Get-MgGroupMember -GroupId $groupId -All |
    Where-Object { $_.Id -eq $userId }

if (-not $existingMember) {
    Add-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId
}


# ============================================================
# Get-MgSubscribedSku
# What it does:
#   Lists tenant licence SKUs.
# When to use:
#   Licence availability checks and SKU ID lookup.
# ============================================================

Get-MgSubscribedSku |
    Select-Object SkuId, SkuPartNumber, ConsumedUnits


# ============================================================
# Set-MgUserLicense
# What it does:
#   Assigns or removes licences from users.
# When to use:
#   User onboarding, role changes, licence remediation.
# ============================================================

$sku = (Get-MgSubscribedSku | Where-Object SkuPartNumber -eq "ENTERPRISEPACK").SkuId
$user = Get-MgUser -UserId "john@contoso.com" -Property AssignedLicenses

if ($user.AssignedLicenses.SkuId -notcontains $sku) {
    Set-MgUserLicense -UserId "john@contoso.com" `
        -AddLicenses @{ SkuId = $sku } `
        -RemoveLicenses @()
}

# ============================================================
# Get-MgRoleAssignment
# What it does:
#   Lists Entra ID role assignments.
# When to use:
#   Permission audits, role reviews, access troubleshooting.
# ============================================================

Get-MgRoleAssignment -All |
    Select-Object PrincipalDisplayName, RoleDefinitionId, Scope

# ============================================================
# New-MgRoleAssignment
# What it does:
#   Assigns an Entra ID role to a user or group.
# When to use:
#   Role-based access control, temporary access, delegated administration.
# ============================================================

$role = Get-MgDirectoryRole -Filter "displayName eq 'User Administrator'"
$principalId = (Get-MgUser -UserId "john@contoso.com").Id
New-MgRoleAssignment -DirectoryRoleId $role.Id -PrincipalId $principalId -Scope "/"

# ============================================================
# Get-MgAuthenticationMethod
# What it does:
#   Lists a user's authentication methods.
# When to use:
#   MFA audits, passwordless adoption reviews, authentication troubleshooting.
# ============================================================  
Get-MgUserAuthenticationMethod -UserId "john@contoso.com"

# ============================================================
# New-MgAuthenticationMethod
# What it does:
#   Adds an authentication method for a user.
# When to use:
#   MFA onboarding, passwordless setup, or test method creation.
# ============================================================  
New-MgUserAuthenticationMethod -UserId "john@contoso.com" -MethodType "PhoneAppOTP"

# ============================================================
# Remove-MgAuthenticationMethod
# What it does:
#   Removes an authentication method from a user.
# When to use:
#   MFA offboarding, passwordless method cleanup, or test method removal.
# ============================================================
Remove-MgUserAuthenticationMethod -UserId "john@contoso.com" -AuthenticationMethodId "<AUTHENTICATION-METHOD-GUID>"

# ============================================================
# Disconnect-MgGraph
# What it does:
#   Disconnects the current Microsoft Graph session.
# When to use:
#   End of script, cleanup, or when switching accounts.
# ============================================================
Disconnect-MgGraph