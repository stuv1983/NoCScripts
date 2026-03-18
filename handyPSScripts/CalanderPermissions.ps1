# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName a.jones@contoso.com

# Check calendar permissions for a specific user's calendar
Get-MailboxFolderPermission -Identity "j.smith@contoso.com:\Calendar"

# Check calendar permissions filtered to a specific user
Get-MailboxFolderPermission -Identity "j.smith@contoso.com:\Calendar" | Where-Object { $_.User -eq "b.taylor@contoso.com" }

# Grant a user Editor permissions on the calendar (can read, create and edit items)
Add-MailboxFolderPermission -Identity "j.smith@contoso.com:\Calendar" -User "b.taylor@contoso.com" -AccessRights Editor

# Update an existing user's calendar permissions
Set-MailboxFolderPermission -Identity "j.smith@contoso.com:\Calendar" -User "b.taylor@contoso.com" -AccessRights Reviewer

# Remove a user's calendar permissions
Remove-MailboxFolderPermission -Identity "j.smith@contoso.com:\Calendar" -User "b.taylor@contoso.com"

# Disconnect the Exchange Online session when done
Disconnect-ExchangeOnline -Confirm:$false
