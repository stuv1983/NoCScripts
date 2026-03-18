# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName a.jones@contoso.com

# Check who has FullAccess permissions on the target mailbox
Get-MailboxPermission -Identity "j.smith@contoso.com" | Where-Object { $_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false }

# Check who has SendAs permissions on the target mailbox
Get-RecipientPermission -Identity "j.smith@contoso.com" | Where-Object { $_.AccessRights -eq "SendAs" -and $_.IsInherited -eq $false }

# Check who has SendOnBehalf permissions on the target mailbox
Get-Mailbox -Identity "j.smith@contoso.com" | Select-Object -ExpandProperty GrantSendOnBehalfTo

# Disconnect the Exchange Online session when done
Disconnect-ExchangeOnline -Confirm:$false
