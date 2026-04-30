<#
.SYNOPSIS
    MSP PowerShell Bible - Exchange On-Prem / Hybrid

.PURPOSE
    On-prem mailbox databases, transport services, hybrid migrations, and mail queues.

.REQUIRED ENVIRONMENT
    Exchange Management Shell, or a remote Exchange PowerShell session.
#>

# ============================================================
# CONNECT TO EXCHANGE ON-PREM / HYBRID
# ============================================================

# Option 1:
# Run this file from the Exchange Management Shell on an Exchange server.

# Option 2:
# Remote PowerShell session to on-prem Exchange.
$ExchangeServer = "exchange01.domain.local"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri "http://$ExchangeServer/PowerShell/" `
    -Authentication Kerberos

Import-PSSession $Session -DisableNameChecking


# ============================================================
# Get-MailboxDatabase
# What it does:
#   Lists mailbox databases and mount status.
# When to use:
#   Exchange health checks and database status reviews.
# ============================================================

Get-MailboxDatabase |
    Select-Object Name, Server, Mounted


# ============================================================
# Get-TransportService
# What it does:
#   Checks Exchange transport service configuration/status.
# When to use:
#   Mail flow troubleshooting.
# ============================================================

Get-TransportService |
    Select-Object Name, Server, Status


# ============================================================
# Test-MAPIConnectivity
# What it does:
#   Tests MAPI connectivity to a mailbox database/server.
# When to use:
#   Outlook/MAPI connectivity troubleshooting.
# ============================================================

Test-MAPIConnectivity -Identity "server.domain.local"


# ============================================================
# New-MoveRequest
# What it does:
#   Starts a mailbox move request.
# When to use:
#   Hybrid migration or mailbox database moves.
# ============================================================

New-MoveRequest -Identity "user@contoso.com" -Remote


# ============================================================
# Get-Queue
# What it does:
#   Lists Exchange transport queues.
# When to use:
#   Investigate stuck mail or transport backlog.
# ============================================================

Get-Queue |
    Select-Object Identity, MessageCount, NextHopDomain, Status


# ============================================================
# DISCONNECT REMOTE SESSION
# ============================================================

Remove-PSSession $Session
