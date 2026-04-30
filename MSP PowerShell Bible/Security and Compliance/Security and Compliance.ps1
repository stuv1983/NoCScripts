<#
.SYNOPSIS
    MSP PowerShell Bible - Security & Compliance Center

.PURPOSE
    Audit logs, retention policies, content search, eDiscovery, and legal hold workflows.

.REQUIRED MODULE
    ExchangeOnlineManagement
#>

# ============================================================
# CONNECT TO SECURITY & COMPLIANCE / PURVIEW POWERSHELL
# ============================================================

Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Import-Module ExchangeOnlineManagement

Connect-IPPSSession


# ============================================================
# Search-UnifiedAuditLog
# What it does:
#   Queries Microsoft 365 unified audit logs.
# When to use:
#   User activity investigation, file access review, admin activity checks.
# ============================================================

Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-7) `
    -EndDate (Get-Date) `
    -UserIds "john@contoso.com"


# ============================================================
# Get-RetentionCompliancePolicy
# What it does:
#   Lists retention compliance policies.
# When to use:
#   Retention policy review and governance checks.
# ============================================================

Get-RetentionCompliancePolicy


# ============================================================
# New-ComplianceSearch
# What it does:
#   Creates a compliance/eDiscovery content search.
# When to use:
#   Legal, HR, phishing, or incident response searches.
# ============================================================

New-ComplianceSearch -Name "Search1" `
    -ExchangeLocation All `
    -ContentMatchQuery 'from:"ceo@contoso.com"'


# ============================================================
# Start-ComplianceSearch
# What it does:
#   Starts a created compliance search.
# When to use:
#   Run an eDiscovery/content search after creating it.
# ============================================================

Start-ComplianceSearch -Identity "Search1"


# ============================================================
# Get-ComplianceCase
# What it does:
#   Lists compliance/eDiscovery cases.
# When to use:
#   eDiscovery case review.
# ============================================================

Get-ComplianceCase
