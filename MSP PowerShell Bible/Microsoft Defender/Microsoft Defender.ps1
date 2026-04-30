<#
.SYNOPSIS
    MSP PowerShell Bible - Microsoft Defender / Microsoft 365 Defender

.PURPOSE
    Incidents, alerts, advanced hunting, secure score, and vulnerability visibility.

.REQUIRED MODULE
    Microsoft.Graph
#>

# ============================================================
# CONNECT TO MICROSOFT DEFENDER / SECURITY GRAPH
# ============================================================

Install-Module Microsoft.Graph -Scope CurrentUser -Force
Import-Module Microsoft.Graph

Connect-MgGraph -Scopes `
    "SecurityIncident.Read.All", `
    "SecurityAlert.Read.All", `
    "SecurityEvents.Read.All"

Select-MgProfile -Name "v1.0"

# Confirm connection
Get-MgContext


# ============================================================
# Get-MgSecurityIncident
# What it does:
#   Lists Microsoft 365 Defender security incidents.
# When to use:
#   Security operations triage and incident review.
# ============================================================

Get-MgSecurityIncident |
    Select-Object Id, DisplayName, Severity, Status, CreatedDateTime


# ============================================================
# Get-MgSecurityAlert
# What it does:
#   Lists Microsoft security alerts.
# When to use:
#   Alert triage, alert export, incident evidence gathering.
# ============================================================

Get-MgSecurityAlert |
    Select-Object Id, Title, Severity, Status, CreatedDateTime


# ============================================================
# Invoke-MgSecurityRunHuntingQuery
# What it does:
#   Runs an advanced hunting query.
# When to use:
#   Threat hunting and incident investigation.
# ============================================================

$query = "DeviceEvents | where Timestamp > ago(1d) | where ActionType contains 'ProcessCreated' | take 50"
Invoke-MgSecurityRunHuntingQuery -Query $query


# ============================================================
# Get-MgSecuritySecureScore
# What it does:
#   Retrieves Microsoft Secure Score.
# When to use:
#   Security posture review and improvement tracking.
# ============================================================

Get-MgSecuritySecureScore |
    Select-Object Id, CurrentScore, MaxScore, CreatedDateTime


# ============================================================
# Get-MgSecurityVulnerability
# What it does:
#   Lists known vulnerabilities where available through Graph.
# When to use:
#   Vulnerability visibility and reporting.
# ============================================================

Get-MgSecurityVulnerability

# Note: Vulnerability data in Graph may be limited; consider using Microsoft Defender for Endpoint APIs for detailed vulnerability information. 

# ============================================================
# DISCONNECT FROM MICROSOFT GRAPH
# ============================================================
Disconnect-MgGraph

# End of Microsoft Defender / Microsoft 365 Defender PowerShell script.

# Note: Always ensure you have the necessary permissions and are compliant with your organization's policies when accessing security data through Microsoft Graph.

# For more advanced scenarios, consider exploring the Microsoft Graph Security API documentation for additional endpoints and capabilities.

# This script is intended for educational purposes and should be tested in a non-production environment before use.

# Always keep your modules up to date to ensure compatibility with the latest Microsoft Graph API changes.
# ============================================================

