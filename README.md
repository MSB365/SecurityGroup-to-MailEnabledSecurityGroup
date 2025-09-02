# Entra ID to Exchange Online Migration Scripts

This repository contains PowerShell scripts to migrate Entra ID security groups to Exchange Online mail-enabled security groups.

## Overview

The migration process consists of two main scripts:

1. **Export-EntraIDGroupMembers.ps1** - Exports members from Entra ID security groups
2. **Create-MailEnabledSecurityGroups.ps1** - Creates mail-enabled security groups in Exchange Online

## Prerequisites

### Required PowerShell Modules
- Microsoft.Graph.Authentication
- Microsoft.Graph.Groups  
- Microsoft.Graph.Users
- ExchangeOnlineManagement

### Required Permissions
- **Entra ID**: Group.Read.All, User.Read.All
- **Exchange Online**: Organization Management or equivalent permissions to create distribution groups

## Quick Start

### Step 1: Prepare Input File
Create a CSV file with the display names of your Entra ID security groups:

\`\`\`csv
DisplayName
IT Security Team
Marketing Department
Finance Team
\`\`\`

### Step 2: Export Group Members
```powershell
.\Export-EntraIDGroupMembers.ps1 -InputCsvPath ".\sample_groups.csv"
