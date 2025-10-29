# Model-Driven Apps Analysis Tools

Collection of PowerShell scripts for analyzing Microsoft Dataverse Model-Driven Applications.

## Overview

This repository contains tools to retrieve detailed information about Model-Driven Apps in Microsoft Power Platform environments, including sharing permissions, user access, team assignments, and usage analytics.

## Prerequisites

- PowerShell 5.1 or later
- Microsoft.PowerApps.Administration.PowerShell module
- Microsoft.Xrm.Data.PowerShell module
- Access to a Microsoft Power Platform environment
- Appropriate permissions to query Dataverse

## Installation

The scripts will automatically install required modules if not present:
- Microsoft.PowerApps.Administration.PowerShell
- Microsoft.Xrm.Data.PowerShell

## Scripts

### Analyze-AllApps.ps1

Main analysis script that performs comprehensive analysis of all Unified Interface apps.

**Features:**
- Lists all Model-Driven Apps (Unified Interface)
- Shows security role assignments
- Identifies users and teams with access (grouped by role)
- Detects orphaned apps
- Tracks last usage dates
- Exports results in multiple formats
- Supports custom environment selection

**Usage:**
```powershell
# Analyze apps in default environment
.\Analyze-AllApps.ps1

# Analyze apps in specific environment by name
.\Analyze-AllApps.ps1 -EnvironmentName "Production"

# Analyze apps in specific environment by ID
.\Analyze-AllApps.ps1 -EnvironmentName "00000000-0000-0000-0000-000000000000"
```

**Parameters:**
- `-EnvironmentName` (optional): Environment name or GUID. If not specified, uses default environment.

**Output:**
- CompleteAnalysis.json - Full data in JSON format
- DetailedReport.txt - Detailed text report with users/teams per role
- SharingReport.txt - Sharing summary by security roles
- UsersAndTeamsReport.txt - Complete users and teams access breakdown
- UsageReport.txt - Usage statistics sorted by last activity
- AppAnalysis.csv - Excel-compatible CSV with all information
- OrphanedApps.txt - List of orphaned apps (if any found)

### Find-SpecificApp.ps1

Search for a specific app by its GUID.

**Usage:**
```powershell
# Search in default environment
.\Find-SpecificApp.ps1

# Search specific app in default environment
.\Find-SpecificApp.ps1 -AppId "your-app-guid-here"

# Search specific app in custom environment
.\Find-SpecificApp.ps1 -AppId "your-app-guid-here" -EnvironmentName "Production"
```

**Parameters:**
- `-AppId` (optional): App GUID to search for. Defaults to example GUID.
- `-EnvironmentName` (optional): Environment name or GUID. If not specified, uses default environment.

**Output:**
- App details displayed in console
- Exported to timestamped folder with JSON and TXT files

### Find-AllUnifiedInterfaceApps.ps1

Retrieve all Unified Interface apps with basic information.

**Usage:**
```powershell
# Search in default environment
.\Find-AllUnifiedInterfaceApps.ps1

# Search in specific environment
.\Find-AllUnifiedInterfaceApps.ps1 -EnvironmentName "Production"
```

**Parameters:**
- `-EnvironmentName` (optional): Environment name or GUID. If not specified, uses default environment.

**Output:**
- List of all apps in console
- Exported to timestamped folder with individual JSON/TXT files per app

## Output Structure

All scripts create timestamped output folders to preserve historical analysis:

```
AppAnalysis_YYYYMMDD_HHMMSS/
├── CompleteAnalysis.json
├── DetailedReport.txt
├── SharingReport.txt
├── UsersAndTeamsReport.txt
├── UsageReport.txt
├── AppAnalysis.csv
└── OrphanedApps.txt (if applicable)
```

## CSV Export Columns

The AppAnalysis.csv file includes:
- AppName
- AppId
- UniqueName
- State
- IsOrphaned
- SharedWithRoles
- SharedCount
- TotalUsers
- TotalTeams
- UsersList
- TeamsList
- LastUsed
- DaysSinceLastUse
- CreatedOn
- CreatedBy
- ModifiedOn
- ModifiedBy

## Authentication

Scripts use interactive OAuth authentication through the PowerApps cmdlets. You will be prompted to sign in when running the scripts.

## Use Cases

- Audit app permissions and user access
- Identify unused or orphaned applications
- Review security role assignments
- Plan app cleanup or consolidation
- Document application inventory
- Track user and team access patterns
- Generate compliance reports

## Notes

- Scripts connect to the default Power Platform environment unless specified
- Can specify environment by display name or GUID using `-EnvironmentName` parameter
- Analysis includes only Unified Interface apps (ClientType = 4)
- User and team lists show only active/enabled accounts
- Audit data may not be available in all environments
- Last usage falls back to last modified date if audit is disabled

## License

This project is provided as-is for educational and operational purposes.

## Author

Created for analyzing Microsoft Dataverse Model-Driven Applications.
