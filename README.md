# Power Platform Model-Driven Apps Scanner

Automated tool to scan and analyze Model-Driven Apps across all Power Platform environments using Service Principal authentication.

## Features

- Service Principal authentication for automated scanning
- Multi-environment scanning with single command
- Consolidated CSV report with all apps from all environments
- Security role analysis for each app
- Environment-level statistics and summaries

## Quick Start

### 1. Setup & Deploy Application User

```powershell
.\1-Setup.ps1
```

This script will:
- Authenticate interactively via browser (required for PAC admin operations)
- Save Service Principal credentials to encrypted file
- List all Power Platform environments
- Deploy application user with System Administrator role to each environment

### 2. Scan All Environments

```powershell
.\2-Scan-AllEnvironments.ps1
```

This script will:
- Load Service Principal credentials
- Connect to each environment automatically
- Scan all Model-Driven Apps
- Generate consolidated reports in timestamped folder

## Output Files

All reports are saved in: `ConsolidatedReport_YYYYMMDD_HHMMSS/`

### Main Report: `AllEnvironments_Apps.csv`

Excel-friendly CSV with complete app inventory across all environments.

**Columns:**
- `Environment` - Environment display name
- `EnvironmentId` - Environment GUID
- `EnvironmentURL` - Dataverse organization URL
- `AppName` - Application display name
- `UniqueName` - Unique identifier
- `AppId` - App module GUID
- `Description` - App description
- `CreatedOn` - Creation timestamp
- `ModifiedOn` - Last modified timestamp
- `PublishedOn` - Last published timestamp
- `RoleCount` - Number of assigned security roles
- `Roles` - Security role names (semicolon-separated)
- `IsShared` - "Yes" if shared with roles
- `IsOrphaned` - "Yes" if not shared with any role

### Additional Reports

- `EnvironmentSummary.csv` - Statistics per environment (app count, shared/orphaned apps)
- `DetailedReport.txt` - Human-readable text report
- `CompleteData.json` - Full data in JSON format

## Prerequisites

- **PowerShell 5.1** (Windows PowerShell Desktop edition)
- **Power Platform CLI** (`pac`) - [Install Guide](https://learn.microsoft.com/power-platform/developer/cli/introduction)
- **Azure AD Service Principal** with:
  - Application (client) ID
  - Client secret
  - Tenant ID

## Service Principal Setup

### 1. Create App Registration in Azure AD

1. Go to **Azure Portal** → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Enter a name (e.g., "Model-Driven App Scanner")
4. Click **Register**
5. Note the **Application (client) ID** and **Tenant ID**
6. Go to **Certificates & secrets** → **New client secret**
7. Note the **client secret value** (copy immediately, it won't be shown again)

### 2. Configure Credentials

Set environment variables with your Service Principal credentials:

```powershell
# Set environment variables (PowerShell)
$env:PP_APP_ID = "your-application-id-here"
$env:PP_CLIENT_SECRET = "your-client-secret-here"
$env:PP_TENANT_ID = "your-tenant-id-here"
```

**Note:** These are temporary session variables. To make them permanent:
- **Windows:** Use System Properties → Environment Variables
- **PowerShell Profile:** Add the commands to your `$PROFILE` script

Alternatively, copy `.env.example` to `.env` and edit it (requires manual loading).

### 3. Run Setup

```powershell
.\1-Setup.ps1
```

The script will deploy the application user to all environments automatically.

## Individual Scripts

### Analyze-AllApps.ps1

Single environment analyzer (supports both interactive and Service Principal modes).

**Interactive mode:**
```powershell
.\Analyze-AllApps.ps1 -EnvironmentName "My Environment"
```

**Service Principal mode:**
```powershell
.\Analyze-AllApps.ps1 -EnvironmentName "My Environment" -UseServicePrincipal -CredentialFile "sp-credentials.xml"
```

## Workflow

```
┌─────────────────────────┐
│   1-Setup.ps1           │
│   - Auth interactively  │
│   - Save SP credentials │
│   - Deploy app users    │
└───────────┬─────────────┘
            │ Creates sp-credentials.xml
            ↓
┌─────────────────────────┐
│ 2-Scan-AllEnvironments  │
│ .ps1                    │
│   - Load SP credentials │
│   - Scan all envs       │
│   - Generate reports    │
└───────────┬─────────────┘
            │
            ↓
┌─────────────────────────┐
│ ConsolidatedReport_*/   │
│   - AllEnvironments_    │
│     Apps.csv            │
│   - EnvironmentSummary  │
│     .csv                │
│   - JSON & TXT reports  │
└─────────────────────────┘
```

## Troubleshooting

### HttpRequestException - Browser non si apre

Se il browser non si apre durante `pac auth create`, usa il **device code flow**:

```powershell
pac auth create --deviceCode
```

Questo mostra un codice da inserire manualmente in un browser (anche su altro dispositivo).

**Vedi [TROUBLESHOOTING.md](TROUBLESHOOTING.md) per soluzioni dettagliate.**

### "Connection Failed" for an environment
- The application user must exist in the environment
- The application user must have **System Administrator** role
- Ensure the environment has a Dataverse database provisioned

### "Credential file not found"
Run `.\1-Setup.ps1` first to create the credentials file.

### "pac command not found"
Install Power Platform CLI:
```powershell
# Using dotnet tool
dotnet tool install --global Microsoft.PowerApps.CLI.Tool

# Or download installer from:
# https://aka.ms/PowerAppsCLI
```

### "Application user deployment failed"
Some environments may have restrictions on creating application users. The error typically indicates:
- Insufficient user permissions (need admin role in that environment)
- Environment-level security policies

You can manually add the application user via Power Platform Admin Center.

## Security Notes

- Credentials in `sp-credentials.xml` are encrypted using **Windows DPAPI**
- Encrypted credentials are machine and user specific
- **Never commit** `sp-credentials.xml` to version control (already in `.gitignore`)
- Service Principal has access only to environments where application user is deployed

## Technical Details

### Authentication Flow

1. **Setup** (`1-Setup.ps1`):
   - Uses **interactive authentication** for PAC admin operations
   - Saves Service Principal credentials for scanner

2. **Scanning** (`2-Scan-AllEnvironments.ps1`):
   - Uses **Service Principal authentication** to connect to Dataverse
   - No user interaction required

### Data Extraction

- Uses **FetchXML queries** via `Microsoft.Xrm.Data.PowerShell` module
- Filters for **Unified Interface apps** (clienttype = 4)
- Links with `appmoduleroles` entity to get security role assignments
- Extracts metadata: creation dates, modification dates, descriptions

### Performance

- Scans run sequentially per environment
- Typical scan time: 5-15 seconds per environment
- Output generation: < 1 second

## License

MIT
