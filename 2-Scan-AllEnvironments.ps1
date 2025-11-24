# Scan All Environments - Consolidated Report
# Scans all Power Platform environments and creates a single consolidated CSV report

param(
    [Parameter(Mandatory=$false)]
    [string]$CredentialFile = "sp-credentials.xml"
)

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$OutputFolder = "ConsolidatedReport_$timestamp"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Multi-Environment App Scanner" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check credential file
if (!(Test-Path $CredentialFile)) {
    Write-Host "ERROR: Credential file not found: $CredentialFile" -ForegroundColor Red
    Write-Host "Please run: .\1-Setup-Authentication.ps1" -ForegroundColor Yellow
    exit 1
}

# Load credentials
Write-Host "Loading Service Principal credentials..." -ForegroundColor Cyan
$credFile = Import-Clixml -Path $CredentialFile
$appId = $credFile.ApplicationId
$tenantId = $credFile.TenantId
$secret = $credFile.ClientSecret
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
$secretPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

Write-Host "Application ID: $appId" -ForegroundColor Gray
Write-Host "Tenant ID: $tenantId" -ForegroundColor Gray
Write-Host ""

# Get all environments
Write-Host "Retrieving environments..." -ForegroundColor Cyan

$pacOutput = pac admin list | Out-String
$lines = $pacOutput -split "`n"

$environments = @()
$headerPassed = $false

foreach ($line in $lines) {
    # Skip header
    if ($line -match "Active Environment") {
        $headerPassed = $true
        continue
    }
    
    if (!$headerPassed) { continue }
    
    # Parse environment lines (handle both with and without asterisk)
    if ($line.Trim() -match "^\*?\s*(.+?)\s+([a-f0-9\-]{36})\s+(https://[^\s]+)\s+") {
        $envName = $matches[1].Trim()
        $envId = $matches[2].Trim()
        $envUrl = $matches[3].Trim()
        
        # Skip trial environments or empty names
        if ($envName -and $envName -ne "Active" -and $envName -ne "Environment") {
            $environments += [PSCustomObject]@{
                DisplayName = $envName
                EnvironmentName = $envId
                URL = $envUrl
            }
        }
    }
}

Write-Host "Found $($environments.Count) environments" -ForegroundColor Green
Write-Host ""

# Create output folder
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

# Initialize consolidated data
$allApps = @()
$environmentSummary = @()

# Load Dataverse module
if (!(Get-Module -ListAvailable -Name "Microsoft.Xrm.Data.PowerShell")) {
    Install-Module -Name "Microsoft.Xrm.Data.PowerShell" -Force -AllowClobber -Scope CurrentUser
}
Import-Module Microsoft.Xrm.Data.PowerShell -WarningAction SilentlyContinue

# Scan each environment
$envCount = 0
foreach ($env in $environments) {
    $envCount++
    $envName = $env.DisplayName
    $envUrl = $env.URL
    $envId = $env.EnvironmentName
    
    Write-Host "[$envCount/$($environments.Count)] Scanning: $envName" -ForegroundColor Yellow
    Write-Host "    URL: $envUrl" -ForegroundColor Gray
    
    try {
        # Connect to Dataverse
        $conn = Connect-CrmOnline -ServerUrl $envUrl -OAuthClientId $appId -ClientSecret $secretPlain -ErrorAction Stop
        
        if (!$conn -or !$conn.IsReady) {
            Write-Host "    - Failed to connect" -ForegroundColor Red
            $environmentSummary += [PSCustomObject]@{
                Environment = $envName
                EnvironmentId = $envId
                URL = $envUrl
                Status = "Connection Failed"
                AppCount = 0
                SharedApps = 0
                OrphanedApps = 0
            }
            continue
        }
        
        # Fetch Unified Interface apps
        $fetchXml = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">'
        $fetchXml += '<entity name="appmodule">'
        $fetchXml += '<attribute name="appmoduleid" />'
        $fetchXml += '<attribute name="name" />'
        $fetchXml += '<attribute name="uniquename" />'
        $fetchXml += '<attribute name="description" />'
        $fetchXml += '<attribute name="createdon" />'
        $fetchXml += '<attribute name="modifiedon" />'
        $fetchXml += '<attribute name="publishedon" />'
        $fetchXml += '<attribute name="statecode" />'
        $fetchXml += '<filter type="and">'
        $fetchXml += '<condition attribute="clienttype" operator="eq" value="4" />'
        $fetchXml += '</filter>'
        $fetchXml += '</entity>'
        $fetchXml += '</fetch>'
        
        $apps = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml
        
        # Handle the results - Get-CrmRecordsByFetch returns a dictionary where each value IS the record
        # We need to extract the actual records
        $appsList = @()
        if ($apps -is [System.Collections.IDictionary]) {
            # Dictionary: iterate through keys and build array of records
            foreach ($key in $apps.Keys) {
                $appsList += $apps[$key]
            }
        } else {
            $appsList = @($apps)
        }
        
        Write-Host "    - Found $($appsList.Count) apps" -ForegroundColor Green
        
        $sharedCount = 0
        $orphanedCount = 0
        
        # Process each app
        foreach ($app in $appsList) {
            # Skip non-record entries (metadata)
            if (!$app.appmoduleid) {
                continue
            }
            
            $appGuid = $app.appmoduleid
            
            # Get app roles
            $rolesFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">'
            $rolesFetch += '<entity name="role">'
            $rolesFetch += '<attribute name="name" />'
            $rolesFetch += '<attribute name="roleid" />'
            $rolesFetch += '<link-entity name="appmoduleroles" from="roleid" to="roleid">'
            $rolesFetch += '<filter>'
            $rolesFetch += "<condition attribute=`"appmoduleid`" operator=`"eq`" value=`"{$appGuid}`" />"
            $rolesFetch += '</filter>'
            $rolesFetch += '</link-entity>'
            $rolesFetch += '</entity>'
            $rolesFetch += '</fetch>'
            
            try {
                $roles = Get-CrmRecordsByFetch -conn $conn -Fetch $rolesFetch -ErrorAction Stop
            
                # Convert Dictionary to array - same logic as apps
                $rolesList = @()
                if ($roles -is [System.Collections.IDictionary]) {
                    foreach ($key in $roles.Keys) {
                        $rolesList += $roles[$key]
                    }
                } else {
                    $rolesList = @($roles)
                }
            
                $roleNames = ($rolesList | ForEach-Object { $_.name }) -join "; "
                $roleCount = $rolesList.Count
            } catch {
                # If role fetch fails, just set to unknown
                $roleNames = "Error fetching roles"
                $roleCount = 0
            }
            
            # Get users with access to this app (through roles)
            $usersList = @()
            $teamsList = @()
            $totalUsers = 0
            $totalTeams = 0
            
            if ($roleCount -gt 0) {
                try {
                    # Get users assigned to these roles
                    $usersFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">'
                    $usersFetch += '<entity name="systemuser">'
                    $usersFetch += '<attribute name="fullname" />'
                    $usersFetch += '<attribute name="systemuserid" />'
                    $usersFetch += '<link-entity name="systemuserroles" from="systemuserid" to="systemuserid">'
                    $usersFetch += '<link-entity name="role" from="roleid" to="roleid">'
                    $usersFetch += '<link-entity name="appmoduleroles" from="roleid" to="roleid">'
                    $usersFetch += '<filter>'
                    $usersFetch += "<condition attribute=`"appmoduleid`" operator=`"eq`" value=`"{$appGuid}`" />"
                    $usersFetch += '</filter>'
                    $usersFetch += '</link-entity>'
                    $usersFetch += '</link-entity>'
                    $usersFetch += '</link-entity>'
                    $usersFetch += '<filter>'
                    $usersFetch += '<condition attribute="isdisabled" operator="eq" value="0" />'
                    $usersFetch += '</filter>'
                    $usersFetch += '</entity>'
                    $usersFetch += '</fetch>'
                    
                    $users = Get-CrmRecordsByFetch -conn $conn -Fetch $usersFetch -ErrorAction SilentlyContinue
                    
                    if ($users -is [System.Collections.IDictionary]) {
                        foreach ($key in $users.Keys) {
                            if ($users[$key].fullname) {
                                $usersList += $users[$key]
                            }
                        }
                    }
                    $totalUsers = $usersList.Count
                } catch {
                    $totalUsers = 0
                }
                
                try {
                    # Get teams assigned to these roles
                    $teamsFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">'
                    $teamsFetch += '<entity name="team">'
                    $teamsFetch += '<attribute name="name" />'
                    $teamsFetch += '<attribute name="teamid" />'
                    $teamsFetch += '<link-entity name="teamroles" from="teamid" to="teamid">'
                    $teamsFetch += '<link-entity name="role" from="roleid" to="roleid">'
                    $teamsFetch += '<link-entity name="appmoduleroles" from="roleid" to="roleid">'
                    $teamsFetch += '<filter>'
                    $teamsFetch += "<condition attribute=`"appmoduleid`" operator=`"eq`" value=`"{$appGuid}`" />"
                    $teamsFetch += '</filter>'
                    $teamsFetch += '</link-entity>'
                    $teamsFetch += '</link-entity>'
                    $teamsFetch += '</link-entity>'
                    $teamsFetch += '</entity>'
                    $teamsFetch += '</fetch>'
                    
                    $teams = Get-CrmRecordsByFetch -conn $conn -Fetch $teamsFetch -ErrorAction SilentlyContinue
                    
                    if ($teams -is [System.Collections.IDictionary]) {
                        foreach ($key in $teams.Keys) {
                            if ($teams[$key].name) {
                                $teamsList += $teams[$key]
                            }
                        }
                    }
                    $totalTeams = $teamsList.Count
                } catch {
                    $totalTeams = 0
                }
            }
            
            # Get last usage information from audit logs
            $lastUsed = $null
            $daysSinceLastUse = $null
            
            try {
                $auditFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" top="1">'
                $auditFetch += '<entity name="audit">'
                $auditFetch += '<attribute name="createdon" />'
                $auditFetch += '<filter>'
                $auditFetch += "<condition attribute=`"objectid`" operator=`"eq`" value=`"{$appGuid}`" />"
                $auditFetch += '<condition attribute="operation" operator="eq" value="64" />'  # Access operation
                $auditFetch += '</filter>'
                $auditFetch += '<order attribute="createdon" descending="true" />'
                $auditFetch += '</entity>'
                $auditFetch += '</fetch>'
                
                $auditRecords = Get-CrmRecordsByFetch -conn $conn -Fetch $auditFetch -ErrorAction SilentlyContinue
                
                if ($auditRecords -is [System.Collections.IDictionary]) {
                    $auditList = @()
                    foreach ($key in $auditRecords.Keys) {
                        if ($auditRecords[$key].createdon) {
                            $auditList += $auditRecords[$key]
                        }
                    }
                    if ($auditList.Count -gt 0) {
                        $lastUsed = $auditList[0].createdon
                        $daysSinceLastUse = [math]::Round((New-TimeSpan -Start $lastUsed -End (Get-Date)).TotalDays)
                    }
                }
            } catch {
                # Audit may not be enabled or accessible
                $lastUsed = $null
                $daysSinceLastUse = $null
            }
            
            # Determine state
            $appState = switch ($app.statecode) {
                "0" { "Active" }
                "1" { "Inactive" }
                default { "Unknown" }
            }
            
            # Format user and team lists
            $usersListStr = if ($usersList.Count -gt 0) { ($usersList | ForEach-Object { $_.fullname }) -join "; " } else { "" }
            $teamsListStr = if ($teamsList.Count -gt 0) { ($teamsList | ForEach-Object { $_.name }) -join "; " } else { "" }
            $sharedWithRoles = $roleNames
            
            $isShared = "No"
            $isOrphaned = "Yes"
            
            if ($roleCount -gt 0) {
                $isShared = "Yes"
                $isOrphaned = "No"
                $sharedCount++
            } else {
                $orphanedCount++
            }
            
            # Add to consolidated data
            $allApps += [PSCustomObject]@{
                Environment = $envName
                EnvironmentId = $envId
                EnvironmentURL = $envUrl
                AppName = $app.name
                UniqueName = $app.uniquename
                AppId = $app.appmoduleid
                Description = $app.description
                State = $appState
                CreatedOn = $app.createdon
                ModifiedOn = $app.modifiedon
                PublishedOn = $app.publishedon
                RoleCount = $roleCount
                SharedWithRoles = $sharedWithRoles
                SharedCount = $roleCount
                TotalUsers = $totalUsers
                TotalTeams = $totalTeams
                UsersList = $usersListStr
                TeamsList = $teamsListStr
                LastUsed = $lastUsed
                DaysSinceLastUse = $daysSinceLastUse
                IsShared = $isShared
                IsOrphaned = $isOrphaned
            }
        }
        
        $environmentSummary += [PSCustomObject]@{
            Environment = $envName
            EnvironmentId = $envId
            URL = $envUrl
            Status = "Success"
            AppCount = $appsList.Count
            SharedApps = $sharedCount
            OrphanedApps = $orphanedCount
        }
        
    } catch {
        Write-Host "    - Error: $($_.Exception.Message)" -ForegroundColor Red
        $environmentSummary += [PSCustomObject]@{
            Environment = $envName
            EnvironmentId = $envId
            URL = $envUrl
            Status = "Error"
            AppCount = 0
            SharedApps = 0
            OrphanedApps = 0
        }
    }
    
    Write-Host ""
}

# Clean up secret
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

# Export consolidated reports
Write-Host "Generating consolidated reports..." -ForegroundColor Cyan

# 1. All apps CSV (Excel-friendly)
$csvPath = Join-Path $OutputFolder "AllEnvironments_Apps.csv"
$allApps | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "  + All Apps Report: $csvPath" -ForegroundColor Green

# 2. Environment summary
$summaryPath = Join-Path $OutputFolder "EnvironmentSummary.csv"
$environmentSummary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host "  + Environment Summary: $summaryPath" -ForegroundColor Green

# 3. JSON export
$jsonPath = Join-Path $OutputFolder "CompleteData.json"
$jsonData = @{
    GeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    TotalEnvironments = $environments.Count
    TotalApps = $allApps.Count
    Environments = $environmentSummary
    Apps = $allApps
}
$jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
Write-Host "  + JSON Data: $jsonPath" -ForegroundColor Green

# 4. Text report
$reportPath = Join-Path $OutputFolder "DetailedReport.txt"
$reportLines = @()
$reportLines += "========================================"
$reportLines += "Multi-Environment App Analysis Report"
$reportLines += "========================================"
$reportLines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$reportLines += "Total Environments Scanned: $($environments.Count)"
$reportLines += "Total Apps Found: $($allApps.Count)"
$reportLines += ""
$reportLines += "========================================"
$reportLines += "Environment Summary"
$reportLines += "========================================"
$reportLines += ""

foreach ($envSum in $environmentSummary) {
    $reportLines += "Environment: $($envSum.Environment)"
    $reportLines += "  Status: $($envSum.Status)"
    $reportLines += "  Apps: $($envSum.AppCount)"
    $reportLines += "  Shared: $($envSum.SharedApps)"
    $reportLines += "  Orphaned: $($envSum.OrphanedApps)"
    $reportLines += "  URL: $($envSum.URL)"
    $reportLines += ""
}

$reportLines += ""
$reportLines += "========================================"
$reportLines += "Orphaned Apps (Not Shared)"
$reportLines += "========================================"
$reportLines += ""

$orphanedApps = $allApps | Where-Object { $_.IsOrphaned -eq "Yes" }
if ($orphanedApps.Count -eq 0) {
    $reportLines += "No orphaned apps found."
} else {
    foreach ($app in $orphanedApps) {
        $reportLines += "[$($app.Environment)] $($app.AppName)"
    }
}

$reportLines += ""
$reportLines += "========================================"
$reportLines += "All Apps by Environment"
$reportLines += "========================================"
$reportLines += ""

foreach ($env in ($allApps | Group-Object Environment)) {
    $reportLines += ""
    $reportLines += "--- $($env.Name) ---"
    foreach ($app in $env.Group) {
        if ($app.Roles) {
            $reportLines += "  - $($app.AppName) - Roles: $($app.Roles)"
        } else {
            $reportLines += "  - $($app.AppName) - NOT SHARED"
        }
    }
}

$reportLines | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "  + Detailed Report: $reportPath" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Scan Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  Environments Scanned: $($environments.Count)" -ForegroundColor White
Write-Host "  Total Apps Found: $($allApps.Count)" -ForegroundColor White
Write-Host "  Orphaned Apps: $($orphanedApps.Count)" -ForegroundColor White
Write-Host ""
Write-Host "Reports Location: $OutputFolder" -ForegroundColor Yellow
Write-Host ""
