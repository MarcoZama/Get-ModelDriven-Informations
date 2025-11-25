<#
.SYNOPSIS
    Generate CompleteAnalysis JSON report with UsersByRole structure
.DESCRIPTION
    This script scans all environments and generates a JSON report 
    with the same structure as CompleteAnalysis.json, including
    detailed UsersByRole and TeamsByRole information.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$CredentialFile = "sp-credentials.xml"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Complete Analysis Report Generator" -ForegroundColor Cyan
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
$applicationId = $credFile.ApplicationId
$tenantId = $credFile.TenantId
$secret = $credFile.ClientSecret
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
$clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
Write-Host "Application ID: $applicationId" -ForegroundColor Gray
Write-Host "Tenant ID: $tenantId" -ForegroundColor Gray
Write-Host ""

# Get all environments using PAC CLI
Write-Host "Retrieving environments..." -ForegroundColor Yellow
$envListJson = pac admin list --json 2>$null
$environments = $envListJson | ConvertFrom-Json

Write-Host "Found $($environments.Count) environments" -ForegroundColor Green
Write-Host ""

# Import required module
Import-Module Microsoft.Xrm.Data.PowerShell

# Array to hold all apps data
$allAppsData = @()
$envCounter = 0

# Process each environment
foreach ($env in $environments) {
    $envCounter++
    $envName = $env.FriendlyName
    $envId = $env.EnvironmentId
    $envUrl = $env.EnvironmentUrl
    
    Write-Host "[$envCounter/$($environments.Count)] Scanning: $envName" -ForegroundColor Cyan
    Write-Host "    URL: $envUrl" -ForegroundColor Gray
    
    # Connect to Dataverse
    try {
        $conn = Connect-CrmOnline -ServerUrl $envUrl -OAuthClientId $applicationId -ClientSecret $clientSecret -ErrorAction Stop
        
        if (-not $conn.IsReady) {
            Write-Host "    - Failed to connect" -ForegroundColor Red
            continue
        }
    } catch {
        Write-Host "    - Connection error: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }
    
    # Fetch apps with creator and modifier info
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
    $fetchXml += '<link-entity name="systemuser" from="systemuserid" to="createdby" alias="creator">'
    $fetchXml += '<attribute name="fullname" />'
    $fetchXml += '</link-entity>'
    $fetchXml += '<link-entity name="systemuser" from="systemuserid" to="modifiedby" alias="modifier">'
    $fetchXml += '<attribute name="fullname" />'
    $fetchXml += '</link-entity>'
    $fetchXml += '<filter type="and">'
    $fetchXml += '<condition attribute="clienttype" operator="eq" value="4" />'
    $fetchXml += '</filter>'
    $fetchXml += '</entity>'
    $fetchXml += '</fetch>'
    
    $apps = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml
    
    # Convert to array
    $appsList = @()
    if ($apps -is [System.Collections.IDictionary]) {
        foreach ($key in $apps.Keys) {
            $appsList += $apps[$key]
        }
    } else {
        $appsList = @($apps)
    }
    
    Write-Host "    - Found $($appsList.Count) apps" -ForegroundColor Green
    
    # Process each app
    foreach ($app in $appsList) {
        if (!$app.appmoduleid) { continue }
        
        $appGuid = $app.appmoduleid
        
        # Get roles for this app
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
        
        $roles = Get-CrmRecordsByFetch -conn $conn -Fetch $rolesFetch -ErrorAction SilentlyContinue
        
        $rolesList = @()
        if ($roles -is [System.Collections.IDictionary]) {
            foreach ($key in $roles.Keys) {
                $rolesList += $roles[$key]
            }
        } else {
            $rolesList = @($roles)
        }
        
        # Initialize data structures
        $usersByRole = @{}
        $teamsByRole = @{}
        $sharedWith = @()
        $sharedWithUsers = @()
        $sharedWithTeams = @()
        
        # Process each role
        foreach ($role in $rolesList) {
            $roleName = $role.name
            $roleId = $role.roleid
            
            # Skip roles without a name
            if ([string]::IsNullOrWhiteSpace($roleName)) {
                continue
            }
            
            $sharedWith += $roleName
            
            # Initialize arrays for this role
            $usersByRole[$roleName] = @()
            $teamsByRole[$roleName] = @()
            
            # Get users for this role
            try {
                $usersFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">'
                $usersFetch += '<entity name="systemuser">'
                $usersFetch += '<attribute name="fullname" />'
                $usersFetch += '<attribute name="systemuserid" />'
                $usersFetch += '<attribute name="applicationid" />'
                $usersFetch += '<link-entity name="systemuserroles" from="systemuserid" to="systemuserid">'
                $usersFetch += '<filter>'
                $usersFetch += "<condition attribute=`"roleid`" operator=`"eq`" value=`"{$roleId}`" />"
                $usersFetch += '</filter>'
                $usersFetch += '</link-entity>'
                $usersFetch += '<filter>'
                $usersFetch += '<condition attribute="isdisabled" operator="eq" value="0" />'
                $usersFetch += '</filter>'
                $usersFetch += '</entity>'
                $usersFetch += '</fetch>'
                
                $roleUsers = Get-CrmRecordsByFetch -conn $conn -Fetch $usersFetch -ErrorAction SilentlyContinue
                
                if ($roleUsers -is [System.Collections.IDictionary]) {
                    $userList = @()
                    foreach ($key in $roleUsers.Keys) {
                        $userList += $roleUsers[$key]
                    }
                    
                    foreach ($user in $userList) {
                        if ($user.fullname) {
                            $userName = $user.fullname
                            # Mark application users with #
                            if ($user.applicationid -ne $null) {
                                $userName = "# $userName"
                            }
                            # Ensure each user is a separate element
                            if (-not ($usersByRole[$roleName] -contains $userName)) {
                                $usersByRole[$roleName] += $userName
                            }
                            if ($sharedWithUsers -notcontains $userName) {
                                $sharedWithUsers += $userName
                            }
                        }
                    }
                }
            } catch {
                # Continue with next role
            }
            
            # Get teams for this role
            try {
                $teamsFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">'
                $teamsFetch += '<entity name="team">'
                $teamsFetch += '<attribute name="name" />'
                $teamsFetch += '<attribute name="teamid" />'
                $teamsFetch += '<link-entity name="teamroles" from="teamid" to="teamid">'
                $teamsFetch += '<filter>'
                $teamsFetch += "<condition attribute=`"roleid`" operator=`"eq`" value=`"{$roleId}`" />"
                $teamsFetch += '</filter>'
                $teamsFetch += '</link-entity>'
                $teamsFetch += '</entity>'
                $teamsFetch += '</fetch>'
                
                $roleTeams = Get-CrmRecordsByFetch -conn $conn -Fetch $teamsFetch -ErrorAction SilentlyContinue
                
                if ($roleTeams -is [System.Collections.IDictionary]) {
                    foreach ($key in $roleTeams.Keys) {
                        if ($roleTeams[$key].name) {
                            $teamName = $roleTeams[$key].name
                            $teamsByRole[$roleName] += $teamName
                            if ($sharedWithTeams -notcontains $teamName) {
                                $sharedWithTeams += $teamName
                            }
                        }
                    }
                }
            } catch {
                # Continue with next role
            }
        }
        
        # Get last usage from audit
        $lastUsed = $null
        try {
            $auditFetch = '<fetch version="1.0" output-format="xml-platform" mapping="logical" top="1">'
            $auditFetch += '<entity name="audit">'
            $auditFetch += '<attribute name="createdon" />'
            $auditFetch += '<filter>'
            $auditFetch += "<condition attribute=`"objectid`" operator=`"eq`" value=`"{$appGuid}`" />"
            $auditFetch += '<condition attribute="operation" operator="eq" value="64" />'
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
                }
            }
        } catch {
            $lastUsed = $null
        }
        
        # Determine state (translate to Italian like in CompleteAnalysis)
        $appState = switch ($app.statecode) {
            "Active" { "Attivo" }
            "Inactive" { "Inattivo" }
            0 { "Attivo" }
            1 { "Inattivo" }
            default { "Sconosciuto" }
        }
        
        # Format dates
        $createdOnFormatted = if ($app.createdon) { (Get-Date $app.createdon -Format "dd/MM/yyyy HH:mm") } else { "" }
        $modifiedOnFormatted = if ($app.modifiedon) { (Get-Date $app.modifiedon -Format "dd/MM/yyyy HH:mm") } else { "" }
        $publishedOnFormatted = if ($app.publishedon) { (Get-Date $app.publishedon -Format "dd/MM/yyyy HH:mm") } else { "" }
        $lastUsedFormatted = if ($lastUsed) { (Get-Date $lastUsed -Format "dd/MM/yyyy HH:mm") } else { "" }
        
        # Create app object
        $appData = [PSCustomObject]@{
            AppName = $app.name
            UniqueName = $app.uniquename
            AppId = $app.appmoduleid
            State = $appState
            CreatedOn = $createdOnFormatted
            CreatedBy = if ($app.'creator.fullname') { $app.'creator.fullname' } else { "" }
            ModifiedOn = $modifiedOnFormatted
            ModifiedBy = if ($app.'modifier.fullname') { $app.'modifier.fullname' } else { "" }
            PublishedOn = $publishedOnFormatted
            LastUsed = $lastUsedFormatted
            SharedWith = $sharedWith
            SharedCount = $sharedWith.Count
            SharedWithUsers = $sharedWithUsers
            SharedWithTeams = $sharedWithTeams
            UsersByRole = $usersByRole
            TeamsByRole = $teamsByRole
            UsageCount = 0  # Placeholder
            IsOrphaned = ($sharedWith.Count -eq 0)
        }
        
        $allAppsData += $appData
    }
    
    Write-Host ""
}

# Clean up secret
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

# Generate output folder
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFolder = "CompleteAnalysis_$timestamp"
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

# Export to JSON
$jsonPath = Join-Path $outputFolder "CompleteAnalysis.json"
$allAppsData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Complete Analysis Generated!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total Apps: $($allAppsData.Count)" -ForegroundColor Green
Write-Host "Output: $jsonPath" -ForegroundColor Green
Write-Host ""
