# Script to analyze Unified Interface apps:
# 1. Check if shared and with whom
# 2. Check if orphaned (no owner or creator)
# 3. Check last usage date

param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentName = $null
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Advanced App Analysis" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check if module is installed
if (!(Get-Module -ListAvailable -Name "Microsoft.Xrm.Data.PowerShell")) {
    Write-Host "Installing Microsoft.Xrm.Data.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name "Microsoft.Xrm.Data.PowerShell" -Force -AllowClobber -Scope CurrentUser
}

Import-Module Microsoft.Xrm.Data.PowerShell

try {
    # Connect to Power Platform
    if ([string]::IsNullOrEmpty($EnvironmentName)) {
        Write-Host "`nConnecting to default environment..." -ForegroundColor Cyan
    } else {
        Write-Host "`nConnecting to environment: $EnvironmentName..." -ForegroundColor Cyan
    }
    
    if (!(Get-Module -ListAvailable -Name "Microsoft.PowerApps.Administration.PowerShell")) {
        Install-Module -Name "Microsoft.PowerApps.Administration.PowerShell" -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module Microsoft.PowerApps.Administration.PowerShell
    
    Add-PowerAppsAccount
    
    # Get target environment
    if ([string]::IsNullOrEmpty($EnvironmentName)) {
        # Get default environment
        $targetEnv = Get-AdminPowerAppEnvironment | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1
        
        if (!$targetEnv) {
            Write-Host "No default environment found. Getting first available environment..." -ForegroundColor Yellow
            $targetEnv = Get-AdminPowerAppEnvironment | Where-Object { $_.CommonDataServiceDatabaseProvisioningState -eq "Succeeded" } | Select-Object -First 1
        }
    } else {
        # Get specified environment by name or ID
        $targetEnv = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentName
        
        if (!$targetEnv) {
            # Try to find by display name
            $targetEnv = Get-AdminPowerAppEnvironment | Where-Object { 
                $_.DisplayName -eq $EnvironmentName -and 
                $_.CommonDataServiceDatabaseProvisioningState -eq "Succeeded" 
            } | Select-Object -First 1
        }
    }
    
    if (!$targetEnv) {
        Write-Host "ERROR: Environment not found!" -ForegroundColor Red
        exit 1
    }
    
    $envUrl = $targetEnv.Internal.properties.linkedEnvironmentMetadata.instanceUrl
    
    Write-Host "Environment: $($targetEnv.DisplayName)" -ForegroundColor Green
    Write-Host "URL: $envUrl`n" -ForegroundColor Gray
    
    # Connect to Dataverse
    Write-Host "Connecting to Dataverse..." -ForegroundColor Cyan
    $conn = Connect-CrmOnline -ServerUrl $envUrl -ForceOAuth
    
    if (!$conn -or !$conn.IsReady) {
        Write-Host "ERROR: Failed to connect to Dataverse" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Connected successfully!`n" -ForegroundColor Green
    
    # Get all Unified Interface apps
    Write-Host "Retrieving all Unified Interface apps..." -ForegroundColor Cyan
    
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="appmodule">
    <attribute name="appmoduleid" />
    <attribute name="name" />
    <attribute name="uniquename" />
    <attribute name="createdon" />
    <attribute name="modifiedon" />
    <attribute name="publishedon" />
    <attribute name="createdby" />
    <attribute name="modifiedby" />
    <attribute name="statecode" />
    <attribute name="statuscode" />
    <filter type="and">
      <condition attribute="clienttype" operator="eq" value="4" />
    </filter>
    <order attribute="name" descending="false" />
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml
    $totalApps = $result.CrmRecords.Count
    
    Write-Host "Found $totalApps apps. Analyzing...`n" -ForegroundColor Green
    
    $analysisResults = @()
    $counter = 1
    
    foreach ($app in $result.CrmRecords) {
        Write-Host "[$counter/$totalApps] Analyzing: $($app.name)..." -ForegroundColor Gray
        
        $appAnalysis = @{
            AppId = $app.appmoduleid
            AppName = $app.name
            UniqueName = $app.uniquename
            State = $app.statecode
            CreatedOn = $app.createdon
            ModifiedOn = $app.modifiedon
            PublishedOn = $app.publishedon
            CreatedBy = $app.createdby
            ModifiedBy = $app.modifiedby
            IsOrphaned = $false
            SharedWith = @()
            SharedCount = 0
            SharedWithUsers = @()
            SharedWithTeams = @()
            UsersByRole = @{}
            TeamsByRole = @{}
            LastUsed = $null
            UsageCount = 0
        }
        
        # Check if orphaned (no creator)
        if ([string]::IsNullOrEmpty($app.createdby)) {
            $appAnalysis.IsOrphaned = $true
        }
        
        # Get app role assignments (sharing information)
        $sharesFetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="appmoduleroles">
    <attribute name="appmoduleroleid" />
    <attribute name="appmoduleid" />
    <attribute name="roleid" />
    <filter type="and">
      <condition attribute="appmoduleid" operator="eq" value="$($app.appmoduleid)" />
    </filter>
    <link-entity name="role" from="roleid" to="roleid" alias="role">
      <attribute name="name" />
      <attribute name="roleid" />
    </link-entity>
  </entity>
</fetch>
"@
        
        try {
            $sharesResult = Get-CrmRecordsByFetch -conn $conn -Fetch $sharesFetchXml
            if ($sharesResult.CrmRecords.Count -gt 0) {
                foreach ($share in $sharesResult.CrmRecords) {
                    $roleName = $share.original["role.name"]
                    $roleId = $share.original["role.roleid"]
                    
                    if ($roleName) {
                        $appAnalysis.SharedWith += $roleName
                        
                        # Get users with this role
                        $userRoleFetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="systemuser">
    <attribute name="systemuserid" />
    <attribute name="fullname" />
    <attribute name="domainname" />
    <attribute name="isdisabled" />
    <filter type="and">
      <condition attribute="isdisabled" operator="eq" value="0" />
    </filter>
    <link-entity name="systemuserroles" from="systemuserid" to="systemuserid">
      <link-entity name="role" from="roleid" to="roleid">
        <filter type="and">
          <condition attribute="roleid" operator="eq" value="$roleId" />
        </filter>
      </link-entity>
    </link-entity>
    <order attribute="fullname" descending="false" />
  </entity>
</fetch>
"@
                        
                        try {
                            $usersResult = Get-CrmRecordsByFetch -conn $conn -Fetch $userRoleFetchXml
                            $usersList = @()
                            foreach ($user in $usersResult.CrmRecords) {
                                $userName = $user.fullname
                                if ($userName) {
                                    $usersList += $userName
                                    if (-not ($appAnalysis.SharedWithUsers -contains $userName)) {
                                        $appAnalysis.SharedWithUsers += $userName
                                    }
                                }
                            }
                            if ($usersList.Count -gt 0) {
                                $appAnalysis.UsersByRole[$roleName] = $usersList
                            }
                        } catch {
                            # Users not available
                        }
                        
                        # Get teams with this role
                        $teamRoleFetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="team">
    <attribute name="teamid" />
    <attribute name="name" />
    <filter type="and">
      <condition attribute="isdefault" operator="eq" value="0" />
    </filter>
    <link-entity name="teamroles" from="teamid" to="teamid">
      <link-entity name="role" from="roleid" to="roleid">
        <filter type="and">
          <condition attribute="roleid" operator="eq" value="$roleId" />
        </filter>
      </link-entity>
    </link-entity>
    <order attribute="name" descending="false" />
  </entity>
</fetch>
"@
                        
                        try {
                            $teamsResult = Get-CrmRecordsByFetch -conn $conn -Fetch $teamRoleFetchXml
                            $teamsList = @()
                            foreach ($team in $teamsResult.CrmRecords) {
                                $teamName = $team.name
                                if ($teamName) {
                                    $teamsList += $teamName
                                    if (-not ($appAnalysis.SharedWithTeams -contains $teamName)) {
                                        $appAnalysis.SharedWithTeams += $teamName
                                    }
                                }
                            }
                            if ($teamsList.Count -gt 0) {
                                $appAnalysis.TeamsByRole[$roleName] = $teamsList
                            }
                        } catch {
                            # Teams not available
                        }
                    }
                }
                $appAnalysis.SharedCount = $sharesResult.CrmRecords.Count
            }
        } catch {
            # Sharing info not available
        }
        
        # Try to get last usage from audit logs
        $auditFetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="audit">
    <attribute name="createdon" />
    <attribute name="action" />
    <attribute name="operation" />
    <filter type="and">
      <condition attribute="objectid" operator="eq" value="$($app.appmoduleid)" />
      <condition attribute="action" operator="in">
        <value>64</value>
        <value>1</value>
      </condition>
    </filter>
    <order attribute="createdon" descending="true" />
  </entity>
</fetch>
"@
        
        try {
            $auditResult = Get-CrmRecordsByFetch -conn $conn -Fetch $auditFetchXml
            if ($auditResult.CrmRecords.Count -gt 0) {
                $appAnalysis.LastUsed = $auditResult.CrmRecords[0].createdon
                $appAnalysis.UsageCount = $auditResult.CrmRecords.Count
            }
        } catch {
            # Audit not available or not enabled
        }
        
        # If no audit data, use modified date as proxy for last activity
        if (!$appAnalysis.LastUsed -and $app.modifiedon) {
            $appAnalysis.LastUsed = $app.modifiedon
        }
        
        $analysisResults += $appAnalysis
        $counter++
    }
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Analysis Complete!" -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Green
    
    # Display summary statistics
    $orphanedApps = ($analysisResults | Where-Object { $_.IsOrphaned -eq $true }).Count
    $sharedApps = ($analysisResults | Where-Object { $_.SharedCount -gt 0 }).Count
    $notSharedApps = ($analysisResults | Where-Object { $_.SharedCount -eq 0 }).Count
    
    Write-Host "Summary Statistics:" -ForegroundColor Cyan
    Write-Host "  Total Apps Analyzed: $totalApps" -ForegroundColor White
    Write-Host "  Orphaned Apps: " -ForegroundColor Yellow -NoNewline
    Write-Host "$orphanedApps" -ForegroundColor $(if($orphanedApps -gt 0){"Red"}else{"Green"})
    Write-Host "  Shared Apps: $sharedApps" -ForegroundColor White
    Write-Host "  Not Shared Apps: $notSharedApps" -ForegroundColor White
    
    # Export results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $folder = "AppAnalysis_$timestamp"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    
    # Save complete analysis to JSON
    $analysisResults | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folder\CompleteAnalysis.json" -Encoding UTF8
    
    # Create detailed report
    $reportText = "UNIFIED INTERFACE APPS - DETAILED ANALYSIS REPORT`n"
    $reportText += "=" * 80 + "`n`n"
    $reportText += "Analysis Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $reportText += "Environment: $($targetEnv.DisplayName)`n"
    $reportText += "Total Apps: $totalApps`n`n"
    $reportText += "SUMMARY STATISTICS`n"
    $reportText += "-" * 80 + "`n"
    $reportText += "Orphaned Apps: $orphanedApps`n"
    $reportText += "Shared Apps: $sharedApps`n"
    $reportText += "Not Shared Apps: $notSharedApps`n`n"
    $reportText += "=" * 80 + "`n`n"
    
    # Detailed app information
    $counter = 1
    foreach ($appData in $analysisResults | Sort-Object AppName) {
        $reportText += "[$counter] $($appData.AppName)`n"
        $reportText += "-" * 80 + "`n"
        $reportText += "  App ID: $($appData.AppId)`n"
        $reportText += "  Unique Name: $($appData.UniqueName)`n"
        $reportText += "  State: $($appData.State)`n"
        $reportText += "`n"
        
        # Orphaned status
        $reportText += "  IS ORPHANED: "
        if ($appData.IsOrphaned) {
            $reportText += "YES (WARNING!)`n"
        } else {
            $reportText += "No`n"
        }
        $reportText += "`n"
        
        # Sharing information
        $reportText += "  SHARING INFORMATION:`n"
        if ($appData.SharedCount -eq 0) {
            $reportText += "    Status: NOT SHARED with any security roles`n"
        } else {
            $reportText += "    Status: SHARED with $($appData.SharedCount) security role(s)`n"
            $reportText += "    Shared with roles:`n"
            foreach ($role in $appData.SharedWith) {
                $reportText += "      - $role`n"
                
                # Show users with this role
                if ($appData.UsersByRole.ContainsKey($role) -and $appData.UsersByRole[$role].Count -gt 0) {
                    $reportText += "        Users ($($appData.UsersByRole[$role].Count)):`n"
                    foreach ($user in $appData.UsersByRole[$role]) {
                        $reportText += "          * $user`n"
                    }
                }
                
                # Show teams with this role
                if ($appData.TeamsByRole.ContainsKey($role) -and $appData.TeamsByRole[$role].Count -gt 0) {
                    $reportText += "        Teams ($($appData.TeamsByRole[$role].Count)):`n"
                    foreach ($team in $appData.TeamsByRole[$role]) {
                        $reportText += "          + $team`n"
                    }
                }
            }
            
            # Summary
            $reportText += "`n    Total unique users with access: $($appData.SharedWithUsers.Count)`n"
            $reportText += "    Total unique teams with access: $($appData.SharedWithTeams.Count)`n"
        }
        $reportText += "`n"
        
        # Usage information
        $reportText += "  USAGE INFORMATION:`n"
        if ($appData.LastUsed) {
            $reportText += "    Last Activity: $($appData.LastUsed)`n"
            $daysSinceUse = ((Get-Date) - [DateTime]$appData.LastUsed).Days
            $reportText += "    Days Since Last Activity: $daysSinceUse days`n"
            if ($appData.UsageCount -gt 0) {
                $reportText += "    Audit Records Found: $($appData.UsageCount)`n"
            }
        } else {
            $reportText += "    Last Activity: Unknown (no audit data)`n"
        }
        $reportText += "`n"
        
        # Metadata
        $reportText += "  METADATA:`n"
        $reportText += "    Created: $($appData.CreatedOn) by $($appData.CreatedBy)`n"
        $reportText += "    Modified: $($appData.ModifiedOn) by $($appData.ModifiedBy)`n"
        if ($appData.PublishedOn) {
            $reportText += "    Published: $($appData.PublishedOn)`n"
        }
        $reportText += "`n"
        $reportText += "  Access URL: ${envUrl}main.aspx?appid=$($appData.AppId)`n"
        $reportText += "`n" + ("=" * 80) + "`n`n"
        
        $counter++
    }
    
    $reportText | Out-File -FilePath "$folder\DetailedReport.txt" -Encoding UTF8
    
    # Create orphaned apps report
    $orphanedList = $analysisResults | Where-Object { $_.IsOrphaned -eq $true }
    if ($orphanedList.Count -gt 0) {
        $orphanedText = "ORPHANED APPS REPORT`n"
        $orphanedText += "=" * 80 + "`n`n"
        $orphanedText += "Total Orphaned Apps: $($orphanedList.Count)`n`n"
        
        foreach ($app in $orphanedList) {
            $orphanedText += "- $($app.AppName)`n"
            $orphanedText += "  ID: $($app.AppId)`n"
            $orphanedText += "  Created By: $($app.CreatedBy)`n`n"
        }
        
        $orphanedText | Out-File -FilePath "$folder\OrphanedApps.txt" -Encoding UTF8
    }
    
    # Create sharing summary
    $sharingText = "APP SHARING SUMMARY`n"
    $sharingText += "=" * 80 + "`n`n"
    
    $sharingText += "NOT SHARED APPS ($notSharedApps):`n"
    $sharingText += "-" * 80 + "`n"
    foreach ($app in ($analysisResults | Where-Object { $_.SharedCount -eq 0 } | Sort-Object AppName)) {
        $sharingText += "- $($app.AppName) (ID: $($app.AppId))`n"
    }
    $sharingText += "`n`n"
    
    $sharingText += "SHARED APPS ($sharedApps):`n"
    $sharingText += "-" * 80 + "`n"
    foreach ($app in ($analysisResults | Where-Object { $_.SharedCount -gt 0 } | Sort-Object AppName)) {
        $sharingText += "- $($app.AppName) - Shared with $($app.SharedCount) role(s)`n"
        foreach ($role in $app.SharedWith) {
            $sharingText += "    * $role`n"
        }
        $sharingText += "`n"
    }
    
    $sharingText | Out-File -FilePath "$folder\SharingReport.txt" -Encoding UTF8
    
    # Create users and teams report
    $usersTeamsText = "USERS AND TEAMS ACCESS REPORT`n"
    $usersTeamsText += "=" * 80 + "`n`n"
    $usersTeamsText += "This report shows which users and teams have access to each app`n"
    $usersTeamsText += "organized by security roles.`n`n"
    $usersTeamsText += "=" * 80 + "`n`n"
    
    foreach ($app in ($analysisResults | Where-Object { $_.SharedCount -gt 0 } | Sort-Object AppName)) {
        $usersTeamsText += "APP: $($app.AppName)`n"
        $usersTeamsText += "-" * 80 + "`n"
        $usersTeamsText += "App ID: $($app.AppId)`n`n"
        
        foreach ($role in $app.SharedWith) {
            $usersTeamsText += "  ROLE: $role`n"
            
            # Users
            if ($app.UsersByRole.ContainsKey($role) -and $app.UsersByRole[$role].Count -gt 0) {
                $usersTeamsText += "    Users with this role ($($app.UsersByRole[$role].Count)):`n"
                foreach ($user in $app.UsersByRole[$role]) {
                    $usersTeamsText += "      * $user`n"
                }
            } else {
                $usersTeamsText += "    Users: None`n"
            }
            
            # Teams
            if ($app.TeamsByRole.ContainsKey($role) -and $app.TeamsByRole[$role].Count -gt 0) {
                $usersTeamsText += "    Teams with this role ($($app.TeamsByRole[$role].Count)):`n"
                foreach ($team in $app.TeamsByRole[$role]) {
                    $usersTeamsText += "      + $team`n"
                }
            } else {
                $usersTeamsText += "    Teams: None`n"
            }
            $usersTeamsText += "`n"
        }
        
        $usersTeamsText += "  SUMMARY:`n"
        $usersTeamsText += "    Total unique users: $($app.SharedWithUsers.Count)`n"
        $usersTeamsText += "    Total unique teams: $($app.SharedWithTeams.Count)`n"
        $usersTeamsText += "`n" + ("=" * 80) + "`n`n"
    }
    
    $usersTeamsText | Out-File -FilePath "$folder\UsersAndTeamsReport.txt" -Encoding UTF8
    
    # Create usage report
    $usageText = "APP USAGE REPORT`n"
    $usageText += "=" * 80 + "`n`n"
    
    $usageText += "Apps sorted by last activity (most recent first):`n`n"
    
    foreach ($app in ($analysisResults | Sort-Object { if($_.LastUsed){[DateTime]$_.LastUsed}else{[DateTime]::MinValue} } -Descending)) {
        if ($app.LastUsed) {
            $daysSince = ((Get-Date) - [DateTime]$app.LastUsed).Days
            $usageText += "- $($app.AppName)`n"
            $usageText += "  Last Activity: $($app.LastUsed) ($daysSince days ago)`n"
        } else {
            $usageText += "- $($app.AppName)`n"
            $usageText += "  Last Activity: Unknown`n"
        }
        $usageText += "`n"
    }
    
    $usageText | Out-File -FilePath "$folder\UsageReport.txt" -Encoding UTF8
    
    # Create CSV for Excel analysis
    $csvData = @()
    foreach ($app in $analysisResults) {
        $csvData += [PSCustomObject]@{
            AppName = $app.AppName
            AppId = $app.AppId
            UniqueName = $app.UniqueName
            State = $app.State
            IsOrphaned = $app.IsOrphaned
            SharedWithRoles = ($app.SharedWith -join "; ")
            SharedCount = $app.SharedCount
            TotalUsers = $app.SharedWithUsers.Count
            TotalTeams = $app.SharedWithTeams.Count
            UsersList = ($app.SharedWithUsers -join "; ")
            TeamsList = ($app.SharedWithTeams -join "; ")
            LastUsed = $app.LastUsed
            DaysSinceLastUse = if($app.LastUsed){((Get-Date) - [DateTime]$app.LastUsed).Days}else{"Unknown"}
            CreatedOn = $app.CreatedOn
            CreatedBy = $app.CreatedBy
            ModifiedOn = $app.ModifiedOn
            ModifiedBy = $app.ModifiedBy
        }
    }
    $csvData | Export-Csv -Path "$folder\AppAnalysis.csv" -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Files exported to: $folder" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  - CompleteAnalysis.json" -ForegroundColor Gray
    Write-Host "  - DetailedReport.txt (with users/teams per role)" -ForegroundColor Gray
    Write-Host "  - SharingReport.txt" -ForegroundColor Gray
    Write-Host "  - UsersAndTeamsReport.txt (NEW!)" -ForegroundColor Green
    Write-Host "  - UsageReport.txt" -ForegroundColor Gray
    Write-Host "  - AppAnalysis.csv (Excel - with users/teams)" -ForegroundColor Gray
    if ($orphanedApps -gt 0) {
        Write-Host "  - OrphanedApps.txt" -ForegroundColor Red
    }
    
    Write-Host "`nKey Findings:" -ForegroundColor Cyan
    if ($orphanedApps -gt 0) {
        Write-Host "  WARNING: $orphanedApps orphaned app(s) found!" -ForegroundColor Red
    } else {
        Write-Host "  No orphaned apps found." -ForegroundColor Green
    }
    Write-Host "  $notSharedApps app(s) are not shared with any security roles" -ForegroundColor Yellow
    Write-Host "  $sharedApps app(s) are shared with security roles" -ForegroundColor Green
    
}
catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
