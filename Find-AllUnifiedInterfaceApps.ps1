# Script to find all Unified Interface apps (clienttype = 4)

param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentName = $null
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Searching for Unified Interface Apps" -ForegroundColor Cyan
Write-Host "Client Type = 4" -ForegroundColor Cyan
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
    Write-Host "URL: $envUrl" -ForegroundColor Gray
    
    # Connect to Dataverse
    Write-Host "`nConnecting to Dataverse..." -ForegroundColor Cyan
    $conn = Connect-CrmOnline -ServerUrl $envUrl -ForceOAuth
    
    if (!$conn -or !$conn.IsReady) {
        Write-Host "ERROR: Failed to connect to Dataverse" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    
    # Search for all Unified Interface apps
    Write-Host "`nSearching for Unified Interface apps (clienttype = 4)..." -ForegroundColor Cyan
    
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="appmodule">
    <all-attributes />
    <filter type="and">
      <condition attribute="clienttype" operator="eq" value="4" />
    </filter>
    <order attribute="name" descending="false" />
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml
    
    $appCount = $result.CrmRecords.Count
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "FOUND $appCount UNIFIED INTERFACE APPS" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    if ($appCount -eq 0) {
        Write-Host "`nNo Unified Interface apps found in the default environment." -ForegroundColor Yellow
        exit 0
    }
    
    # Display summary
    Write-Host "`nApp List:" -ForegroundColor Cyan
    $counter = 1
    foreach ($app in $result.CrmRecords) {
        Write-Host "`n[$counter] " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.name)" -ForegroundColor White
        Write-Host "    ID: " -ForegroundColor Gray -NoNewline
        Write-Host "$($app.appmoduleid)" -ForegroundColor Cyan
        Write-Host "    Unique Name: " -ForegroundColor Gray -NoNewline
        Write-Host "$($app.uniquename)" -ForegroundColor White
        Write-Host "    State: " -ForegroundColor Gray -NoNewline
        Write-Host "$($app.statecode)" -ForegroundColor White
        Write-Host "    Created: " -ForegroundColor Gray -NoNewline
        Write-Host "$($app.createdon)" -ForegroundColor White
        Write-Host "    Modified: " -ForegroundColor Gray -NoNewline
        Write-Host "$($app.modifiedon)" -ForegroundColor White
        $counter++
    }
    
    # Export results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $folder = "UnifiedInterfaceApps_$timestamp"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    
    # Save all apps to JSON
    $allAppsData = @()
    foreach ($app in $result.CrmRecords) {
        $allAppsData += $app.original
    }
    $allAppsData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folder\AllApps_Full.json" -Encoding UTF8
    
    # Save summary
    $summaryText = "Unified Interface Apps - Summary`n"
    $summaryText += "=====================`n`n"
    $summaryText += "Search Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $summaryText += "Environment: $($targetEnv.DisplayName)`n"
    $summaryText += "URL: $envUrl`n"
    $summaryText += "Total Apps Found: $appCount`n`n"
    $summaryText += "Filter: clienttype = 4 (Unified Interface)`n`n"
    $summaryText += "Apps List:`n"
    $summaryText += "==========`n`n"
    
    $counter = 1
    foreach ($app in $result.CrmRecords) {
        $summaryText += "[$counter] $($app.name)`n"
        $summaryText += "    ID: $($app.appmoduleid)`n"
        $summaryText += "    Unique Name: $($app.uniquename)`n"
        $summaryText += "    State: $($app.statecode)`n"
        $summaryText += "    Status: $($app.statuscode)`n"
        $summaryText += "    Is Managed: $($app.ismanaged)`n"
        $summaryText += "    Created: $($app.createdon)`n"
        $summaryText += "    Created By: $($app.createdby)`n"
        $summaryText += "    Modified: $($app.modifiedon)`n"
        $summaryText += "    Modified By: $($app.modifiedby)`n"
        if ($app.description) {
            $summaryText += "    Description: $($app.description)`n"
        }
        $summaryText += "    URL: ${envUrl}main.aspx?appid=$($app.appmoduleid)`n"
        $summaryText += "`n"
        $counter++
    }
    
    $summaryText | Out-File -FilePath "$folder\AppsSummary.txt" -Encoding UTF8
    
    # Save individual app files
    $individualFolder = "$folder\IndividualApps"
    New-Item -ItemType Directory -Path $individualFolder -Force | Out-Null
    
    foreach ($app in $result.CrmRecords) {
        $appName = $app.name -replace '[^\w\s-]', '' -replace '\s+', '_'
        $appFileName = "${appName}_$($app.appmoduleid)"
        
        # JSON
        $app.original | ConvertTo-Json -Depth 10 | Out-File -FilePath "$individualFolder\${appFileName}.json" -Encoding UTF8
        
        # Readable text
        $app.original | Format-List | Out-File -FilePath "$individualFolder\${appFileName}.txt" -Encoding UTF8
    }
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Files exported to: $folder" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  - AllApps_Full.json (all apps)" -ForegroundColor Gray
    Write-Host "  - AppsSummary.txt (summary)" -ForegroundColor Gray
    Write-Host "  - IndividualApps\ (folder with individual app files)" -ForegroundColor Gray
    Write-Host "    └─ $appCount JSON files" -ForegroundColor Gray
    Write-Host "    └─ $appCount TXT files" -ForegroundColor Gray
    
}
catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
