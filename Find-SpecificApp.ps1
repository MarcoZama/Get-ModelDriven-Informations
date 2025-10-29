# Script to find specific app by ID

param(
    [Parameter(Mandatory=$false)]
    [string]$AppId = "28d151a9-08e6-ee11-904c-000d3a26cee9",
    
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentName = $null
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Searching for App ID: $AppId" -ForegroundColor Cyan
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
    
    # Search for the app
    Write-Host "`nSearching for app..." -ForegroundColor Cyan
    
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="appmodule">
    <all-attributes />
    <filter type="and">
      <condition attribute="appmoduleid" operator="eq" value="$AppId" />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -eq 0) {
        Write-Host "`n========================================" -ForegroundColor Red
        Write-Host "APP NOT FOUND" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "`nThe app with ID '$AppId' was not found in the default environment." -ForegroundColor Yellow
        exit 1
    }
    
    $app = $result.CrmRecords[0]
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "APP FOUND!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    # Display app details
    Write-Host "`nBasic Information:" -ForegroundColor Cyan
    Write-Host "  Name: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($app.name)" -ForegroundColor White
    
    if ($app.uniquename) {
        Write-Host "  Unique Name: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.uniquename)" -ForegroundColor White
    }
    
    if ($app.description) {
        Write-Host "  Description: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.description)" -ForegroundColor White
    }
    
    Write-Host "`nIdentifiers:" -ForegroundColor Cyan
    Write-Host "  App ID: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($app.appmoduleid)" -ForegroundColor White
    
    if ($app.appmoduleidunique) {
        Write-Host "  App ID Unique: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.appmoduleidunique)" -ForegroundColor White
    }
    
    Write-Host "`nConfiguration:" -ForegroundColor Cyan
    Write-Host "  Client Type: " -ForegroundColor Yellow -NoNewline
    $clientTypeDesc = switch ([int]$app.clienttype) {
        1 { "Web" }
        2 { "Outlook" }
        3 { "Mobile" }
        4 { "Unified Interface" }
        8 { "Teams" }
        default { "Unknown" }
    }
    Write-Host "$clientTypeDesc ($($app.clienttype))" -ForegroundColor White
    
    Write-Host "  State: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($app.statecode)" -ForegroundColor White
    
    Write-Host "  Status: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($app.statuscode)" -ForegroundColor White
    
    Write-Host "  Is Managed: " -ForegroundColor Yellow -NoNewline
    Write-Host "$($app.ismanaged)" -ForegroundColor White
    
    Write-Host "`nDates:" -ForegroundColor Cyan
    if ($app.createdon) {
        Write-Host "  Created On: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.createdon)" -ForegroundColor White
    }
    
    if ($app.modifiedon) {
        Write-Host "  Modified On: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.modifiedon)" -ForegroundColor White
    }
    
    if ($app.publishedon) {
        Write-Host "  Published On: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($app.publishedon)" -ForegroundColor White
    }
    
    Write-Host "`nAccess URL:" -ForegroundColor Cyan
    Write-Host "  ${envUrl}main.aspx?appid=$AppId" -ForegroundColor Green
    
    # Export details
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $appName = $app.name -replace '[^\w\s-]', '' -replace '\s+', '_'
    $folder = "App_${appName}_$timestamp"
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    
    # Save full JSON
    $app.original | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folder\AppDetails_Full.json" -Encoding UTF8
    
    # Save summary
    $summaryText = "App Details - Summary`n"
    $summaryText += "=====================`n`n"
    $summaryText += "Search Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $summaryText += "App ID: $AppId`n`n"
    $summaryText += "Environment:`n"
    $summaryText += "  Name: $($targetEnv.DisplayName)`n"
    $summaryText += "  URL: $envUrl`n"
    $summaryText += "  Environment ID: $($targetEnv.EnvironmentName)`n`n"
    $summaryText += "App Information:`n"
    $summaryText += "  Name: $($app.name)`n"
    $summaryText += "  Unique Name: $($app.uniquename)`n"
    $summaryText += "  Description: $($app.description)`n"
    $summaryText += "  Client Type: $clientTypeDesc ($($app.clienttype))`n"
    $summaryText += "  State: $($app.statecode)`n"
    $summaryText += "  Status: $($app.statuscode)`n"
    $summaryText += "  Is Managed: $($app.ismanaged)`n"
    $summaryText += "  Created: $($app.createdon)`n"
    $summaryText += "  Modified: $($app.modifiedon)`n"
    $summaryText += "  Published: $($app.publishedon)`n`n"
    $summaryText += "Access URL:`n"
    $summaryText += "${envUrl}main.aspx?appid=$AppId`n"
    
    $summaryText | Out-File -FilePath "$folder\AppSummary.txt" -Encoding UTF8
    
    # Save readable text format
    $app.original | Format-List | Out-File -FilePath "$folder\AppDetails_Readable.txt" -Encoding UTF8
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Files exported to: $folder" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  - AppDetails_Full.json" -ForegroundColor Gray
    Write-Host "  - AppSummary.txt" -ForegroundColor Gray
    Write-Host "  - AppDetails_Readable.txt" -ForegroundColor Gray
    
}
catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
