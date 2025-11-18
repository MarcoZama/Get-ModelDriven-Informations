# Setup Script - Deploy Application User to All Environments
# Uses PAC CLI to authenticate and deploy application user

# Load credentials from environment variables
$applicationId = $env:PP_APP_ID
$clientSecret = $env:PP_CLIENT_SECRET
$tenantId = $env:PP_TENANT_ID

if (!$applicationId -or !$clientSecret -or !$tenantId) {
    Write-Host "ERROR: Missing required environment variables!" -ForegroundColor Red
    Write-Host "Please set the following environment variables:" -ForegroundColor Yellow
    Write-Host "  PP_APP_ID - Application (client) ID" -ForegroundColor Yellow
    Write-Host "  PP_CLIENT_SECRET - Client secret value" -ForegroundColor Yellow
    Write-Host "  PP_TENANT_ID - Tenant ID" -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Yellow
    Write-Host "Example:" -ForegroundColor Cyan
    Write-Host '  $env:PP_APP_ID = "your-app-id"' -ForegroundColor Gray
    Write-Host '  $env:PP_CLIENT_SECRET = "your-secret"' -ForegroundColor Gray
    Write-Host '  $env:PP_TENANT_ID = "your-tenant-id"' -ForegroundColor Gray
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Application User Deployment" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Authenticate with Interactive Login (needed for admin operations)
Write-Host "Step 1: Authenticating to Power Platform..." -ForegroundColor Yellow
Write-Host "  This will open a browser for interactive login" -ForegroundColor Gray
Write-Host "  (Required for admin operations)" -ForegroundColor Gray
Write-Host ""

pac auth create

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: Authentication failed!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Authentication successful!" -ForegroundColor Green
Write-Host ""

# Step 2: Save credentials for scanner
Write-Host "Step 2: Saving credentials for scanner..." -ForegroundColor Yellow

$secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force

$credObject = [PSCustomObject]@{
    ApplicationId = $applicationId
    TenantId = $tenantId
    ClientSecret = $secureSecret
}

$credObject | Export-Clixml -Path "sp-credentials.xml"
Write-Host "  Credentials saved to: sp-credentials.xml" -ForegroundColor Green
Write-Host ""

# Step 3: Get all environments
Write-Host "Step 3: Retrieving environments..." -ForegroundColor Yellow

$pacOutput = pac admin list 2>&1 | Out-String
$lines = $pacOutput -split "`n"

$environments = @()
$headerPassed = $false

foreach ($line in $lines) {
    if ($line -match "Active Environment") {
        $headerPassed = $true
        continue
    }
    
    if (!$headerPassed) { continue }
    
    if ($line.Trim() -match "^\*?\s*(.+?)\s+([a-f0-9\-]{36})\s+(https://[^\s]+)\s+") {
        $envName = $matches[1].Trim()
        $envId = $matches[2].Trim()
        
        if ($envName -and $envName -ne "Active" -and $envName -ne "Environment") {
            $environments += [PSCustomObject]@{
                DisplayName = $envName
                EnvironmentId = $envId
            }
        }
    }
}

Write-Host "  Found $($environments.Count) environments" -ForegroundColor Green
Write-Host ""

# Step 4: Deploy application user to each environment
Write-Host "Step 4: Deploying application user to all environments..." -ForegroundColor Yellow
Write-Host ""

$successCount = 0
$alreadyExistsCount = 0
$failCount = 0

foreach ($env in $environments) {
    $envName = $env.DisplayName
    $envId = $env.EnvironmentId
    
    Write-Host "  [$($successCount + $alreadyExistsCount + $failCount + 1)/$($environments.Count)] $envName" -ForegroundColor Cyan
    
    try {
        # Assign application user with System Administrator role
        Write-Host "    - Assigning application user with System Administrator role..." -ForegroundColor Gray
        $assignOutput = pac admin assign-user --environment $envId --user $applicationId --role "System Administrator" --application-user 2>&1 | Out-String
        
        if ($LASTEXITCODE -eq 0 -or $assignOutput -like "*Successfully assigned*") {
            Write-Host "    - SUCCESS: Application user assigned" -ForegroundColor Green
            $successCount++
        } elseif ($assignOutput -like "*already has*" -or $assignOutput -like "*already exists*") {
            Write-Host "    - Application user already exists with role" -ForegroundColor Yellow
            $alreadyExistsCount++
        } else {
            Write-Host "    - FAILED: $assignOutput" -ForegroundColor Red
            $failCount++
        }
        
    } catch {
        Write-Host "    - ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $failCount++
    }
    
    Write-Host ""
}

# Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Deployment Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  Total Environments: $($environments.Count)" -ForegroundColor White
Write-Host "  Successfully Deployed: $successCount" -ForegroundColor Green
Write-Host "  Already Existed: $alreadyExistsCount" -ForegroundColor Yellow
Write-Host "  Failed: $failCount" -ForegroundColor Red
Write-Host ""
Write-Host "Next step: Run the scanner" -ForegroundColor Yellow
Write-Host "  .\2-Scan-AllEnvironments.ps1" -ForegroundColor White
Write-Host ""
