# Troubleshooting Guide

## HttpRequestException - Browser Login Non Si Apre

Se ricevi errore `HttpRequestException` e la finestra del browser non si apre durante `pac auth create`, prova queste soluzioni:

### Soluzione 1: Device Code Flow (Consigliata)

Usa il device code flow che non richiede il browser:

```powershell
# Invece di eseguire .\1-Setup.ps1 direttamente, usa:
pac auth create --deviceCode
```

Questo mostrerà:
1. Un codice da copiare (es: `A1B2C3D4`)
2. Un link da aprire manualmente in un browser (anche su un altro dispositivo)
3. Inserisci il codice nel browser
4. Completa il login

Dopo l'autenticazione con device code, puoi continuare con lo script:

```powershell
# Imposta le variabili d'ambiente
$env:PP_APP_ID = "your-app-id"
$env:PP_CLIENT_SECRET = "your-secret"
$env:PP_TENANT_ID = "your-tenant-id"

# Salta l'autenticazione (già fatto) ed esegui solo il deployment
# Vedi sotto "Script Alternativo"
```

### Soluzione 2: Usa Browser Specifico

Forza PAC CLI ad usare un browser specifico:

```powershell
# Imposta browser di default
$env:BROWSER = "chrome"  # oppure "msedge", "firefox"

# Poi esegui normalmente
.\1-Setup.ps1
```

### Soluzione 3: Autenticazione Manuale

Se le soluzioni precedenti non funzionano, autentica manualmente e salta lo step 1:

**Step 1 - Autentica manualmente:**
```powershell
# Prova con device code
pac auth create --deviceCode --cloud Public --tenant "your-tenant-id"

# OPPURE prova a specificare il browser
pac auth create --cloud Public --tenant "your-tenant-id"
```

**Step 2 - Esegui setup manuale:**

Crea un file `1-Setup-Manual.ps1`:

```powershell
# Setup manuale (assumendo che pac auth create è già stato fatto)

# Carica credenziali
$applicationId = $env:PP_APP_ID
$clientSecret = $env:PP_CLIENT_SECRET
$tenantId = $env:PP_TENANT_ID

if (!$applicationId -or !$clientSecret -or !$tenantId) {
    Write-Host "ERROR: Imposta le variabili d'ambiente!" -ForegroundColor Red
    exit 1
}

Write-Host "Saving credentials..." -ForegroundColor Yellow
$secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$credObject = [PSCustomObject]@{
    ApplicationId = $applicationId
    TenantId = $tenantId
    ClientSecret = $secureSecret
}
$credObject | Export-Clixml -Path "sp-credentials.xml"
Write-Host "Credentials saved to: sp-credentials.xml" -ForegroundColor Green

Write-Host "Retrieving environments..." -ForegroundColor Yellow
$pacOutput = pac admin list | Out-String
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

Write-Host "Found $($environments.Count) environments" -ForegroundColor Green

Write-Host "Deploying application user to all environments..." -ForegroundColor Yellow

$successCount = 0
$alreadyExistsCount = 0
$failCount = 0

foreach ($env in $environments) {
    $envName = $env.DisplayName
    $envId = $env.EnvironmentId
    
    Write-Host "  [$($successCount + $alreadyExistsCount + $failCount + 1)/$($environments.Count)] $envName" -ForegroundColor Cyan
    
    try {
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
```

Poi esegui:
```powershell
.\1-Setup-Manual.ps1
```

### Soluzione 4: Controlla Proxy/Firewall

Se sei dietro un proxy aziendale:

```powershell
# Configura proxy per PAC CLI
$env:HTTP_PROXY = "http://proxy-server:port"
$env:HTTPS_PROXY = "http://proxy-server:port"

# Poi esegui normalmente
.\1-Setup.ps1
```

### Soluzione 5: Aggiorna PAC CLI

Assicurati di avere l'ultima versione:

```powershell
# Controlla versione
pac --version

# Aggiorna
dotnet tool update --global Microsoft.PowerApps.CLI.Tool
```

## Errore Comune: "Connection Failed"

Se dopo l'autenticazione vedi "Connection Failed" per alcuni environment:
- L'application user NON è stato creato in quell'environment
- Riprova il deployment con lo script setup

## Verificare Autenticazione Riuscita

Dopo `pac auth create` (con qualsiasi metodo), verifica:

```powershell
# Controlla autenticazione attiva
pac auth list

# Dovresti vedere il tuo account con asterisco (*)
```

## Contatto

Se nessuna soluzione funziona, contatta l'amministratore tenant per verificare:
- Conditional Access policies
- MFA requirements
- Browser restrictions
- Network policies
