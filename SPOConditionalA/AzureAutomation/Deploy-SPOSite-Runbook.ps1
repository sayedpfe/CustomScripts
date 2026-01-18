# Deploy Set-SPOSite Runbook to Azure Automation
# This script sets up the runbook that uses Microsoft.Online.SharePoint.PowerShell

param(
    [Parameter(Mandatory=$false)]
    [string]$AdminUsername,
    
    [Parameter(Mandatory=$false)]
    [SecureString]$AdminPassword
)

$ErrorActionPreference = "Stop"

Write-Host "`n🚀 Deploying Set-SPOSite Runbook to Azure Automation" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Load configuration
$configPath = Join-Path $PSScriptRoot "automation-config.json"
if (-not (Test-Path $configPath)) {
    Write-Host "❌ Configuration file not found: $configPath" -ForegroundColor Red
    exit
}

$config = Get-Content $configPath | ConvertFrom-Json

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  Resource Group: $($config.ResourceGroupName)" -ForegroundColor White
Write-Host "  Automation Account: $($config.AutomationAccountName)" -ForegroundColor White
Write-Host "  App ID: $($config.AppId)" -ForegroundColor White
Write-Host "  Tenant ID: $($config.TenantId)`n" -ForegroundColor White

# Step 1: Create Automation Variable for tenant name
Write-Host "📦 Step 1: Creating Automation Variable..." -ForegroundColor Yellow

$tenantName = "M365CPI90282478"
try {
    $existingVar = Get-AzAutomationVariable -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'SPOTenantName' -ErrorAction SilentlyContinue
    
    if ($existingVar) {
        Write-Host "   Variable 'SPOTenantName' already exists with value: $($existingVar.Value)" -ForegroundColor Green
    } else {
        New-AzAutomationVariable -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'SPOTenantName' -Value $tenantName -Encrypted $false | Out-Null
        Write-Host "   ✅ Created variable 'SPOTenantName' = $tenantName" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to create variable: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 2: Create Automation Credential
Write-Host "`n📦 Step 2: Creating Automation Credential..." -ForegroundColor Yellow

if (-not $AdminUsername -or -not $AdminPassword) {
    Write-Host "   ⚠️ Admin credentials not provided as parameters" -ForegroundColor Yellow
    Write-Host "   Please enter SharePoint Admin credentials:" -ForegroundColor Yellow
    $AdminUsername = Read-Host "   SharePoint Admin Username (e.g., admin@tenant.onmicrosoft.com)"
    $AdminPassword = Read-Host "   SharePoint Admin Password" -AsSecureString
}

try {
    $existingCred = Get-AzAutomationCredential -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'SPOAdminCredential' -ErrorAction SilentlyContinue
    
    if ($existingCred) {
        Write-Host "   Credential 'SPOAdminCredential' already exists for: $($existingCred.UserName)" -ForegroundColor Yellow
        $update = Read-Host "   Do you want to update it? (y/n)"
        if ($update -eq 'y') {
            $credential = New-Object System.Management.Automation.PSCredential($AdminUsername, $AdminPassword)
            Set-AzAutomationCredential -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'SPOAdminCredential' -Value $credential | Out-Null
            Write-Host "   ✅ Updated credential 'SPOAdminCredential'" -ForegroundColor Green
        }
    } else {
        $credential = New-Object System.Management.Automation.PSCredential($AdminUsername, $AdminPassword)
        New-AzAutomationCredential -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'SPOAdminCredential' -Value $credential | Out-Null
        Write-Host "   ✅ Created credential 'SPOAdminCredential' for: $AdminUsername" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to create credential: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 3: Import Microsoft.Online.SharePoint.PowerShell module
Write-Host "`n📦 Step 3: Importing Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow

try {
    $existingModule = Get-AzAutomationModule -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
    
    if ($existingModule -and $existingModule.ProvisioningState -eq 'Succeeded') {
        Write-Host "   ✅ Module already imported (version $($existingModule.Version))" -ForegroundColor Green
    } else {
        Write-Host "   Importing module from PowerShell Gallery..." -ForegroundColor Gray
        $moduleUri = "https://www.powershellgallery.com/api/v2/package/Microsoft.Online.SharePoint.PowerShell"
        New-AzAutomationModule -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name 'Microsoft.Online.SharePoint.PowerShell' -ContentLinkUri $moduleUri | Out-Null
        Write-Host "   ✅ Module import started (this may take several minutes)" -ForegroundColor Green
        Write-Host "   ℹ️  Check Azure Portal to confirm import completion before testing runbook" -ForegroundColor Cyan
    }
} catch {
    Write-Host "   ❌ Failed to import module: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   You may need to import this module manually in the Azure Portal" -ForegroundColor Yellow
}

# Step 4: Import and publish the runbook
Write-Host "`n📦 Step 4: Publishing runbook..." -ForegroundColor Yellow

$runbookPath = Join-Path $PSScriptRoot "ApplyConditionalAccessPolicy-SPO.ps1"
$runbookName = "ApplyConditionalAccessPolicy-SPO"

if (-not (Test-Path $runbookPath)) {
    Write-Host "   ❌ Runbook file not found: $runbookPath" -ForegroundColor Red
    exit
}

try {
    # Import or update runbook
    $existingRunbook = Get-AzAutomationRunbook -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Name $runbookName -ErrorAction SilentlyContinue
    
    if ($existingRunbook) {
        Write-Host "   Updating existing runbook..." -ForegroundColor Gray
        Import-AzAutomationRunbook -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Path $runbookPath -Name $runbookName -Type PowerShell -Force -Published | Out-Null
        Write-Host "   ✅ Runbook updated and published" -ForegroundColor Green
    } else {
        Write-Host "   Creating new runbook..." -ForegroundColor Gray
        Import-AzAutomationRunbook -ResourceGroupName $config.ResourceGroupName -AutomationAccountName $config.AutomationAccountName -Path $runbookPath -Name $runbookName -Type PowerShell -Published | Out-Null
        Write-Host "   ✅ Runbook created and published" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to publish runbook: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

Write-Host "`n═══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "✅ Deployment completed successfully!" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════`n" -ForegroundColor Cyan

Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Wait for module import to complete (check Azure Portal)" -ForegroundColor White
Write-Host "2. Test the runbook with:" -ForegroundColor White
Write-Host "   `$params = @{" -ForegroundColor Gray
Write-Host "       SiteUrl = 'https://m365cpi90282478.sharepoint.com/sites/DeutschKurs'" -ForegroundColor Gray
Write-Host "       AuthenticationContextName = 'c1'" -ForegroundColor Gray
Write-Host "   }" -ForegroundColor Gray
Write-Host "   Start-AzAutomationRunbook -ResourceGroupName '$($config.ResourceGroupName)' ```" -ForegroundColor Gray
Write-Host "       -AutomationAccountName '$($config.AutomationAccountName)' ```" -ForegroundColor Gray
Write-Host "       -Name '$runbookName' -Parameters `$params" -ForegroundColor Gray
Write-Host "3. Create/update webhook for Power Automate integration`n" -ForegroundColor White
