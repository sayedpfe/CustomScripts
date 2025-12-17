# Setup Azure Automation Account for SharePoint Conditional Access
# This script creates all required Azure resources

param(
    [Parameter(Mandatory=$false)]
    [string]$ResourceGroupName = "rg-sharepoint-automation",
    
    [Parameter(Mandatory=$false)]
    [string]$Location = "EastUS",
    
    [Parameter(Mandatory=$false)]
    [string]$AutomationAccountName = "aa-spo-conditional-access",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Azure Automation Setup for SPO Conditional Access" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check if Azure PowerShell is installed
if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    Write-Host "`n❌ Azure PowerShell module not installed" -ForegroundColor Red
    Write-Host "Installing Az modules..." -ForegroundColor Yellow
    Install-Module -Name Az.Accounts, Az.Automation, Az.Resources -Force -AllowClobber -Scope CurrentUser
}

# Step 1: Connect to Azure
Write-Host "`n📌 Step 1: Connecting to Azure..." -ForegroundColor Cyan
try {
    $context = Get-AzContext
    if (-not $context) {
        Connect-AzAccount
    }
    
    if ($SubscriptionId) {
        Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    }
    
    $context = Get-AzContext
    Write-Host "   ✅ Connected to Azure" -ForegroundColor Green
    Write-Host "   Subscription: $($context.Subscription.Name)" -ForegroundColor White
    Write-Host "   Tenant: $($context.Tenant.Id)" -ForegroundColor White
} catch {
    Write-Host "   ❌ Failed to connect to Azure" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 2: Create Resource Group
Write-Host "`n📌 Step 2: Creating Resource Group..." -ForegroundColor Cyan
try {
    $rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    
    if ($rg) {
        Write-Host "   ℹ️  Resource Group already exists" -ForegroundColor Yellow
    } else {
        $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
        Write-Host "   ✅ Resource Group created: $ResourceGroupName" -ForegroundColor Green
    }
    
    Write-Host "   Location: $($rg.Location)" -ForegroundColor White
} catch {
    Write-Host "   ❌ Failed to create Resource Group" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 3: Create Automation Account
Write-Host "`n📌 Step 3: Creating Automation Account..." -ForegroundColor Cyan
try {
    $aa = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
    
    if ($aa) {
        Write-Host "   ℹ️  Automation Account already exists" -ForegroundColor Yellow
    } else {
        $aa = New-AzAutomationAccount `
            -ResourceGroupName $ResourceGroupName `
            -Name $AutomationAccountName `
            -Location $Location `
            -AssignSystemIdentity
        
        Write-Host "   ✅ Automation Account created: $AutomationAccountName" -ForegroundColor Green
    }
    
    # Ensure System-Assigned Managed Identity is enabled
    if (-not $aa.Identity.PrincipalId) {
        Write-Host "   Enabling System-Assigned Managed Identity..." -ForegroundColor Yellow
        Set-AzAutomationAccount `
            -ResourceGroupName $ResourceGroupName `
            -Name $AutomationAccountName `
            -AssignSystemIdentity | Out-Null
        
        $aa = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName
    }
    
    $principalId = $aa.Identity.PrincipalId
    Write-Host "   Managed Identity Principal ID: $principalId" -ForegroundColor White
    
} catch {
    Write-Host "   ❌ Failed to create Automation Account" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 4: Import Required PowerShell Modules
Write-Host "`n📌 Step 4: Importing PowerShell Modules..." -ForegroundColor Cyan

$modulesToImport = @(
    @{Name="PnP.PowerShell"; Version="latest"}
)

foreach ($module in $modulesToImport) {
    Write-Host "   Importing $($module.Name)..." -ForegroundColor Yellow
    try {
        # Check if module already exists
        $existingModule = Get-AzAutomationModule `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $module.Name `
            -ErrorAction SilentlyContinue
        
        if ($existingModule) {
            Write-Host "   ℹ️  Module $($module.Name) already imported (Status: $($existingModule.ProvisioningState))" -ForegroundColor Yellow
        } else {
            # Get the latest version from PowerShell Gallery
            $moduleUri = "https://www.powershellgallery.com/api/v2/package/$($module.Name)"
            
            New-AzAutomationModule `
                -ResourceGroupName $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -Name $module.Name `
                -ContentLinkUri $moduleUri | Out-Null
            
            Write-Host "   ✅ Module $($module.Name) import started (this may take several minutes)" -ForegroundColor Green
        }
    } catch {
        Write-Host "   ⚠️  Warning: Could not import module $($module.Name)" -ForegroundColor Yellow
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host "`n   Note: Module import runs in background. Check status in Azure Portal." -ForegroundColor Gray

# Step 5: Assign SharePoint Administrator Role to Managed Identity
Write-Host "`n📌 Step 5: Assigning SharePoint Administrator Role..." -ForegroundColor Cyan

# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
    Write-Host "   Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -AllowClobber -Scope CurrentUser
}

try {
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory" -NoWelcome
    
    # Get SharePoint Administrator role
    $role = Get-MgDirectoryRole -Filter "DisplayName eq 'SharePoint Administrator'"
    
    if (-not $role) {
        # Role template might not be activated, activate it
        $roleTemplate = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "SharePoint Administrator" }
        $role = New-MgDirectoryRole -RoleTemplateId $roleTemplate.Id
        Write-Host "   Activated SharePoint Administrator role" -ForegroundColor Yellow
    }
    
    # Check if already assigned
    $existingMember = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id | Where-Object { $_.Id -eq $principalId }
    
    if ($existingMember) {
        Write-Host "   ℹ️  SharePoint Administrator role already assigned" -ForegroundColor Yellow
    } else {
        # Assign the role
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$principalId"
        }
        
        Write-Host "   ✅ SharePoint Administrator role assigned to Managed Identity" -ForegroundColor Green
    }
    
    Disconnect-MgGraph | Out-Null
    
} catch {
    Write-Host "   ⚠️  Warning: Could not assign SharePoint Administrator role" -ForegroundColor Yellow
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "`n   ⚠️  MANUAL STEP REQUIRED:" -ForegroundColor Red
    Write-Host "   1. Go to Azure Portal → Entra ID → Roles and administrators" -ForegroundColor White
    Write-Host "   2. Find 'SharePoint Administrator' role" -ForegroundColor White
    Write-Host "   3. Click '+ Add assignments'" -ForegroundColor White
    Write-Host "   4. Search for '$AutomationAccountName'" -ForegroundColor White
    Write-Host "   5. Select the managed identity and assign the role" -ForegroundColor White
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ SETUP COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host @"

Azure Automation Account Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Resource Group:       $ResourceGroupName
Automation Account:   $AutomationAccountName
Location:             $Location
Managed Identity ID:  $principalId

Next Steps:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. ✅ Wait for PnP.PowerShell module to finish importing (5-10 minutes)
   Check status: Azure Portal → Automation Account → Modules

2. ⏭️  Run: .\2-Create-Runbook.ps1
   This will create the PowerShell runbook that applies conditional access

3. ⏭️  Run: .\3-Setup-Webhook.ps1
   This will create a webhook URL for Power Automate

4. ⏭️  Import the Power Automate flow using the webhook URL

"@ -ForegroundColor White

# Save configuration for next scripts
$config = @{
    ResourceGroupName = $ResourceGroupName
    AutomationAccountName = $AutomationAccountName
    Location = $Location
    PrincipalId = $principalId
    SubscriptionId = $context.Subscription.Id
    TenantId = $context.Tenant.Id
}

$config | ConvertTo-Json | Out-File -FilePath ".\automation-config.json" -Encoding UTF8
Write-Host "Configuration saved to: automation-config.json" -ForegroundColor Gray
