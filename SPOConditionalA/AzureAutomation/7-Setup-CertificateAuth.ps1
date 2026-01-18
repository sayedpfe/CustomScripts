# Setup Certificate-Based Authentication for SharePoint
# This creates an Azure AD App Registration with certificate auth

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\automation-config.json",
    
    [Parameter(Mandatory=$false)]
    [string]$AppDisplayName = "SPO-ConditionalAccess-Automation",
    
    [Parameter(Mandatory=$false)]
    [int]$CertificateValidityYears = 2
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Certificate-Based Authentication Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    Write-Host "❌ Configuration file not found: $ConfigPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content $ConfigPath | ConvertFrom-Json
Write-Host "Loaded configuration" -ForegroundColor Gray

# Connect to Azure
Write-Host "`n📌 Step 1: Connecting to Azure..." -ForegroundColor Cyan
Connect-AzAccount -SubscriptionId $config.SubscriptionId | Out-Null
Write-Host "✅ Connected" -ForegroundColor Green

# Connect to Microsoft Graph
Write-Host "`n📌 Step 2: Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -NoWelcome
Write-Host "✅ Connected" -ForegroundColor Green

# Create self-signed certificate
Write-Host "`n📌 Step 3: Creating self-signed certificate..." -ForegroundColor Cyan

$certName = "CN=$AppDisplayName"
$certPath = Join-Path $env:TEMP "$AppDisplayName.pfx"
$certPassword = -join ((65..90) + (97..122) + (48..57) | Get-Random -Count 32 | ForEach-Object {[char]$_})
$securePassword = ConvertTo-SecureString -String $certPassword -Force -AsPlainText

$cert = New-SelfSignedCertificate `
    -Subject $certName `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($CertificateValidityYears)

Write-Host "✅ Certificate created" -ForegroundColor Green
Write-Host "   Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
Write-Host "   Valid until: $($cert.NotAfter)" -ForegroundColor White

# Export certificate
Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $securePassword | Out-Null
$cerBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
$cerBase64 = [System.Convert]::ToBase64String($cerBytes)

Write-Host "✅ Certificate exported" -ForegroundColor Green

# Create or update Azure AD App Registration
Write-Host "`n📌 Step 4: Creating Azure AD App Registration..." -ForegroundColor Cyan

$existingApp = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue

if ($existingApp) {
    Write-Host "   App already exists, updating..." -ForegroundColor Yellow
    $app = $existingApp
    $appId = $app.AppId
} else {
    # Create new app
    $appParams = @{
        DisplayName = $AppDisplayName
        SignInAudience = "AzureADMyOrg"
        KeyCredentials = @(
            @{
                Type = "AsymmetricX509Cert"
                Usage = "Verify"
                Key = [System.Text.Encoding]::ASCII.GetBytes($cerBase64)
            }
        )
    }
    
    $app = New-MgApplication @appParams
    $appId = $app.AppId
    
    Write-Host "✅ App Registration created" -ForegroundColor Green
    Write-Host "   App ID: $appId" -ForegroundColor White
}

# Add certificate to app if not already there
if (-not $existingApp) {
    Write-Host "   Certificate already added during creation" -ForegroundColor Gray
} else {
    Write-Host "   Adding certificate to existing app..." -ForegroundColor Yellow
    $keyCredential = @{
        Type = "AsymmetricX509Cert"
        Usage = "Verify"
        Key = [System.Text.Encoding]::ASCII.GetBytes($cerBase64)
        DisplayName = "Automation Certificate"
    }
    
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials @($keyCredential)
    Write-Host "✅ Certificate added" -ForegroundColor Green
}

# Create Service Principal if doesn't exist
Write-Host "`n📌 Step 5: Creating Service Principal..." -ForegroundColor Cyan
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue

if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $appId
    Write-Host "✅ Service Principal created" -ForegroundColor Green
} else {
    Write-Host "   Service Principal already exists" -ForegroundColor Yellow
}

Write-Host "   Service Principal ID: $($sp.Id)" -ForegroundColor White

# Grant SharePoint permissions (Sites.FullControl.All)
Write-Host "`n📌 Step 6: Granting SharePoint permissions..." -ForegroundColor Cyan

$sharepointSP = Get-MgServicePrincipal -Filter "displayName eq 'Office 365 SharePoint Online'" -ErrorAction SilentlyContinue
if ($sharepointSP) {
    $permission = $sharepointSP.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" }
    
    if ($permission) {
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $sp.Id `
                -PrincipalId $sp.Id `
                -ResourceId $sharepointSP.Id `
                -AppRoleId $permission.Id | Out-Null
            
            Write-Host "✅ Sites.FullControl.All permission granted" -ForegroundColor Green
        } catch {
            if ($_.Exception.Message -like "*already exists*") {
                Write-Host "   Permission already granted" -ForegroundColor Yellow
            } else {
                Write-Host "   ⚠️ Failed to grant permission: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

# Assign SharePoint Administrator role
Write-Host "`n📌 Step 7: Assigning SharePoint Administrator role..." -ForegroundColor Cyan

$role = Get-MgDirectoryRole -Filter "DisplayName eq 'SharePoint Administrator'"
if (-not $role) {
    $roleTemplate = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "SharePoint Administrator" }
    $role = New-MgDirectoryRole -RoleTemplateId $roleTemplate.Id
}

try {
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"
    } | Out-Null
    Write-Host "✅ SharePoint Administrator role assigned" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exist*") {
        Write-Host "   Role already assigned" -ForegroundColor Yellow
    } else {
        Write-Host "   ⚠️ Failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Upload certificate to Automation Account
Write-Host "`n📌 Step 8: Uploading certificate to Automation Account..." -ForegroundColor Cyan

try {
    New-AzAutomationCertificate `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name "SPOAppCertificate" `
        -Path $certPath `
        -Password $securePassword `
        -Exportable | Out-Null
    
    Write-Host "✅ Certificate uploaded as 'SPOAppCertificate'" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "   Certificate already exists, updating..." -ForegroundColor Yellow
        Remove-AzAutomationCertificate `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name "SPOAppCertificate" `
            -Force
        
        New-AzAutomationCertificate `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name "SPOAppCertificate" `
            -Path $certPath `
            -Password $securePassword `
            -Exportable | Out-Null
        
        Write-Host "✅ Certificate updated" -ForegroundColor Green
    } else {
        throw
    }
}

# Store App ID as Automation Variable
Write-Host "`n📌 Step 9: Storing App ID in Automation Account..." -ForegroundColor Cyan

try {
    New-AzAutomationVariable `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name "SPOAppId" `
        -Value $appId `
        -Encrypted $false | Out-Null
    
    Write-Host "✅ App ID stored as 'SPOAppId'" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "   Variable already exists, updating..." -ForegroundColor Yellow
        Set-AzAutomationVariable `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name "SPOAppId" `
            -Value $appId `
            -Encrypted $false | Out-Null
        
        Write-Host "✅ App ID updated" -ForegroundColor Green
    } else {
        throw
    }
}

# Clean up local certificate
Remove-Item $certPath -Force -ErrorAction SilentlyContinue
Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force -ErrorAction SilentlyContinue

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ SETUP COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host @"

Certificate-Based Authentication configured:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
App Registration:     $AppDisplayName
App ID (Client ID):   $appId
Service Principal:    $($sp.Id)
Certificate Name:     SPOAppCertificate
Certificate Expires:  $($cert.NotAfter)
Automation Variable:  SPOAppId = $appId

Permissions Granted:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Sites.FullControl.All (SharePoint API)
✅ SharePoint Administrator (Directory Role)

Next Steps:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. ⏭️  Wait for Microsoft.Online.SharePoint.PowerShell module to finish importing
   Check: Azure Portal → Automation Account → Modules

2. ⏭️  Run: .\6-Update-Runbook-SPO.ps1
   This updates the runbook to use certificate authentication

3. ⏭️  Test with: .\4-Test-Webhook.ps1

"@ -ForegroundColor White

Disconnect-MgGraph | Out-Null
