# Local test for PnP.PowerShell certificate authentication
# This will test the same approach we'll use in Azure Automation

$ErrorActionPreference = "Stop"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PnP.PowerShell Certificate Auth Test" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Configuration
$appId = "7bf03e05-9069-412e-a27c-f2f6326fc69f"
$tenantId = "b22f8675-8375-455b-941a-67bee4cf7747"
$certThumbprint = "335A71B840074E572F616257A6D435784E83F1AB"
$adminUrl = "https://m365cpi90282478-admin.sharepoint.com"
$testSiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs"
$authContextName = "c1"

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  App ID: $appId"
Write-Host "  Tenant: $tenantId"
Write-Host "  Admin URL: $adminUrl"
Write-Host "  Test Site: $testSiteUrl"
Write-Host "  Auth Context: $authContextName`n"

# Check for PnP.PowerShell module
Write-Host "Checking PnP.PowerShell module..." -ForegroundColor Yellow
$module = Get-Module -Name PnP.PowerShell -ListAvailable | Select-Object -First 1
if ($module) {
    Write-Host "✅ Module found: Version $($module.Version)" -ForegroundColor Green
} else {
    Write-Host "❌ Module not found. Please install with:" -ForegroundColor Red
    Write-Host "   Install-Module -Name PnP.PowerShell -Force -AllowClobber" -ForegroundColor White
    exit
}

# Find certificate in local store
Write-Host "`nSearching for certificate..." -ForegroundColor Yellow
$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
if (-not $cert) {
    $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
}

if (-not $cert) {
    Write-Host "❌ Certificate not found in local stores!" -ForegroundColor Red
    Write-Host "`nThe certificate needs to be installed locally for testing." -ForegroundColor Yellow
    Write-Host "Checking if certificate file exists..." -ForegroundColor Yellow
    
    # Check for certificate file from setup script
    if (Test-Path ".\SPO-ConditionalAccess-Cert.pfx") {
        Write-Host "✅ Found certificate file!" -ForegroundColor Green
        Write-Host "`nImporting certificate..." -ForegroundColor Yellow
        $certPassword = Read-Host "Enter certificate password (from setup script)" -AsSecureString
        $cert = Import-PfxCertificate -FilePath ".\SPO-ConditionalAccess-Cert.pfx" -CertStoreLocation Cert:\CurrentUser\My -Password $certPassword
        Write-Host "✅ Certificate imported to CurrentUser\My" -ForegroundColor Green
    } else {
        Write-Host "❌ Certificate file not found!" -ForegroundColor Red
        Write-Host "`nOptions:" -ForegroundColor Yellow
        Write-Host "1. Re-run the certificate setup script (7-Setup-CertificateAuth.ps1)" -ForegroundColor White
        Write-Host "2. Export the certificate from Azure Automation" -ForegroundColor White
        exit
    }
}

Write-Host "✅ Certificate found!" -ForegroundColor Green
Write-Host "   Location: $($cert.PSPath)" -ForegroundColor Gray
Write-Host "   Subject: $($cert.Subject)" -ForegroundColor Gray
Write-Host "   Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
Write-Host "   Has Private Key: $($cert.HasPrivateKey)" -ForegroundColor Gray

if (-not $cert.HasPrivateKey) {
    Write-Host "❌ Certificate doesn't have private key!" -ForegroundColor Red
    exit
}

# Test connection with PnP
Write-Host "`n--- TEST 1: Connect with Certificate Object ---" -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Tenant "$tenantId" -Certificate $cert
    Write-Host "✅ Connected successfully!" -ForegroundColor Green
    
    # Test getting site
    Write-Host "`nTesting Get-PnPTenantSite..." -ForegroundColor Yellow
    $site = Get-PnPTenantSite -Url $testSiteUrl
    Write-Host "✅ Site retrieved!" -ForegroundColor Green
    Write-Host "   URL: $($site.Url)" -ForegroundColor Gray
    Write-Host "   Conditional Access Policy: $($site.ConditionalAccessPolicy)" -ForegroundColor Gray
    
    Disconnect-PnPOnline
    Write-Host "✅ Test 1 PASSED!`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Test 1 FAILED!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)`n" -ForegroundColor Red
}

# Test with certificate thumbprint
Write-Host "--- TEST 2: Connect with Thumbprint ---" -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Tenant "$tenantId" -Thumbprint $cert.Thumbprint
    Write-Host "✅ Connected successfully!" -ForegroundColor Green
    
    Disconnect-PnPOnline
    Write-Host "✅ Test 2 PASSED!`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Test 2 FAILED!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)`n" -ForegroundColor Red
}

# Test applying conditional access policy
Write-Host "--- TEST 3: Apply Conditional Access Policy ---" -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Tenant "$tenantId" -Certificate $cert
    Write-Host "✅ Connected" -ForegroundColor Green
    
    Write-Host "`nApplying policy..." -ForegroundColor Yellow
    Set-PnPTenantSite -Url $testSiteUrl -ConditionalAccessPolicy AuthenticationContext
    Write-Host "✅ Policy set to AuthenticationContext" -ForegroundColor Green
    
    Write-Host "`nSetting authentication context name: $authContextName" -ForegroundColor Yellow
    # Get site using CSOM
    $ctx = Get-PnPContext
    $site = $ctx.Web.ParentWeb.Site
    $ctx.Load($site)
    $ctx.ExecuteQuery()
    $site.AuthenticationContextName = $authContextName
    $site.Update()
    $ctx.ExecuteQuery()
    Write-Host "✅ Authentication context name set" -ForegroundColor Green
    
    # Verify
    Start-Sleep -Seconds 2
    Write-Host "`nVerifying..." -ForegroundColor Yellow
    $verifiedSite = Get-PnPTenantSite -Url $testSiteUrl
    Write-Host "  Conditional Access Policy: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor Gray
    
    if ($verifiedSite.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Host "✅ Test 3 PASSED!" -ForegroundColor Green
        Write-Host "`n🎉 SUCCESS! The approach works locally!" -ForegroundColor Green
        Write-Host "   We can now use this in Azure Automation.`n" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Policy applied but verification shows: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor Yellow
    }
    
    Disconnect-PnPOnline
} catch {
    Write-Host "❌ Test 3 FAILED!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   $($_.Exception.ToString())`n" -ForegroundColor DarkGray
    try { Disconnect-PnPOnline } catch {}
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Test Complete" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan
