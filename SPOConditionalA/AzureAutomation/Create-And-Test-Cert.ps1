# Create a test certificate locally and test PnP auth
# This cert will be used for local testing only

$ErrorActionPreference = "Stop"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Create Local Test Certificate" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$appId = "7bf03e05-9069-412e-a27c-f2f6326fc69f"
$tenantId = "b22f8675-8375-455b-941a-67bee4cf7747"
$certName = "CN=SPO-Test-Local"

# Create self-signed certificate
Write-Host "Creating self-signed certificate..." -ForegroundColor Yellow
$cert = New-SelfSignedCertificate `
    -Subject $certName `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

Write-Host "✅ Certificate created!" -ForegroundColor Green
Write-Host "   Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
Write-Host "   Subject: $($cert.Subject)" -ForegroundColor Gray

# Export to PFX
$pfxPath = ".\SPO-Test-Cert.pfx"
$pfxPassword = ConvertTo-SecureString -String "TestPassword123!" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword | Out-Null
Write-Host "✅ Exported to: $pfxPath" -ForegroundColor Green

# Upload certificate to Azure AD App
Write-Host "`nUploading certificate to Azure AD App..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome
    
    $cerBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    $cerValue = [System.Convert]::ToBase64String($cerBytes)
    
    Update-MgApplication -ApplicationId (Get-MgApplication -Filter "appId eq '$appId'").Id -KeyCredentials @{
        Type = "AsymmetricX509Cert"
        Usage = "Verify"
        Key = $cerBytes
    }
    
    Write-Host "✅ Certificate uploaded to App Registration" -ForegroundColor Green
} catch {
    Write-Host "⚠️ Failed to upload to Azure AD: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "   You'll need to upload manually or the existing cert in Azure is fine" -ForegroundColor Gray
}

# Now test PnP connection
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Testing PnP Connection" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$adminUrl = "https://m365cpi90282478-admin.sharepoint.com"
$testSiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs"

Write-Host "Connecting with PnP.PowerShell..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Tenant $tenantId -Certificate $cert
    Write-Host "✅ Connected successfully!" -ForegroundColor Green
    
    Write-Host "`nTesting Get-PnPTenantSite..." -ForegroundColor Yellow
    $site = Get-PnPTenantSite -Url $testSiteUrl
    Write-Host "✅ Site retrieved!" -ForegroundColor Green
    Write-Host "   Conditional Access Policy: $($site.ConditionalAccessPolicy)" -ForegroundColor Gray
    
    Write-Host "`nTesting Set-PnPTenantSite..." -ForegroundColor Yellow
    Set-PnPTenantSite -Url $testSiteUrl -ConditionalAccessPolicy AuthenticationContext
    Write-Host "✅ Policy set successfully!" -ForegroundColor Green
    
    Disconnect-PnPOnline
    
    Write-Host "`n🎉 SUCCESS! PnP certificate authentication works!" -ForegroundColor Green
    Write-Host "`nCertificate Details for Azure Automation:" -ForegroundColor Cyan
    Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
    Write-Host "  PFX File: $pfxPath" -ForegroundColor White
    Write-Host "  Password: TestPassword123!" -ForegroundColor White
    
} catch {
    Write-Host "❌ Test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nThis might be because:" -ForegroundColor Yellow
    Write-Host "  1. The certificate wasn't uploaded to the App Registration" -ForegroundColor Gray
    Write-Host "  2. The App doesn't have SharePoint permissions" -ForegroundColor Gray
    Write-Host "  3. Admin consent wasn't granted" -ForegroundColor Gray
}

Write-Host "`n========================================`n" -ForegroundColor Cyan
