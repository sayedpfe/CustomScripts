# Local test script for SharePoint Online certificate authentication
# This will help us find the right way to authenticate before using in Azure Automation

$ErrorActionPreference = "Stop"

# Configuration
$appId = "7bf03e05-9069-412e-a27c-f2f6326fc69f"
$tenantId = "b22f8675-8375-455b-941a-67bee4cf7747"
$certThumbprint = "335A71B840074E572F616257A6D435784E83F1AB"
$adminUrl = "https://m365cpi90282478-admin.sharepoint.com"
$testSiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Testing SharePoint Certificate Authentication" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  App ID: $appId"
Write-Host "  Tenant: $tenantId"
Write-Host "  Cert Thumbprint: $certThumbprint"
Write-Host "  Admin URL: $adminUrl`n"

# Check if module is loaded
Write-Host "Checking Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow
$module = Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object -First 1
if ($module) {
    Write-Host "✅ Module found: Version $($module.Version)" -ForegroundColor Green
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
} else {
    Write-Host "❌ Module not found. Installing..." -ForegroundColor Red
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber
    Import-Module Microsoft.Online.SharePoint.PowerShell
}

# Check Connect-SPOService parameters
Write-Host "`nChecking Connect-SPOService available parameters..." -ForegroundColor Yellow
$cmdlet = Get-Command Connect-SPOService
Write-Host "Available parameter sets:" -ForegroundColor Cyan
$cmdlet.ParameterSets | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor White
    $_.Parameters | Where-Object { $_.Name -like "*Cert*" -or $_.Name -like "*Client*" -or $_.Name -like "*Tenant*" -or $_.Name -like "*Thumb*" } | ForEach-Object {
        Write-Host "    * $($_.Name) [$($_.ParameterType.Name)]" -ForegroundColor Gray
    }
}

# Try to find the certificate in local store
Write-Host "`nSearching for certificate in local stores..." -ForegroundColor Yellow
$cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
if (-not $cert) {
    $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
}

if ($cert) {
    Write-Host "✅ Certificate found in $($cert.PSPath)" -ForegroundColor Green
    Write-Host "   Subject: $($cert.Subject)" -ForegroundColor Gray
    Write-Host "   Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
    Write-Host "   Has Private Key: $($cert.HasPrivateKey)" -ForegroundColor Gray
} else {
    Write-Host "❌ Certificate NOT found in local stores!" -ForegroundColor Red
    Write-Host "   This is expected - certificate is only in Azure Automation" -ForegroundColor Yellow
    Write-Host "   We'll need to test authentication approach in Azure" -ForegroundColor Yellow
}

# Attempt connection with certificate path (if cert exists locally)
if ($cert -and $cert.HasPrivateKey) {
    Write-Host "`nAttempting connection with certificate object..." -ForegroundColor Yellow
    try {
        Connect-SPOService -Url $adminUrl -ClientId $appId -Tenant $tenantId -Certificate $cert
        Write-Host "✅ SUCCESS! Connected with certificate object" -ForegroundColor Green
        
        # Test Get-SPOSite
        Write-Host "`nTesting Get-SPOSite..." -ForegroundColor Yellow
        $site = Get-SPOSite -Identity $testSiteUrl
        Write-Host "✅ Site retrieved: $($site.Url)" -ForegroundColor Green
        Write-Host "   Conditional Access Policy: $($site.ConditionalAccessPolicy)" -ForegroundColor Gray
        
        Disconnect-SPOService
        Write-Host "`n✅ Test completed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "`n⚠️ Cannot test locally - certificate not in local store" -ForegroundColor Yellow
    Write-Host "   The certificate only exists in Azure Automation" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Recommendation for Azure Automation:" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Based on the parameter check above, use:" -ForegroundColor Yellow
Write-Host '  Connect-SPOService -Url $adminUrl -ClientId $appId -Tenant $tenantId -Certificate $cert' -ForegroundColor White
Write-Host "`nThe certificate object from Get-AutomationCertificate should work." -ForegroundColor Yellow
Write-Host ""
