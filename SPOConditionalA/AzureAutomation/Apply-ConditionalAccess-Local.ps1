# Local script to apply SharePoint conditional access policy
# Uses the Azure AD App and certificate we created

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AuthenticationContextName
)

$ErrorActionPreference = "Stop"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Apply Conditional Access Policy - LOCAL" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Configuration (from Azure AD App)
$appId = "7bf03e05-9069-412e-a27c-f2f6326fc69f"
$tenantId = "b22f8675-8375-455b-941a-67bee4cf7747"
$tenantName = "M365CPI90282478"
$adminUrl = "https://$tenantName-admin.sharepoint.com"

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  App ID: $appId"
Write-Host "  Tenant: $tenantName"
Write-Host "  Site: $SiteUrl"
Write-Host "  Auth Context: $AuthenticationContextName`n"

# Find the certificate we just created
Write-Host "Looking for certificate..." -ForegroundColor Yellow
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=SPO-Test-Local" } | Select-Object -First 1

if (-not $cert) {
    Write-Host "❌ Certificate not found!" -ForegroundColor Red
    Write-Host "`nPlease run: .\Create-And-Test-Cert.ps1 first to create the certificate" -ForegroundColor Yellow
    exit
}

Write-Host "✅ Certificate found!" -ForegroundColor Green
Write-Host "   Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
Write-Host "   Subject: $($cert.Subject)`n" -ForegroundColor Gray

# Check SharePoint Online module
Write-Host "Checking Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow
try {
    # Import SPO module
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue
    $spoVersion = (Get-Module Microsoft.Online.SharePoint.PowerShell).Version
    Write-Host "✅ Microsoft.Online.SharePoint.PowerShell version $spoVersion loaded`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to load Microsoft.Online.SharePoint.PowerShell" -ForegroundColor Red
    Write-Host "   Installing module..." -ForegroundColor Yellow
    Install-Module Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
    Import-Module Microsoft.Online.SharePoint.PowerShell
    Write-Host "✅ Microsoft.Online.SharePoint.PowerShell installed and loaded`n" -ForegroundColor Green
}

# Connect to SharePoint using certificate authentication
Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Yellow
try {
    # SPO module supports certificate authentication via ClientId, CertificatePath, and Tenant parameters
    # Export certificate temporarily for connection
    $tempCertPath = Join-Path $env:TEMP "SPOTempCert.pfx"
    $certPassword = ConvertTo-SecureString -String "TempPassword123!" -Force -AsPlainText
    
    # Export the certificate
    Export-PfxCertificate -Cert $cert -FilePath $tempCertPath -Password $certPassword | Out-Null
    
    # Connect using certificate
    Connect-SPOService -Url $adminUrl -ClientId $appId -CertificatePath $tempCertPath -CertificatePassword $certPassword
    
    # Clean up temp certificate
    Remove-Item $tempCertPath -Force -ErrorAction SilentlyContinue
    
    Write-Host "✅ Connected successfully!`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Connection failed!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)`n" -ForegroundColor Red
    Write-Host "This might be because:" -ForegroundColor Yellow
    Write-Host "  1. The certificate wasn't uploaded to the App Registration correctly" -ForegroundColor Gray
    Write-Host "  2. The App doesn't have SharePoint API permissions" -ForegroundColor Gray
    Write-Host "  3. Admin consent wasn't granted for the permissions`n" -ForegroundColor Gray
    exit
}

# Apply conditional access policy
Write-Host "Applying conditional access policy..." -ForegroundColor Yellow
try {
    # Use Set-SPOSite to set both policy type and authentication context name
    Set-SPOSite -Identity $SiteUrl -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName $AuthenticationContextName
    
    Write-Host "✅ Conditional access policy applied!" -ForegroundColor Green
    Write-Host "   Policy Type: AuthenticationContext" -ForegroundColor Green
    Write-Host "   Context Name: $AuthenticationContextName`n" -ForegroundColor Green
    
    # Wait for propagation
    Start-Sleep -Seconds 2
    
    # Verify
    Write-Host "Verifying changes..." -ForegroundColor Yellow
    $verifiedSite = Get-SPOSite -Identity $SiteUrl -Detailed
    
    Write-Host "`nResults:" -ForegroundColor Cyan
    Write-Host "  Site URL: $($verifiedSite.Url)" -ForegroundColor White
    Write-Host "  Conditional Access Policy: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor White
    
    if ($verifiedSite.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Host "`n🎉 SUCCESS! Conditional access policy applied!" -ForegroundColor Green
        Write-Host "   The policy 'AuthenticationContext' is now active on this site.`n" -ForegroundColor Green
    } else {
        Write-Host "`n⚠️ Policy was applied but verification shows: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ Failed to apply policy!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Details: $($_.Exception.ToString())`n" -ForegroundColor DarkGray
} finally {
    Disconnect-SPOService
    Write-Host "Disconnected from SharePoint`n" -ForegroundColor Gray
}

Write-Host "========================================`n" -ForegroundColor Cyan
