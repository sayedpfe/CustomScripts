# Local script to apply SharePoint conditional access policy
# Uses Microsoft.Online.SharePoint.PowerShell module with Set-SPOSite cmdlet

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

# Find the certificate we created
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

# Check if Microsoft.Online.SharePoint.PowerShell module is installed
Write-Host "Checking Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow
$spoModule = Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1

if (-not $spoModule) {
    Write-Host "Module not installed. Installing..." -ForegroundColor Yellow
    Install-Module Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
    Write-Host "✅ Module installed`n" -ForegroundColor Green
} else {
    Write-Host "✅ Module found (version $($spoModule.Version))`n" -ForegroundColor Green
}

# Import module (must not have PnP loaded)
try {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
    Write-Host "✅ Module loaded successfully`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to load module: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nThis is likely because PnP.PowerShell is loaded in the current session." -ForegroundColor Yellow
    Write-Host "Please run this script in a fresh PowerShell window.`n" -ForegroundColor Yellow
    exit
}

# Export certificate temporarily for connection
Write-Host "Preparing certificate for connection..." -ForegroundColor Yellow
$tempCertPath = Join-Path $env:TEMP "SPOTempCert_$([guid]::NewGuid()).pfx"
$certPassword = ConvertTo-SecureString -String "TempPassword123!" -Force -AsPlainText

try {
    Export-PfxCertificate -Cert $cert -FilePath $tempCertPath -Password $certPassword -ErrorAction Stop | Out-Null
    Write-Host "✅ Certificate exported`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to export certificate: $($_.Exception.Message)`n" -ForegroundColor Red
    exit
}

# Connect to SharePoint Online Admin Center
Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Yellow
try {
    Connect-SPOService -Url $adminUrl -ClientId $appId -CertificatePath $tempCertPath -CertificatePassword $certPassword
    Write-Host "✅ Connected successfully!`n" -ForegroundColor Green
} catch {
    Write-Host "❌ Connection failed!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)`n" -ForegroundColor Red
    Write-Host "Possible reasons:" -ForegroundColor Yellow
    Write-Host "  1. Certificate not uploaded to Azure AD App Registration" -ForegroundColor Gray
    Write-Host "  2. App missing SharePoint API permissions (Sites.FullControl.All)" -ForegroundColor Gray
    Write-Host "  3. Admin consent not granted for the permissions`n" -ForegroundColor Gray
    
    # Clean up temp certificate
    if (Test-Path $tempCertPath) {
        Remove-Item $tempCertPath -Force -ErrorAction SilentlyContinue
    }
    exit
}

# Clean up temp certificate after successful connection
if (Test-Path $tempCertPath) {
    Remove-Item $tempCertPath -Force -ErrorAction SilentlyContinue
}

# Apply conditional access policy
Write-Host "Applying conditional access policy..." -ForegroundColor Yellow
try {
    # Use Set-SPOSite to set both policy type and authentication context name in one command
    # This is the official Microsoft-documented approach
    Set-SPOSite -Identity $SiteUrl `
                -ConditionalAccessPolicy AuthenticationContext `
                -AuthenticationContextName $AuthenticationContextName
    
    Write-Host "✅ Conditional access policy applied!" -ForegroundColor Green
    Write-Host "   Policy Type: AuthenticationContext" -ForegroundColor Green
    Write-Host "   Context Name: $AuthenticationContextName`n" -ForegroundColor Green
    
    # Wait for propagation
    Write-Host "Waiting for changes to propagate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    
    # Verify
    Write-Host "Verifying changes..." -ForegroundColor Yellow
    $verifiedSite = Get-SPOSite -Identity $SiteUrl -Detailed
    
    Write-Host "`nResults:" -ForegroundColor Cyan
    Write-Host "  Site URL: $($verifiedSite.Url)" -ForegroundColor White
    Write-Host "  Conditional Access Policy: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor White
    
    if ($verifiedSite.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Host "`n🎉 SUCCESS! Conditional access policy applied!" -ForegroundColor Green
        Write-Host "   The policy 'AuthenticationContext' with context '$AuthenticationContextName' is now active.`n" -ForegroundColor Green
    } else {
        Write-Host "`n⚠️ Policy was applied but verification shows: $($verifiedSite.ConditionalAccessPolicy)" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ Failed to apply policy!" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    
    # Check for common errors
    if ($_.Exception.Message -like "*c1*" -or $_.Exception.Message -like "*authentication context*") {
        Write-Host "`n⚠️ Authentication context 'c1' may not exist in your tenant." -ForegroundColor Yellow
        Write-Host "   Please verify the authentication context exists in Azure AD Conditional Access.`n" -ForegroundColor Yellow
    }
} finally {
    Disconnect-SPOService
    Write-Host "Disconnected from SharePoint`n" -ForegroundColor Gray
}

Write-Host "========================================`n" -ForegroundColor Cyan
