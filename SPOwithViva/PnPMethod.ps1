# Simple PnP PowerShell method to get SharePoint Site Group ID
# This is often the most reliable method for SharePoint sites

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
)

Write-Host "=== PnP POWERSHELL METHOD ===" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Yellow

# Check if PnP PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host ""
    Write-Host "❌ PnP.PowerShell module is not installed!" -ForegroundColor Red
    Write-Host ""
    Write-Host "To install PnP PowerShell, run:" -ForegroundColor Yellow
    Write-Host "Install-Module PnP.PowerShell -Force -AllowClobber" -ForegroundColor White
    Write-Host ""
    Write-Host "After installation, re-run this script." -ForegroundColor Yellow
    exit 1
}

try {
    Write-Host ""
    Write-Host "Connecting to SharePoint site using PnP..." -ForegroundColor Green
    
    # Try multiple authentication methods
    $connected = $false
    
    # Method 1: Try using existing credentials
    try {
        Write-Host "Attempting connection with device login..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $SiteUrl -DeviceLogin
        $connected = $true
        Write-Host "✓ Connected successfully using device login!" -ForegroundColor Green
    }
    catch {
        Write-Host "Device login failed, trying alternative method..." -ForegroundColor Yellow
        
        # Method 2: Try using Web Login (legacy but often works)
        try {
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin
            $connected = $true
            Write-Host "✓ Connected successfully using web login!" -ForegroundColor Green
        }
        catch {
            Write-Host "Web login failed, trying certificate-based auth..." -ForegroundColor Yellow
            
            # Method 3: Try using certificate from config (if available)
            try {
                $config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
                Connect-PnPOnline -Url $SiteUrl -ClientId $config.ClientId -Tenant $config.TenantId -Thumbprint $config.Thumbprint
                $connected = $true
                Write-Host "✓ Connected successfully using certificate!" -ForegroundColor Green
            }
            catch {
                throw "All authentication methods failed. Error: $($_.Exception.Message)"
            }
        }
    }
    
    if (-not $connected) {
        throw "Could not connect to SharePoint site"
    }
    
    # Get site information including Group ID
    Write-Host ""
    Write-Host "Getting site information..." -ForegroundColor Yellow
    $site = Get-PnPSite -Include "GroupId"
    
    Write-Host ""
    Write-Host "=== SITE INFORMATION ===" -ForegroundColor Cyan
    Write-Host "Site Title: $($site.Title)" -ForegroundColor White
    Write-Host "Site URL: $($site.Url)" -ForegroundColor Gray
    Write-Host "Group ID: $($site.GroupId)" -ForegroundColor Green
    
    if ($site.GroupId -and $site.GroupId -ne [System.Guid]::Empty) {
        Write-Host ""
        Write-Host "✅ This is a GROUP-CONNECTED site!" -ForegroundColor Green
        
        # Get the Microsoft 365 Group details
        Write-Host ""
        Write-Host "Getting Microsoft 365 Group details..." -ForegroundColor Yellow
        
        try {
            $group = Get-PnPMicrosoft365Group -Identity $site.GroupId
            
            Write-Host ""
            Write-Host "=== GROUP INFORMATION ===" -ForegroundColor Cyan
            Write-Host "Group Display Name: $($group.DisplayName)" -ForegroundColor White
            Write-Host "Group Email: $($group.Mail)" -ForegroundColor Gray
            Write-Host "Group ID: $($group.Id)" -ForegroundColor Gray
            Write-Host "Description: $($group.Description)" -ForegroundColor Gray
            
            # Check for Viva Engage (Yammer)
            if ($group.ResourceProvisioningOptions) {
                Write-Host "Resource Provisioning Options: $($group.ResourceProvisioningOptions -join ', ')" -ForegroundColor White
                
                $hasYammer = $group.ResourceProvisioningOptions -contains "Yammer"
                Write-Host ""
                if ($hasYammer) {
                    Write-Host "🎉 SUCCESS: This site HAS Viva Engage (Yammer) enabled!" -ForegroundColor Green
                } else {
                    Write-Host "⚠️  RESULT: This site does NOT have Viva Engage (Yammer) enabled" -ForegroundColor Yellow
                    Write-Host "   The group exists but Viva Engage is not provisioned." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Resource Provisioning Options: None" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "⚠️  RESULT: No resource provisioning options found" -ForegroundColor Yellow
                Write-Host "   This group does not have Viva Engage enabled." -ForegroundColor Yellow
            }
            
        }
        catch {
            Write-Host ""
            Write-Host "❌ Could not get Microsoft 365 Group details: $($_.Exception.Message)" -ForegroundColor Red
        }
        
    } else {
        Write-Host ""
        Write-Host "❌ This is NOT a group-connected site" -ForegroundColor Red
        Write-Host "   Only group-connected sites can have Viva Engage communities." -ForegroundColor Yellow
    }
    
    # Additional site properties that might be useful
    Write-Host ""
    Write-Host "=== ADDITIONAL SITE PROPERTIES ===" -ForegroundColor Cyan
    $web = Get-PnPWeb
    Write-Host "Web Title: $($web.Title)" -ForegroundColor White
    Write-Host "Web Template: $($web.WebTemplate)" -ForegroundColor Gray
    Write-Host "Created: $($web.Created)" -ForegroundColor Gray
    Write-Host "Last Modified: $($web.LastItemModifiedDate)" -ForegroundColor Gray
    
    Write-Host ""
    Write-Host "✅ PnP PowerShell analysis completed!" -ForegroundColor Green
    
}
catch {
    Write-Host ""
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Always disconnect
    try {
        Disconnect-PnPOnline
        Write-Host ""
        Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
    }
    catch {
        # Ignore disconnect errors
    }
}

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "✅ PnP PowerShell provides the most reliable access to SharePoint site properties" -ForegroundColor Green
Write-Host "✅ It can directly access the GroupId property of the site" -ForegroundColor Green
Write-Host "✅ It can also retrieve Microsoft 365 Group details including Viva Engage status" -ForegroundColor Green