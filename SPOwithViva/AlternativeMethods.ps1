# Alternative Methods to Get SharePoint Site Group ID
# This script demonstrates multiple ways to get the group ID for a SharePoint site

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
)

Write-Host "=== MULTIPLE METHODS TO GET SHAREPOINT SITE GROUP ID ===" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl" -ForegroundColor Yellow
Write-Host ""

# Method 1: PnP PowerShell
Write-Host "--- Method 1: PnP PowerShell ---" -ForegroundColor Green
try {
    # Check if PnP PowerShell is installed
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Host "❌ PnP.PowerShell module not installed" -ForegroundColor Red
        Write-Host "   Install with: Install-Module PnP.PowerShell -Force" -ForegroundColor Yellow
    } else {
        Write-Host "✓ PnP.PowerShell module found" -ForegroundColor Green
        
        # Connect using interactive authentication
        Write-Host "Connecting to PnP..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        # Get site information including group ID
        $siteInfo = Get-PnPSite -Include "GroupId"
        Write-Host "Site Title: $($siteInfo.Title)" -ForegroundColor White
        Write-Host "Group ID (PnP): $($siteInfo.GroupId)" -ForegroundColor Green
        
        # Alternative: Get web properties
        $web = Get-PnPWeb -Includes "AssociatedOwnerGroup"
        Write-Host "Web Title: $($web.Title)" -ForegroundColor White
        
        # Get Microsoft 365 Group info if available
        if ($siteInfo.GroupId -and $siteInfo.GroupId -ne [System.Guid]::Empty) {
            try {
                $groupInfo = Get-PnPMicrosoft365Group -Identity $siteInfo.GroupId
                Write-Host "Group Display Name: $($groupInfo.DisplayName)" -ForegroundColor Green
                Write-Host "Group Mail: $($groupInfo.Mail)" -ForegroundColor Gray
                
                # Check for Yammer in resource provisioning
                if ($groupInfo.ResourceProvisioningOptions) {
                    Write-Host "Resource Provisioning Options: $($groupInfo.ResourceProvisioningOptions -join ', ')" -ForegroundColor White
                    $hasYammer = $groupInfo.ResourceProvisioningOptions -contains "Yammer"
                    Write-Host "Has Viva Engage (Yammer): $hasYammer" -ForegroundColor $(if($hasYammer){'Green'}else{'Red'})
                } else {
                    Write-Host "No Resource Provisioning Options found" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "Could not get M365 Group details: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "❌ No Group ID found - not a group-connected site" -ForegroundColor Red
        }
        
        Disconnect-PnPOnline
    }
}
catch {
    Write-Host "❌ PnP PowerShell method failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Method 2: SharePoint Online Management Shell
Write-Host "--- Method 2: SharePoint Online Management Shell ---" -ForegroundColor Green
try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
        Write-Host "❌ SharePoint Online Management Shell not installed" -ForegroundColor Red
        Write-Host "   Install with: Install-Module Microsoft.Online.SharePoint.PowerShell -Force" -ForegroundColor Yellow
    } else {
        Write-Host "✓ SharePoint Online Management Shell found" -ForegroundColor Green
        
        # Extract tenant name from URL
        $tenantName = ($SiteUrl -split "//")[1].Split(".")[0]
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        
        Write-Host "Connecting to SPO Admin..." -ForegroundColor Yellow
        Connect-SPOService -Url $adminUrl
        
        # Get site properties
        $siteProperties = Get-SPOSite -Identity $SiteUrl -Detailed
        Write-Host "Site Title: $($siteProperties.Title)" -ForegroundColor White
        Write-Host "Group ID (SPO): $($siteProperties.GroupId)" -ForegroundColor Green
        Write-Host "Is Group Connected: $($siteProperties.IsGroupConnected)" -ForegroundColor White
        
        if ($siteProperties.GroupId -and $siteProperties.GroupId -ne [System.Guid]::Empty) {
            Write-Host "✓ Group-connected site found" -ForegroundColor Green
        } else {
            Write-Host "❌ Not a group-connected site" -ForegroundColor Red
        }
        
        Disconnect-SPOService
    }
}
catch {
    Write-Host "❌ SPO Management Shell method failed: $($_.Exception.Message)" -ForegroundColor Red
}
}

Write-Host ""

# Method 3: Microsoft Graph REST API directly
Write-Host "--- Method 3: Graph REST API (Alternative Endpoints) ---" -ForegroundColor Green
try {
    # Connect to Graph if not already connected
    $config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
    Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint -NoWelcome
    
    # Method 3a: Get site by hostname:path and check properties
    $hostPath = "m365cpi90282478.sharepoint.com:/sites/welcometoviva"
    $siteDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$hostPath" -OutputType PSObject
    
    Write-Host "Site Name: $($siteDetails.displayName)" -ForegroundColor White
    Write-Host "Site ID: $($siteDetails.id)" -ForegroundColor Gray
    
    # Check all possible group ID locations
    Write-Host "Checking all possible group ID properties:" -ForegroundColor Yellow
    Write-Host "  root.siteCollection.groupId: '$($siteDetails.root.siteCollection.groupId)'" -ForegroundColor Gray
    Write-Host "  siteCollection.groupId: '$($siteDetails.siteCollection.groupId)'" -ForegroundColor Gray
    Write-Host "  groupId: '$($siteDetails.groupId)'" -ForegroundColor Gray
    
    # Method 3b: Try to get associated group by searching
    Write-Host "`nSearching for associated Microsoft 365 Group..." -ForegroundColor Yellow
    $groupSearchResults = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq 'Welcome to Viva'&`$select=id,displayName,resourceProvisioningOptions" -OutputType PSObject
    
    if ($groupSearchResults.value -and $groupSearchResults.value.Count -gt 0) {
        $group = $groupSearchResults.value[0]
        Write-Host "✓ Found matching group!" -ForegroundColor Green
        Write-Host "  Group ID: $($group.id)" -ForegroundColor Green
        Write-Host "  Group Name: $($group.displayName)" -ForegroundColor White
        Write-Host "  Resource Provisioning: $($group.resourceProvisioningOptions -join ', ')" -ForegroundColor White
        
        $hasYammer = $group.resourceProvisioningOptions -contains "Yammer"
        Write-Host "  Has Viva Engage: $hasYammer" -ForegroundColor $(if($hasYammer){'Green'}else{'Red'})
        
        # Method 3c: Verify this group owns the site
        Write-Host "`nVerifying group owns the site..." -ForegroundColor Yellow
        try {
            $groupSite = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/sites/root" -OutputType PSObject
            Write-Host "  Group's root site: $($groupSite.webUrl)" -ForegroundColor Gray
            
            if ($groupSite.webUrl -eq $SiteUrl) {
                Write-Host "  ✅ CONFIRMED: This group owns the site!" -ForegroundColor Green
            } else {
                Write-Host "  ⚠ Group has a different root site" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "  ❌ Could not get group's root site: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ No matching group found" -ForegroundColor Red
    }
    
    Disconnect-MgGraph | Out-Null
}
catch {
    Write-Host "❌ Graph REST API method failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Method 4: SharePoint REST API
Write-Host "--- Method 4: SharePoint REST API ---" -ForegroundColor Green
try {
    Write-Host "SharePoint REST API endpoints to try:" -ForegroundColor Yellow
    Write-Host "  1. $SiteUrl/_api/site" -ForegroundColor Gray
    Write-Host "  2. $SiteUrl/_api/web" -ForegroundColor Gray
    Write-Host "  3. $SiteUrl/_api/site/groupid" -ForegroundColor Gray
    Write-Host "  4. $SiteUrl/_api/web/associatedmembergroup" -ForegroundColor Gray
    Write-Host ""
    Write-Host "💡 These require authentication to the SharePoint site directly" -ForegroundColor Yellow
    Write-Host "   Example: Invoke-RestMethod -Uri '$SiteUrl/_api/site' -Headers @{Authorization='Bearer <token>'}" -ForegroundColor Gray
}
catch {
    Write-Host "❌ SharePoint REST API method info failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "✅ PnP PowerShell: Most reliable for SharePoint-specific properties" -ForegroundColor Green
Write-Host "✅ SPO Management Shell: Good for admin-level site information" -ForegroundColor Green
Write-Host "✅ Graph API with name search: Works when direct group ID isn't available" -ForegroundColor Green
Write-Host "✅ SharePoint REST API: Direct access to site properties (requires auth)" -ForegroundColor Green
Write-Host ""
Write-Host "Recommendation: Use PnP PowerShell for the most accurate results!" -ForegroundColor Yellow