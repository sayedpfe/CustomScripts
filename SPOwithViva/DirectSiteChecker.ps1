# Direct Site Checker for Viva Engage
# This script will try to access your specific site directly

$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint -NoWelcome

$siteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
Write-Host "Checking site: $siteUrl" -ForegroundColor Yellow

# Method 1: Try to get the site directly using the URL
Write-Host "`n=== Method 1: Direct site lookup ===" -ForegroundColor Cyan
try {
    $encodedUrl = [System.Web.HttpUtility]::UrlEncode($siteUrl)
    $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$encodedUrl"
    Write-Host "✓ Site found via direct lookup!" -ForegroundColor Green
    Write-Host "  Site ID: $($site.id)" -ForegroundColor Gray
    Write-Host "  Display Name: $($site.displayName)" -ForegroundColor Gray
    Write-Host "  Group ID: $($site.root.siteCollection.groupId)" -ForegroundColor Gray
    
    if ($site.root.siteCollection.groupId) {
        $groupId = $site.root.siteCollection.groupId
        Write-Host "`n  Checking associated group..." -ForegroundColor Yellow
        try {
            $group = Get-MgGroup -GroupId $groupId -Property "id,displayName,resourceProvisioningOptions"
            Write-Host "  ✓ Group found: $($group.DisplayName)" -ForegroundColor Green
            Write-Host "  Resource provisioning options: $($group.resourceProvisioningOptions -join ', ')" -ForegroundColor Gray
            
            $hasYammer = $group.resourceProvisioningOptions -contains "Yammer"
            Write-Host "  Has Viva Engage (Yammer): $hasYammer" -ForegroundColor $(if($hasYammer){'Green'}else{'Yellow'})
        }
        catch {
            Write-Host "  ✗ Could not get group details: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  ⚠ No group ID - this is not a group-connected site" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "✗ Site not found via direct lookup: $($_.Exception.Message)" -ForegroundColor Red
}

# Method 2: Search for sites with "welcometoviva" in the name/URL
Write-Host "`n=== Method 2: Search for 'welcometoviva' ===" -ForegroundColor Cyan
try {
    $searchResults = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?search=welcometoviva"
    if ($searchResults.value -and $searchResults.value.Count -gt 0) {
        Write-Host "✓ Found $($searchResults.value.Count) site(s) matching 'welcometoviva'" -ForegroundColor Green
        foreach ($site in $searchResults.value) {
            Write-Host "  - $($site.displayName): $($site.webUrl)" -ForegroundColor Gray
            Write-Host "    Group ID: $($site.root.siteCollection.groupId)" -ForegroundColor Gray
        }
    } else {
        Write-Host "✗ No sites found matching 'welcometoviva'" -ForegroundColor Red
    }
}
catch {
    Write-Host "✗ Search failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Method 3: Search for sites with "viva" in the name
Write-Host "`n=== Method 3: All sites containing 'viva' ===" -ForegroundColor Cyan
try {
    $searchResults = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?search=viva"
    if ($searchResults.value -and $searchResults.value.Count -gt 0) {
        Write-Host "✓ Found $($searchResults.value.Count) site(s) containing 'viva'" -ForegroundColor Green
        foreach ($site in $searchResults.value) {
            Write-Host "  - $($site.displayName): $($site.webUrl)" -ForegroundColor Gray
            if ($site.root.siteCollection.groupId) {
                Write-Host "    Group ID: $($site.root.siteCollection.groupId)" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "✗ No sites found containing 'viva'" -ForegroundColor Red
    }
}
catch {
    Write-Host "✗ Search failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Method 4: Check with hostname:path format
Write-Host "`n=== Method 4: Hostname:path format ===" -ForegroundColor Cyan
try {
    $hostPath = "m365cpi90282478.sharepoint.com:/sites/welcometoviva"
    $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$hostPath"
    Write-Host "✓ Site found via hostname:path format!" -ForegroundColor Green
    Write-Host "  Site ID: $($site.id)" -ForegroundColor Gray
    Write-Host "  Display Name: $($site.displayName)" -ForegroundColor Gray
    Write-Host "  Group ID: $($site.root.siteCollection.groupId)" -ForegroundColor Gray
}
catch {
    Write-Host "✗ Site not found via hostname:path format: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nAnalysis complete!" -ForegroundColor Green
Disconnect-MgGraph | Out-Null