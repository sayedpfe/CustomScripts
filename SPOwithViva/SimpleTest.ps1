# Simple test to get group ID correctly
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint -NoWelcome

Write-Host "Connected to Graph" -ForegroundColor Green

$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
Write-Host "Getting site: $SiteUrl" -ForegroundColor Yellow

try {
    # Method 1: Using hostname:path format
    $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/m365cpi90282478.sharepoint.com:/sites/welcometoviva"
    
    Write-Host "Site found!" -ForegroundColor Green
    Write-Host "Display Name: $($site.displayName)"
    Write-Host "Site ID: $($site.id)"
    
    # THE CORRECT WAY to get group ID
    $groupId = $site.root.siteCollection.groupId
    Write-Host "Group ID: '$groupId'" -ForegroundColor Yellow
    
    if ([string]::IsNullOrEmpty($groupId)) {
        Write-Host "❌ No Group ID - This is NOT a group-connected site" -ForegroundColor Red
    } else {
        Write-Host "✅ Group ID found: $groupId" -ForegroundColor Green
    }
    
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Disconnect-MgGraph | Out-Null