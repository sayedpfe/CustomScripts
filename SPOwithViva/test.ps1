# Test script for Get-PnPTenantSite with GroupIdDefined
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to admin site
$tenantUrl = "https://m365cpi90282478-admin.sharepoint.com"
Write-Host "Connecting to tenant admin: $tenantUrl" -ForegroundColor Green
Connect-PnPOnline -Url $tenantUrl -ClientId $config.ClientId -Tenant $config.TenantId -Thumbprint $config.Thumbprint

Write-Host "Getting all group-connected sites..." -ForegroundColor Yellow
$groupSites = Get-PnPTenantSite -GroupIdDefined $true

Write-Host "Found $($groupSites.Count) group-connected sites" -ForegroundColor Green

# Display first few sites
$groupSites | Select-Object -First 5 | Format-Table Url, Title, GroupId -AutoSize

# Check for our specific site
$targetSite = $groupSites | Where-Object { $_.Url -like "*welcometoviva*" }
if ($targetSite) {
    Write-Host "`n✓ Found 'welcometoviva' site!" -ForegroundColor Green
    Write-Host "  URL: $($targetSite.Url)" -ForegroundColor White
    Write-Host "  Title: $($targetSite.Title)" -ForegroundColor White
    Write-Host "  GroupId: $($targetSite.GroupId)" -ForegroundColor Green
} else {
    Write-Host "`n✗ 'welcometoviva' site not found in group-connected sites" -ForegroundColor Red
}

Disconnect-PnPOnline