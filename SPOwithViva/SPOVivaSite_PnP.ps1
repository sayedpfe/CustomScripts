# SharePoint Sites with Viva Engage Community Detection - PnP PowerShell Version
# Uses Get-PnPTenantSite for more reliable group-connected site detection

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "VivaEngageSites_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory=$false)]
    [bool]$ExportToCsv = $true
)

# Check if PnP PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module is not installed. Please run: Install-Module PnP.PowerShell -Force"
    exit 1
}

Write-Host "=== VIVA ENGAGE SITE DETECTION USING PnP POWERSHELL ===" -ForegroundColor Cyan
Write-Host ""

# Load configuration and connect
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
$tenantUrl = "https://m365cpi90282478-admin.sharepoint.com"

Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Green
try {
    Connect-PnPOnline -Url $tenantUrl -ClientId $config.ClientId -Tenant $config.TenantId -Thumbprint $config.Thumbprint
    Write-Host "✓ Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect: $($_.Exception.Message)"
    exit 1
}

# Get all group-connected sites
Write-Host "Retrieving all group-connected SharePoint sites..." -ForegroundColor Yellow
try {
    $groupConnectedSites = Get-PnPTenantSite -GroupIdDefined $true
    Write-Host "✓ Found $($groupConnectedSites.Count) group-connected sites" -ForegroundColor Green
}
catch {
    Write-Error "Failed to get sites: $($_.Exception.Message)"
    Disconnect-PnPOnline
    exit 1
}

# Initialize results
$results = @()
$vivaEngageCount = 0
$processedCount = 0

Write-Host "Analyzing groups for Viva Engage connections..." -ForegroundColor Yellow
Write-Host ""

foreach ($site in $groupConnectedSites) {
    $processedCount++
    Write-Progress -Activity "Analyzing Group-Connected Sites" -Status "Processing $($site.Title)" -PercentComplete (($processedCount / $groupConnectedSites.Count) * 100)
    
    $siteResult = [PSCustomObject]@{
        SiteName                    = $site.Title
        SiteUrl                     = $site.Url
        GroupId                     = $site.GroupId
        GroupDisplayName            = ""
        CreationOptions             = ""
        IsVivaEngageCommunity       = $false
        StorageQuota               = $site.StorageQuota
        LastContentModifiedDate    = $site.LastContentModifiedDate
        ErrorMessage               = ""
    }
    
    if ($site.GroupId -and $site.GroupId -ne [System.Guid]::Empty) {
        try {
            # Get group details using PnP
            $group = Get-PnPMicrosoft365Group -Identity $site.GroupId -ErrorAction Stop
            
            $siteResult.GroupDisplayName = $group.DisplayName
            
            # Check both CreationOptions and ResourceProvisioningOptions for Viva Engage
            $creationOptions = $group.CreationOptions -join ','
            $resourceProvisioningOptions = $group.ResourceProvisioningOptions -join ','
            
            $siteResult.CreationOptions = if ($creationOptions) { $creationOptions } else { $resourceProvisioningOptions }
            
            # Check for Viva Engage (Yammer) - use CreationOptions which contains "YammerProvisioning"
            $hasYammer = ($group.CreationOptions -contains "YammerProvisioning") -or ($group.ResourceProvisioningOptions -contains "Yammer")
            $siteResult.IsVivaEngageCommunity = $hasYammer
            
            if ($hasYammer) {
                $vivaEngageCount++
                Write-Host "✓ Found Viva Engage site: $($site.Title)" -ForegroundColor Green
            }
        }
        catch {
            $siteResult.ErrorMessage = "Could not retrieve group details: $($_.Exception.Message)"
            Write-Warning "Could not get group details for $($site.Title): $($_.Exception.Message)"
        }
    }
    
    $results += $siteResult
}

Write-Progress -Activity "Analyzing Group-Connected Sites" -Completed

# Disconnect from PnP
Disconnect-PnPOnline

# Display results
Write-Host ""
$separator = "=" * 70
Write-Host $separator -ForegroundColor Cyan
Write-Host "VIVA ENGAGE COMMUNITY ANALYSIS SUMMARY (PnP PowerShell)" -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host "Total group-connected sites analyzed: $($results.Count)" -ForegroundColor White
Write-Host "Sites connected to Viva Engage: $vivaEngageCount" -ForegroundColor Green
Write-Host "Group-connected WITHOUT Viva Engage: $($results.Count - $vivaEngageCount)" -ForegroundColor Yellow
Write-Host $separator -ForegroundColor Cyan

# Show Viva Engage connected sites
$vivaEngageSites = $results | Where-Object { $_.IsVivaEngageCommunity -eq $true }
$groupConnectedNoViva = $results | Where-Object { $_.IsVivaEngageCommunity -eq $false }

if ($vivaEngageSites.Count -gt 0) {
    Write-Host ""
    Write-Host "SharePoint Sites Connected to Viva Engage Communities:" -ForegroundColor Green
    $siteSeparator = "-" * 60
    Write-Host $siteSeparator -ForegroundColor Green
    
    $vivaEngageSites | ForEach-Object {
        Write-Host "• $($_.SiteName)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        Write-Host "  Group: $($_.GroupDisplayName)" -ForegroundColor Gray
        Write-Host "  Group ID: $($_.GroupId)" -ForegroundColor DarkGray
        Write-Host ""
    }
} else {
    Write-Host ""
    Write-Host "No SharePoint sites found connected to Viva Engage communities." -ForegroundColor Yellow
}

# Optionally show sites without Viva Engage
if ($groupConnectedNoViva.Count -gt 0 -and $groupConnectedNoViva.Count -le 10) {
    Write-Host ""
    Write-Host "Sample Group-Connected Sites WITHOUT Viva Engage:" -ForegroundColor Yellow
    $siteSeparator = "-" * 60
    Write-Host $siteSeparator -ForegroundColor Yellow
    
    $groupConnectedNoViva | Select-Object -First 10 | ForEach-Object {
        Write-Host "• $($_.SiteName)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        Write-Host "  Group: $($_.GroupDisplayName)" -ForegroundColor Gray
        Write-Host ""
    }
}

# Export to CSV if requested
if ($ExportToCsv) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Results exported to: $OutputPath" -ForegroundColor Green
        
        # Export Viva Engage sites only
        if ($vivaEngageSites.Count -gt 0) {
            $vivaOnlyPath = $OutputPath -replace '\.csv$', '_VivaEngageOnly.csv'
            $vivaEngageSites | Export-Csv -Path $vivaOnlyPath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ Viva Engage sites exported to: $vivaOnlyPath" -ForegroundColor Green
        }
        
        # Export sites without Viva Engage
        if ($groupConnectedNoViva.Count -gt 0) {
            $noVivaPath = $OutputPath -replace '\.csv$', '_NoVivaEngage.csv'
            $groupConnectedNoViva | Export-Csv -Path $noVivaPath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ Group-connected sites without Viva Engage exported to: $noVivaPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# Return summary
$summary = [PSCustomObject]@{
    TotalGroupConnectedSites = $results.Count
    VivaEngageSites = $vivaEngageCount
    GroupConnectedNoViva = $results.Count - $vivaEngageCount
    AnalysisDate = Get-Date
    OutputFile = if ($ExportToCsv) { $OutputPath } else { "Not exported" }
    VivaEngageConnectedSites = $vivaEngageSites
}

Write-Host ""
Write-Host "✅ Analysis completed successfully!" -ForegroundColor Green
return $summary