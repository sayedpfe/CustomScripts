# Simplified SharePoint Sites with Viva Engage Community Detection
# This version uses a more reliable approach for authentication and site retrieval

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "VivaEngageSites_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory=$false)]
    [bool]$ExportToCsv = $true,
    
    [Parameter(Mandatory=$false)]
    [string]$DebugSiteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
)

# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module is not installed. Please run: Install-Module Microsoft.Graph -Force"
    exit 1
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
# Connect using device code (more reliable for automation)
try {
     Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Host "Please ensure you have the required permissions and try again." -ForegroundColor Yellow
    exit 1
}

# Verify connection
$context = Get-MgContext
if (-not $context) {
    Write-Error "Not connected to Microsoft Graph."
    exit 1
}

Write-Host "Authenticated as: $($context.Account)" -ForegroundColor Gray

# Get all SharePoint sites using Get-MgSite
Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Yellow

try {
    $allSites = Get-MgSite -All -Property "id,displayName,webUrl,createdDateTime,lastModifiedDateTime,root"
    Write-Host "Found $($allSites.Count) SharePoint sites" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve sites: $($_.Exception.Message)"
    exit 1
}

if ($allSites.Count -eq 0) {
    Write-Warning "No sites found. This might be due to insufficient permissions or tenant configuration."
    exit 1
}

# Initialize results
$results = @()
$vivaEngageCount = 0
$processedCount = 0

Write-Host "Analyzing sites for Viva Engage connections..." -ForegroundColor Yellow

# Check if our debug site is in the list
$debugSite = $allSites | Where-Object { $_.webUrl -eq $DebugSiteUrl }
if ($debugSite) {
    Write-Host "✓ Found debug site '$($debugSite.displayName)' in the site list" -ForegroundColor Green
    Write-Host "  Site ID: $($debugSite.id)" -ForegroundColor Gray
    Write-Host "  Group ID: $($debugSite.root.siteCollection.groupId)" -ForegroundColor Gray
} else {
    Write-Host "⚠ Debug site '$DebugSiteUrl' not found in site list" -ForegroundColor Yellow
    Write-Host "Available sites containing 'viva':" -ForegroundColor Gray
    $allSites | Where-Object { $_.webUrl -like "*viva*" -or $_.displayName -like "*viva*" } | ForEach-Object {
        Write-Host "  - $($_.displayName): $($_.webUrl)" -ForegroundColor Gray
    }
}

foreach ($site in $allSites) {
    $processedCount++
    Write-Progress -Activity "Analyzing Sites" -Status "Processing $($site.displayName)" -PercentComplete (($processedCount / $allSites.Count) * 100)
    
    # Special debugging for our target site
    $isDebugSite = $site.webUrl -eq $DebugSiteUrl
    if ($isDebugSite) {
        Write-Host "`n=== DEBUGGING TARGET SITE ===" -ForegroundColor Magenta
        Write-Host "Site: $($site.displayName)" -ForegroundColor Magenta
        Write-Host "URL: $($site.webUrl)" -ForegroundColor Magenta
    }
    
    $siteResult = [PSCustomObject]@{
        SiteName                    = $site.displayName
        SiteUrl                     = $site.webUrl
        SiteId                      = $site.id
        GroupId                     = ""
        GroupDisplayName            = ""
        ResourceProvisioningOptions = ""
        IsVivaEngageCommunity       = $false
        LastModified               = $site.lastModifiedDateTime
        CreatedDateTime            = $site.createdDateTime
        ErrorMessage               = ""
    }
    
    try {
        # Check if site has a group ID
        $groupId = $site.root.siteCollection.groupId
        if ($isDebugSite) {
            Write-Host "Group ID from site: $groupId" -ForegroundColor Magenta
        }
        
        if ($groupId) {
            $siteResult.GroupId = $groupId
            
            try {
                # Get group details
                $group = Get-MgGroup -GroupId $groupId -Property "id,displayName,resourceProvisioningOptions" -ErrorAction Stop
                
                if ($isDebugSite) {
                    Write-Host "Group found: $($group.DisplayName)" -ForegroundColor Magenta
                    Write-Host "Resource provisioning options: $($group.resourceProvisioningOptions -join ', ')" -ForegroundColor Magenta
                }
                
                $siteResult.GroupDisplayName = $group.DisplayName
                $siteResult.ResourceProvisioningOptions = ($group.resourceProvisioningOptions -join ',')
                
                # Check for Viva Engage (Yammer) connection
                $hasYammer = $group.resourceProvisioningOptions -contains "Yammer"
                $siteResult.IsVivaEngageCommunity = $hasYammer
                
                if ($isDebugSite) {
                    Write-Host "Has Yammer: $hasYammer" -ForegroundColor Magenta
                    if (!$hasYammer) {
                        Write-Host "Group is connected but Viva Engage is NOT enabled" -ForegroundColor Yellow
                    }
                    Write-Host "=== END DEBUG ===" -ForegroundColor Magenta
                }
                
                if ($hasYammer) {
                    $vivaEngageCount++
                    Write-Host "✓ Found Viva Engage site: $($site.displayName)" -ForegroundColor Green
                } elseif ($isDebugSite) {
                    Write-Host "ℹ Group-connected but no Viva Engage: $($site.displayName)" -ForegroundColor Yellow
                }
            }
            catch {
                $errorMsg = "Could not retrieve group details: $($_.Exception.Message)"
                $siteResult.ErrorMessage = $errorMsg
                if ($isDebugSite) {
                    Write-Host "ERROR getting group: $errorMsg" -ForegroundColor Red
                    Write-Host "=== END DEBUG ===" -ForegroundColor Magenta
                }
                Write-Warning "Could not get group details for $($site.displayName): $($_.Exception.Message)"
            }
        } else {
            if ($isDebugSite) {
                Write-Host "No group ID found - this is not a group-connected site" -ForegroundColor Yellow
                Write-Host "=== END DEBUG ===" -ForegroundColor Magenta
            }
        }
    }
    catch {
        $errorMsg = "Error processing site: $($_.Exception.Message)"
        $siteResult.ErrorMessage = $errorMsg
        if ($isDebugSite) {
            Write-Host "ERROR processing site: $errorMsg" -ForegroundColor Red
            Write-Host "=== END DEBUG ===" -ForegroundColor Magenta
        }
        Write-Warning "Error processing $($site.displayName): $($_.Exception.Message)"
    }
    
    $results += $siteResult
}

Write-Progress -Activity "Analyzing Sites" -Completed

# Display results
Write-Host "`n"
$separator = "=" * 60
Write-Host $separator -ForegroundColor Cyan
Write-Host "VIVA ENGAGE COMMUNITY ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host "Total SharePoint sites analyzed: $($results.Count)" -ForegroundColor White
Write-Host "Sites connected to Viva Engage: $vivaEngageCount" -ForegroundColor Green
Write-Host "Sites NOT connected to Viva Engage: $($results.Count - $vivaEngageCount)" -ForegroundColor Yellow
Write-Host $separator -ForegroundColor Cyan

# Show Viva Engage connected sites
$vivaEngageSites = $results | Where-Object { $_.IsVivaEngageCommunity -eq $true }

if ($vivaEngageSites.Count -gt 0) {
    Write-Host "`nSharePoint Sites Connected to Viva Engage Communities:" -ForegroundColor Green
    $siteSeparator = "-" * 55
    Write-Host $siteSeparator -ForegroundColor Green
    
    $vivaEngageSites | ForEach-Object {
        Write-Host "• $($_.SiteName)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        Write-Host "  Group: $($_.GroupDisplayName)" -ForegroundColor Gray
        Write-Host ""
    }
} else {
    Write-Host "`nNo SharePoint sites found connected to Viva Engage communities." -ForegroundColor Yellow
}

# Export to CSV if requested
if ($ExportToCsv) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Export Viva Engage sites only
        if ($vivaEngageSites.Count -gt 0) {
            $vivaOnlyPath = $OutputPath -replace '\.csv$', '_VivaEngageOnly.csv'
            $vivaEngageSites | Export-Csv -Path $vivaOnlyPath -NoTypeInformation -Encoding UTF8
            Write-Host "Viva Engage sites exported to: $vivaOnlyPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# Return summary
$summary = [PSCustomObject]@{
    TotalSitesAnalyzed = $results.Count
    VivaEngageSites = $vivaEngageCount
    NonVivaEngageSites = $results.Count - $vivaEngageCount
    AnalysisDate = Get-Date
    OutputFile = if ($ExportToCsv) { $OutputPath } else { "Not exported" }
    VivaEngageConnectedSites = $vivaEngageSites
}

Write-Host "`nAnalysis completed successfully!" -ForegroundColor Green
Disconnect-MgGraph | Out-Null
return $summary