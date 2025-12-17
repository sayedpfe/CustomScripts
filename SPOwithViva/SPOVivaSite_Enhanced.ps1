# Enhanced SharePoint Sites with Viva Engage Community Detection
# This version includes fallback methods to find group connections

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "VivaEngageSites_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory=$false)]
    [bool]$ExportToCsv = $true,
    
    [Parameter(Mandatory=$false)]
    [string]$DebugSiteUrl = "https://m365cpi90282478.sharepoint.com/sites/welcometoviva"
)

# Connect to Microsoft Graph
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint -NoWelcome

Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green

# Get all SharePoint sites
Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Yellow
$allSites = Get-MgSite -All -Property "id,displayName,webUrl,createdDateTime,lastModifiedDateTime,root"
Write-Host "Found $($allSites.Count) SharePoint sites" -ForegroundColor Green

# Get all groups for fallback matching
Write-Host "Retrieving all groups for name matching..." -ForegroundColor Yellow
$allGroups = Get-MgGroup -All -Property "id,displayName,resourceProvisioningOptions"
Write-Host "Found $($allGroups.Count) groups" -ForegroundColor Green

# Function to find group by site name
function Find-GroupBySiteName {
    param($siteName, $allGroups)
    
    # Try exact match first
    $exactMatch = $allGroups | Where-Object { $_.DisplayName -eq $siteName }
    if ($exactMatch) { return $exactMatch }
    
    # Try partial matches
    $partialMatch = $allGroups | Where-Object { $_.DisplayName -like "*$siteName*" -or $siteName -like "*$($_.DisplayName)*" }
    if ($partialMatch -and $partialMatch.Count -eq 1) { return $partialMatch }
    
    return $null
}

# Initialize results
$results = @()
$vivaEngageCount = 0
$groupConnectedCount = 0
$processedCount = 0

Write-Host "Analyzing sites for Viva Engage connections..." -ForegroundColor Yellow

foreach ($site in $allSites) {
    $processedCount++
    Write-Progress -Activity "Analyzing Sites" -Status "Processing $($site.displayName)" -PercentComplete (($processedCount / $allSites.Count) * 100)
    
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
        IsGroupConnected           = $false
        MatchMethod                = ""
        LastModified               = $site.lastModifiedDateTime
        CreatedDateTime            = $site.createdDateTime
        ErrorMessage               = ""
    }
    
    try {
        # Method 1: Try to get group ID from site directly
        $groupId = $site.root.siteCollection.groupId
        
        if ($isDebugSite) {
            Write-Host "Direct Group ID: '$groupId'" -ForegroundColor Magenta
        }
        
        # Method 2: If no direct group ID, try to find group by name
        if ([string]::IsNullOrEmpty($groupId)) {
            $matchedGroup = Find-GroupBySiteName -siteName $site.displayName -allGroups $allGroups
            if ($matchedGroup) {
                $groupId = $matchedGroup.Id
                $siteResult.MatchMethod = "Name Match"
                if ($isDebugSite) {
                    Write-Host "Found group by name matching: $($matchedGroup.DisplayName)" -ForegroundColor Magenta
                }
            }
        } else {
            $siteResult.MatchMethod = "Direct"
        }
        
        if (![string]::IsNullOrEmpty($groupId)) {
            $siteResult.GroupId = $groupId
            $siteResult.IsGroupConnected = $true
            $groupConnectedCount++
            
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
                Write-Host "No group found through any method" -ForegroundColor Yellow
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
$separator = "=" * 70
Write-Host $separator -ForegroundColor Cyan
Write-Host "ENHANCED VIVA ENGAGE COMMUNITY ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host "Total SharePoint sites analyzed: $($results.Count)" -ForegroundColor White
Write-Host "Group-connected sites found: $groupConnectedCount" -ForegroundColor Blue
Write-Host "Sites connected to Viva Engage: $vivaEngageCount" -ForegroundColor Green
Write-Host "Group-connected but NO Viva Engage: $($groupConnectedCount - $vivaEngageCount)" -ForegroundColor Yellow
Write-Host "Sites NOT group-connected: $($results.Count - $groupConnectedCount)" -ForegroundColor Gray
Write-Host $separator -ForegroundColor Cyan

# Show Viva Engage connected sites
$vivaEngageSites = $results | Where-Object { $_.IsVivaEngageCommunity -eq $true }
$groupConnectedNoViva = $results | Where-Object { $_.IsGroupConnected -eq $true -and $_.IsVivaEngageCommunity -eq $false }

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

if ($groupConnectedNoViva.Count -gt 0) {
    Write-Host "`nGroup-Connected Sites WITHOUT Viva Engage (potential candidates):" -ForegroundColor Yellow
    $siteSeparator = "-" * 55
    Write-Host $siteSeparator -ForegroundColor Yellow
    
    $groupConnectedNoViva | ForEach-Object {
        Write-Host "• $($_.SiteName)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        Write-Host "  Group: $($_.GroupDisplayName)" -ForegroundColor Gray
        Write-Host "  Match Method: $($_.MatchMethod)" -ForegroundColor Gray
        Write-Host ""
    }
}

# Export to CSV if requested
if ($ExportToCsv) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Export different categories
        if ($vivaEngageSites.Count -gt 0) {
            $vivaOnlyPath = $OutputPath -replace '\.csv$', '_VivaEngageOnly.csv'
            $vivaEngageSites | Export-Csv -Path $vivaOnlyPath -NoTypeInformation -Encoding UTF8
            Write-Host "Viva Engage sites exported to: $vivaOnlyPath" -ForegroundColor Green
        }
        
        if ($groupConnectedNoViva.Count -gt 0) {
            $candidatesPath = $OutputPath -replace '\.csv$', '_GroupConnectedNoViva.csv'
            $groupConnectedNoViva | Export-Csv -Path $candidatesPath -NoTypeInformation -Encoding UTF8
            Write-Host "Group-connected sites without Viva Engage exported to: $candidatesPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# Return summary
$summary = [PSCustomObject]@{
    TotalSitesAnalyzed = $results.Count
    GroupConnectedSites = $groupConnectedCount
    VivaEngageSites = $vivaEngageCount
    GroupConnectedNoViva = $groupConnectedCount - $vivaEngageCount
    NonGroupConnectedSites = $results.Count - $groupConnectedCount
    AnalysisDate = Get-Date
    OutputFile = if ($ExportToCsv) { $OutputPath } else { "Not exported" }
    VivaEngageConnectedSites = $vivaEngageSites
    GroupConnectedNoVivaSites = $groupConnectedNoViva
}

Write-Host "`nAnalysis completed successfully!" -ForegroundColor Green
Disconnect-MgGraph | Out-Null
return $summary