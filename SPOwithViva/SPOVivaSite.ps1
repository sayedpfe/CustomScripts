# Requires Microsoft.Graph modules (e.g., Microsoft.Graph.Beta or v2 SDK core modules)
# Connect with sufficient permissions:
#   - Sites.Read.All (to read all SPO sites)
#   - Group.Read.All (to read M365 Groups)
#   - Sites.FullControl.All (for admin access to all sites)

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "VivaEngageSites_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCsv = $true,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 50
)

# 1) Connect to Graph (interactive)
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
$config = Get-Content "C:\temp\GraphDemo\config\clientconfiguration.json" -Raw | ConvertFrom-Json
try {
    # Try device code authentication first (more reliable for automation)
    Connect-MgGraph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint
    Write-Host "Successfully connected to Microsoft Graph using device authentication" -ForegroundColor Green
}
catch {
    Write-Host "Device authentication failed, trying interactive browser..." -ForegroundColor Yellow
    try {
        # Fallback to interactive browser
        Connect-MgGraph -Scopes "Sites.Read.All","Group.Read.All","Sites.FullControl.All" -NoWelcome
        Write-Host "Successfully connected to Microsoft Graph using browser authentication" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        exit 1
    }
}

# Verify connection
$context = Get-MgContext
if (-not $context) {
    Write-Error "Not connected to Microsoft Graph. Please ensure you have the required permissions."
    exit 1
}

# 2) Get all SharePoint sites
Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Yellow
$allSites = @()
$uri = "https://graph.microsoft.com/v1.0/sites?search=*"

do {
    try {
        Write-Host "Calling Graph API: $uri" -ForegroundColor Gray
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        if ($response.value) {
            $allSites += $response.value
            Write-Host "Retrieved $($allSites.Count) sites so far..." -ForegroundColor Gray
        }
        $uri = $response.'@odata.nextLink'
    }
    catch {
        Write-Error "Error retrieving sites: $($_.Exception.Message)"
        Write-Host "Response details: $($_.Exception.Response)" -ForegroundColor Red
        
        # Try alternative approach with Get-MgSite
        Write-Host "Trying alternative approach with Get-MgSite..." -ForegroundColor Yellow
        try {
            $sites = Get-MgSite -All
            $allSites += $sites
            Write-Host "Successfully retrieved $($allSites.Count) sites using Get-MgSite" -ForegroundColor Green
            break
        }
        catch {
            Write-Error "Alternative approach also failed: $($_.Exception.Message)"
            break
        }
    }
} while ($uri)

Write-Host "Found $($allSites.Count) total SharePoint sites" -ForegroundColor Green

# 3) Initialize results collection
$vivaEngageResults = @()
$processedCount = 0
$vivaEngageCount = 0

Write-Host "`nAnalyzing sites for Viva Engage connections..." -ForegroundColor Yellow

# 4) Process each site to check for Viva Engage connection
foreach ($site in $allSites) {
    $processedCount++
    Write-Progress -Activity "Analyzing SharePoint Sites" -Status "Processing site $processedCount of $($allSites.Count)" -PercentComplete (($processedCount / $allSites.Count) * 100)
    
    try {
        # Skip root SharePoint site and search center
        if ($site.root.siteCollection.hostname -match "sharepoint\.com$" -and 
            ($site.displayName -eq "SharePoint" -or $site.displayName -like "*Search*")) {
            continue
        }
        
        $groupId = $site.root.siteCollection.groupId
        $siteResult = [PSCustomObject]@{
            SiteName                    = $site.displayName
            SiteUrl                     = $site.webUrl
            SiteId                      = $site.id
            GroupId                     = $groupId
            GroupDisplayName            = ""
            ResourceProvisioningOptions = ""
            IsVivaEngageCommunity       = $false
            LastModified               = $site.lastModifiedDateTime
            CreatedDateTime            = $site.createdDateTime
        }
        
        if ($groupId) {
            try {
                # Get the Microsoft 365 Group details
                $group = Get-MgGroup -GroupId $groupId -Property "id,displayName,resourceProvisioningOptions" -ErrorAction SilentlyContinue
                
                if ($group) {
                    $hasYammer = $group.resourceProvisioningOptions -contains "Yammer"
                    
                    $siteResult.GroupDisplayName = $group.DisplayName
                    $siteResult.ResourceProvisioningOptions = ($group.resourceProvisioningOptions -join ',')
                    $siteResult.IsVivaEngageCommunity = [bool]$hasYammer
                    
                    if ($hasYammer) {
                        $vivaEngageCount++
                        Write-Host "✓ Found Viva Engage site: $($site.displayName)" -ForegroundColor Green
                    }
                }
            }
            catch {
                Write-Warning "Could not retrieve group details for GroupId: $groupId (Site: $($site.displayName))"
            }
        }
        
        $vivaEngageResults += $siteResult
        
        # Add small delay to avoid throttling
        if ($processedCount % $BatchSize -eq 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    catch {
        Write-Warning "Error processing site '$($site.displayName)': $($_.Exception.Message)"
    }
}

Write-Progress -Activity "Analyzing SharePoint Sites" -Completed

# 5) Display summary results
Write-Host "`n"
$separator = "=" * 60
Write-Host $separator -ForegroundColor Cyan
Write-Host "VIVA ENGAGE COMMUNITY ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host "Total SharePoint sites analyzed: $($vivaEngageResults.Count)" -ForegroundColor White
Write-Host "Sites connected to Viva Engage: $vivaEngageCount" -ForegroundColor Green
Write-Host "Sites NOT connected to Viva Engage: $($vivaEngageResults.Count - $vivaEngageCount)" -ForegroundColor Yellow
Write-Host $separator -ForegroundColor Cyan

# 6) Display Viva Engage connected sites
$vivaEngageSites = $vivaEngageResults | Where-Object { $_.IsVivaEngageCommunity -eq $true }

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

# 7) Export results to CSV if requested
if ($ExportToCsv) {
    try {
        $vivaEngageResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Green
        
        # Also create a filtered file with only Viva Engage sites
        if ($vivaEngageSites.Count -gt 0) {
            $vivaEngageOnlyPath = $OutputPath -replace '\.csv$', '_VivaEngageOnly.csv'
            $vivaEngageSites | Export-Csv -Path $vivaEngageOnlyPath -NoTypeInformation -Encoding UTF8
            Write-Host "Viva Engage sites only exported to: $vivaEngageOnlyPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# 8) Return summary object
$summary = [PSCustomObject]@{
    TotalSitesAnalyzed = $vivaEngageResults.Count
    VivaEngageSites = $vivaEngageCount
    NonVivaEngageSites = $vivaEngageResults.Count - $vivaEngageCount
    AnalysisDate = Get-Date
    OutputFile = if ($ExportToCsv) { $OutputPath } else { "Not exported" }
    VivaEngageConnectedSites = $vivaEngageSites
}

Write-Host "`nAnalysis completed successfully!" -ForegroundColor Green
return $summary