<#
.SYNOPSIS
    Export SharePoint site structure and lists for Offboarding Copilot Agent

.DESCRIPTION
    This script exports the complete SharePoint site structure including:
    - Site information
    - All lists and libraries
    - List schemas (columns, content types, views)
    - List data (optional)
    
.PARAMETER SiteUrl
    The URL of the SharePoint site to export
    
.PARAMETER ExportPath
    The path where export files will be saved
    
.PARAMETER IncludeData
    Switch to include list item data in the export
    
.PARAMETER Credential
    Optional credentials for authentication

.EXAMPLE
    .\Export-SharePointSiteStructure.ps1 -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" -ExportPath ".\Export"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess",
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".\Export",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeData,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

# Install PnP PowerShell if not already installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell -ErrorAction Stop

# Create export directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportFolder = Join-Path $ExportPath "SharePointExport_$timestamp"
New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null

Write-Host "Starting SharePoint site export..." -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl" -ForegroundColor White
Write-Host "Export Path: $exportFolder" -ForegroundColor White

try {
    # Load config from file if ClientId not provided
    $clientSecret = $null
    $tenantId = $null
    
    if (-not $ClientId) {
        $configPath = Join-Path $PSScriptRoot "AppConfig.json"
        if (Test-Path $configPath) {
            Write-Host "Loading configuration from AppConfig.json..." -ForegroundColor Gray
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            $ClientId = $config.ClientId
            $clientSecret = $config.ClientSecret
            $tenantId = $config.TenantId
        }
    }
    
    # Connect to SharePoint
    Write-Host "`nConnecting to SharePoint site..." -ForegroundColor Yellow
    
    if ($Credential) {
        Write-Host "Using provided credentials..." -ForegroundColor Gray
        Connect-PnPOnline -Url $SiteUrl -Credentials $Credential
    } elseif ($ClientId -and $clientSecret) {
        # Use app-only authentication with client secret
        # Note: This requires registering as SharePoint ACS app or adding Site.Selected permission
        Write-Host "Using app-only authentication..." -ForegroundColor Gray
        Write-Host "Note: If this fails, the app needs DELEGATED permissions for interactive auth" -ForegroundColor Gray
        
        try {
            # Try Azure AD app with client secret (for Sites.Selected permission)
            $secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force
            $tenantName = $tenantId -replace '\.onmicrosoft\.com$', ''
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $clientSecret
        } catch {
            Write-Host "App-only auth failed. Trying interactive mode..." -ForegroundColor Yellow
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
        }
    } elseif ($ClientId) {
        # Use interactive authentication (requires DELEGATED permissions)
        Write-Host "Using interactive authentication (browser window will open)..." -ForegroundColor Gray
        Write-Host "This requires DELEGATED SharePoint permissions in the app registration" -ForegroundColor Gray
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    } else {
        Write-Host "`nNo authentication method available. Please either:" -ForegroundColor Yellow
        Write-Host "1. Run Setup-EntraIDApp.ps1 to register an app, OR" -ForegroundColor White
        Write-Host "2. Provide -ClientId parameter" -ForegroundColor White
        throw "Authentication configuration required"
    }
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    
    # Export site information
    Write-Host "`n=== Exporting Site Information ===" -ForegroundColor Cyan
    $web = Get-PnPWeb -Includes Title, Description, ServerRelativeUrl, Created, Language
    
    $siteInfo = @{
        Title = $web.Title
        Description = $web.Description
        ServerRelativeUrl = $web.ServerRelativeUrl
        Created = $web.Created
        Language = $web.Language
        SiteUrl = $SiteUrl
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    $siteInfoPath = Join-Path $exportFolder "SiteInfo.json"
    $siteInfo | ConvertTo-Json -Depth 10 | Out-File $siteInfoPath -Encoding UTF8
    Write-Host "Site information exported to: SiteInfo.json" -ForegroundColor Green
    
    # Export using PnP Provisioning Template
    Write-Host "`n=== Using PnP Provisioning Engine ===" -ForegroundColor Cyan
    Write-Host "Extracting site template (this may take a few minutes)..." -ForegroundColor Yellow
    
    $templatePath = Join-Path $exportFolder "SiteTemplate.xml"
    $pnpTemplatePath = Join-Path $exportFolder "SiteTemplate.pnp"
    
    # Extract template with handlers for all components
    $handlers = "Lists", "Fields", "ContentTypes", "CustomActions", "Features", "Navigation", "PropertyBagEntries", "Publishing", "RegionalSettings", "SearchSettings", "SitePolicy", "SupportedUILanguages", "TermGroups", "Workflows", "SiteSecurity", "ComposedLook", "Tenant", "WebSettings", "SiteFooter", "Theme", "ApplicationLifecycleManagement"
    
    Get-PnPSiteTemplate -Out $templatePath -Handlers $handlers -IncludeAllPages -IncludeHiddenLists:$false -PersistBrandingFiles -PersistPublishingFiles
    
    Write-Host "PnP Provisioning Template exported to: SiteTemplate.xml" -ForegroundColor Green
    
    # Also save as PnP file (compressed format)
    Save-PnPSiteTemplate -Out $pnpTemplatePath -Template (Read-PnPSiteTemplate -Path $templatePath)
    Write-Host "Compressed template saved to: SiteTemplate.pnp" -ForegroundColor Green
    
    # Export list data if requested
    if ($IncludeData) {
        Write-Host "`n=== Exporting List Data ===" -ForegroundColor Cyan
        
        # Get all lists (excluding system lists)
        $lists = Get-PnPList | Where-Object { 
            -not $_.Hidden -and 
            $_.BaseTemplate -ne 109 -and
            $_.Title -notlike "appdata" -and
            $_.Title -notlike "appfiles"
        }
        
        foreach ($list in $lists) {
            Write-Host "Exporting data from: $($list.Title)..." -ForegroundColor Gray
            
            $items = Get-PnPListItem -List $list.Title -PageSize 1000
    foreach ($list in $lists) {
        $listsExport += @{
            Title = $list.Title
            Description = $list.Description
            BaseTemplate = $list.BaseTemplate
            ItemCount = $list.ItemCount
        }
    }
    
    # Export all lists summary
    $listsPath = Join-Path $exportFolder "AllLists.json"
    $listsExport | ConvertTo-Json -Depth 10 | Out-File $listsPath -Encoding UTF8
    Write-Host "Lists summary| ConvertTo-Json -Depth 10 | Out-File $listDataPath -Encoding UTF8
                Write-Host "  - Exported $($items.Count) items" -ForegroundColor Green
                
                # Also export to CSV for easy viewing
                if ($itemsData.Count -gt 0) {
                    $csvPath = Join-Path $exportFolder "List_$($list.Title -replace '[^\w\-]', '_')_Data.csv"
                    $itemsData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                }
            } else {
                Write-Host "  - No items to export" -ForegroundColor Gray
            }
        }
        
        Write-Host "  - List exported successfully" -ForegroundColor Green
    }
    
    # Export all lists summary
    $listsPath = Join-Path $exportFolder "AllLists.json"
    $listsExport | ConvertTo-Json -Depth 10 | Out-File $listsPath -Encoding UTF8
    Write-Host "`nAll lists exported to: AllLists.json" -ForegroundColor Green
    
    # Create deployment manifest
    Write-Host "`n=== Creating Deployment Manifest ===" -ForegroundColor Cyan
    $manifest = @{
        SiteInfo = $siteInfo
        Lists = $listsExport | Select-Object Title, Description, BaseTemplate, ItemCount
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalLists = $listsExport.Count
        IncludeData = $IncludeData.IsPresent
    }
    
    $manifestPath = Join-Path $exportFolder "DeploymentManifest.json"
    $manifest | ConvertTo-Json -Depth 10 | Out-File $manifestPath -Encoding UTF8
    Write-Host "Deployment manifest created: DeploymentManifest.json" -ForegroundColor Green
    
    # Create summary report
    $summaryPath = Join-Path $exportFolder "ExportSummary.txt"
    $exportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $summaryText = New-Object System.Text.StringBuilder
    [void]$summaryText.AppendLine("SharePoint Site Export Summary")
    [void]$summaryText.AppendLine("==============================")
    [void]$summaryText.AppendLine("Export Date: $exportDate")
    [void]$summaryText.AppendLine("Site URL: $SiteUrl")
    [void]$summaryText.AppendLine("Site Title: $($siteInfo.Title)")
    [void]$summaryText.AppendLine("")
    [void]$summaryText.AppendLine("Lists Exported: $($listsExport.Count)")
    [void]$summaryText.AppendLine("")
    [void]$summaryText.AppendLine("List Details:")
    
    foreach ($listItem in $listsExport) {
        [void]$summaryText.AppendLine("  - $($listItem.Title) - $($listItem.ItemCount) items")
    }
    
    [void]$summaryText.AppendLine("")
    [void]$summaryText.AppendLine("Export Location: $exportFolder")
    $summaryText.ToString() | Out-File $summaryPath -Encoding UTF8
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Export completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Export location: $exportFolder" -ForegroundColor White
    Write-Host "Total lists exported: $($listsExport.Count)" -ForegroundColor White
    Write-Host "`nFiles created:" -ForegroundColor Yellow
    Get-ChildItem -Path $exportFolder | ForEach-Object { 
        Write-Host "  - $($_.Name)" -ForegroundColor Gray 
    }
    
} catch {
    Write-Host "`nError during export: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
} finally {
    Disconnect-PnPOnline
    Write-Host "`nDisconnected from SharePoint" -ForegroundColor Gray
}

Write-Host "`nExport script completed." -ForegroundColor Cyan
