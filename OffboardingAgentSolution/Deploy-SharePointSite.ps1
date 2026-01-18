<#
.SYNOPSIS
    Deploy SharePoint site structure and lists from exported configuration

.DESCRIPTION
    This script creates a new SharePoint site or updates an existing site with:
    - Site configuration
    - Lists and libraries
    - List schemas (columns, content types, views)
    - List data (optional)
    
.PARAMETER TargetSiteUrl
    The URL of the SharePoint site where the structure will be deployed
    
.PARAMETER ExportPath
    The path containing the exported configuration files
    
.PARAMETER CreateSite
    Switch to create a new site (requires admin permissions)
    
.PARAMETER IncludeData
    Switch to import list item data from the export
    
.PARAMETER Credential
    Optional credentials for authentication

.EXAMPLE
    .\Deploy-SharePointSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewOffboarding" -ExportPath ".\Export\SharePointExport_20241218_120000"
    
.EXAMPLE
    .\Deploy-SharePointSite.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewOffboarding" -ExportPath ".\Export\SharePointExport_20241218_120000" -IncludeData
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetSiteUrl,
    
    [Parameter(Mandatory = $true)]
    [string]$ExportPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$CreateSite,
    
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

Write-Host "Starting SharePoint site deployment..." -ForegroundColor Cyan
Write-Host "Target Site URL: $TargetSiteUrl" -ForegroundColor White
Write-Host "Export Path: $ExportPath" -ForegroundColor White

# Validate export path
if (-not (Test-Path $ExportPath)) {
    Write-Host "Error: Export path not found: $ExportPath" -ForegroundColor Red
    exit 1
}

$manifestPath = Join-Path $ExportPath "DeploymentManifest.json"
if (-not (Test-Path $manifestPath)) {
    Write-Host "Error: DeploymentManifest.json not found in export path" -ForegroundColor Red
    exit 1
}

try {
    # Load manifest
    Write-Host "`nLoading deployment manifest..." -ForegroundColor Yellow
    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $siteInfoPath = Join-Path $ExportPath "SiteInfo.json"
    $siteInfo = Get-Content $siteInfoPath -Raw | ConvertFrom-Json
    
    Write-Host "Manifest loaded successfully" -ForegroundColor Green
    Write-Host "  Original Site: $($siteInfo.Title)" -ForegroundColor Gray
    Write-Host "  Lists to deploy: $($manifest.TotalLists)" -ForegroundColor Gray
    
    # Load config from file if ClientId not provided
    $clientSecret = $null
    $tenantId = $null
    
    if (-not $ClientId) {
        $configPath = Join-Path (Split-Path $ExportPath -Parent) "AppConfig.json"
        if (-not (Test-Path $configPath)) {
            $configPath = Join-Path $PSScriptRoot "AppConfig.json"
        }
        if (Test-Path $configPath) {
            Write-Host "Loading configuration from AppConfig.json..." -ForegroundColor Gray
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            $ClientId = $config.ClientId
            $clientSecret = $config.ClientSecret
            $tenantId = $config.TenantId
        }
    }
    
    # Connect to SharePoint
    Write-Host "`nConnecting to target SharePoint site..." -ForegroundColor Yellow
    
    if ($Credential) {
        Write-Host "Using provided credentials..." -ForegroundColor Gray
        Connect-PnPOnline -Url $TargetSiteUrl -Credentials $Credential
    } elseif ($ClientId -and $clientSecret) {
        # Use app-only authentication with client secret
        Write-Host "Using app-only authentication..." -ForegroundColor Gray
        Write-Host "Note: If this fails, the app needs DELEGATED permissions for interactive auth" -ForegroundColor Gray
        
        try {
            # Try Azure AD app with client secret
            Connect-PnPOnline -Url $TargetSiteUrl -ClientId $ClientId -ClientSecret $clientSecret
        } catch {
            Write-Host "App-only auth failed. Trying interactive mode..." -ForegroundColor Yellow
            Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId
        }
    } elseif ($ClientId) {
        # Use interactive authentication (requires DELEGATED permissions)
        Write-Host "Using interactive authentication (browser window will open)..." -ForegroundColor Gray
        Write-Host "This requires DELEGATED SharePoint permissions in the app registration" -ForegroundColor Gray
        Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId
    } else {
        Write-Host "`nNo authentication method available. Please either:" -ForegroundColor Yellow
        Write-Host "1. Run Setup-EntraIDApp.ps1 to register an app, OR" -ForegroundColor White
        Write-Host "2. Provide -ClientId parameter" -ForegroundColor White
        throw "Authentication configuration required"
    }
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    
    # Apply PnP Provisioning Template
    Write-Host "`n=== Applying PnP Provisioning Template ===" -ForegroundColor Cyan
    
    # Check for template files
    $templatePath = Join-Path $ExportPath "SiteTemplate.xml"
    $pnpTemplatePath = Join-Path $ExportPath "SiteTemplate.pnp"
    
    $templateFile = $null
    if (Test-Path $pnpTemplatePath) {
        $templateFile = $pnpTemplatePath
        Write-Host "Using PnP template: SiteTemplate.pnp" -ForegroundColor Gray
    } elseif (Test-Path $templatePath) {
        $templateFile = $templatePath
        Write-Host "Using XML template: SiteTemplate.xml" -ForegroundColor Gray
    } else {
        Write-Host "Warning: No PnP template found, falling back to manual deployment" -ForegroundColor Yellow
    }
    
    if ($templateFile) {
        Write-Host "Applying provisioning template (this may take a few minutes)..." -ForegroundColor Yellow
        
        try {
            # Apply the template with handlers
            $handlers = "Lists", "Fields", "ContentTypes", "CustomActions", "Features", "Navigation", "PropertyBagEntries", "Publishing", "RegionalSettings", "SearchSettings", "SitePolicy", "SupportedUILanguages", "TermGroups", "Workflows", "SiteSecurity", "ComposedLook", "WebSettings", "SiteFooter", "Theme"
            
            Apply-PnPProvisioningTemplate -Path $templateFile -Handlers $handlers -ClearNavigation
            
            Write-Host "Template applied successfully!" -ForegroundColor Green
            
            # Import list data if requested
            if ($IncludeData) {
                Write-Host "`n=== Importing List Data ===" -ForegroundColor Cyan
                
                $dataFiles = Get-ChildItem -Path $ExportPath -Filter "ListData_*.json"
                
                foreach ($dataFile in $dataFiles) {
                    $listName = $dataFile.BaseName -replace '^ListData_', '' -replace '_', ' '
                    Write-Host "Importing data to: $listName..." -ForegroundColor Gray
                    
                    try {
                        $listData = Get-Content $dataFile.FullName -Raw | ConvertFrom-Json
                        
                        $importedCount = 0
                        foreach ($item in $listData) {
                            try {
                                $values = @{}
                                $item.FieldValues.PSObject.Properties | ForEach-Object {
                                    if ($_.Name -ne "ID" -and 
                                        $_.Name -ne "Created" -and 
                                        $_.Name -ne "Modified" -and
                                        $_.Name -ne "Author" -and
                                        $_.Name -ne "Editor" -and
                                        $_.Value -ne $null) {
                                        $values[$_.Name] = $_.Value
                                    }
                                }
                                
                                if ($values.Count -gt 0) {
                                    Add-PnPListItem -List $listName -Values $values -ErrorAction Stop | Out-Null
                                    $importedCount++
                                }
                            } catch {
                                Write-Host "    Warning: Could not import item: $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                        }
                        
                        Write-Host "  Imported $importedCount items" -ForegroundColor Green
                    } catch {
                        Write-Host "  Warning: Could not import data for $listName" -ForegroundColor Yellow
                    }
                }
            }
            
            $deployedLists = $manifest.Lists.Title
            
        } catch {
            Write-Host "Error applying template: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Falling back to manual deployment..." -ForegroundColor Yellow
        }
    }
    
    # Manual deployment fallback (if needed)
    if (-not $templateFile -or $deployedLists.Count -eq 0) {
        Write-Host "`n=== Manual List Deployment ===" -ForegroundColor Cyan
        $deployedLists = @()
        $listCount = 0
        
        foreach ($listInfo in $manifest.Lists) {
        $listCount++
        Write-Host "`nDeploying List ($listCount/$($manifest.TotalLists)): $($listInfo.Title)" -ForegroundColor Yellow
        
        # Load full list schema
        $listSchemaPath = Join-Path $ExportPath "List_$($listInfo.Title -replace '[^\w\-]', '_')_Schema.json"
        
        if (-not (Test-Path $listSchemaPath)) {
            Write-Host "  Warning: Schema file not found for list '$($listInfo.Title)', skipping..." -ForegroundColor Yellow
            continue
        }
        
        $listSchema = Get-Content $listSchemaPath -Raw | ConvertFrom-Json
        
        # Check if list already exists
        $existingList = Get-PnPList -Identity $listSchema.Title -ErrorAction SilentlyContinue
        
        if ($existingList) {
            Write-Host "  List already exists. Updating schema..." -ForegroundColor Yellow
            $list = $existingList
        } else {
            Write-Host "  Creating new list..." -ForegroundColor Gray
            
            # Determine template type
            $template = switch ($listSchema.BaseTemplate) {
                100 { "GenericList" }
                101 { "DocumentLibrary" }
                106 { "Events" }
                107 { "Tasks" }
                108 { "DiscussionBoard" }
                109 { "PictureLibrary" }
                119 { "WebPageLibrary" }
                default { "GenericList" }
            }
            
            # Create the list
            $list = New-PnPList -Title $listSchema.Title -Template $template -ErrorAction Stop
            
            # Update list settings
            Set-PnPList -Identity $listSchema.Title `
                -Description $listSchema.Description `
                -EnableVersioning $listSchema.EnableVersioning `
                -EnableMinorVersions $listSchema.EnableMinorVersions `
                -EnableModeration $listSchema.EnableModeration `
                -EnableAttachments $listSchema.EnableAttachments `
                -EnableFolderCreation $listSchema.EnableFolderCreation `
                -ErrorAction SilentlyContinue
            
            Write-Host "  - List created successfully" -ForegroundColor Green
        }
        
        # Add custom fields
        if ($listSchema.Fields -and $listSchema.Fields.Count -gt 0) {
            Write-Host "  - Adding custom fields ($($listSchema.Fields.Count))..." -ForegroundColor Gray
            
            foreach ($field in $listSchema.Fields) {
                try {
                    # Check if field already exists
                    $existingField = Get-PnPField -List $listSchema.Title -Identity $field.InternalName -ErrorAction SilentlyContinue
                    
                    if ($existingField) {
                        Write-Host "    * Field '$($field.InternalName)' already exists, skipping" -ForegroundColor DarkGray
                        continue
                    }
                    
                    # Create field based on type
                    $fieldParams = @{
                        List = $listSchema.Title
                        DisplayName = $field.Title
                        InternalName = $field.InternalName
                        Type = $field.TypeAsString
                        Required = $field.Required
                    }
                    
                    if ($field.Description) {
                        $fieldParams.Description = $field.Description
                    }
                    
                    # Handle choice fields
                    if ($field.TypeAsString -eq "Choice" -and $field.Choices) {
                        $fieldParams.Choices = $field.Choices
                        if ($field.DefaultValue) {
                            $fieldParams.DefaultValue = $field.DefaultValue
                        }
                    } elseif ($field.DefaultValue) {
                        $fieldParams.DefaultValue = $field.DefaultValue
                    }
                    
                    Add-PnPField @fieldParams -ErrorAction Stop
                    Write-Host "    * Added field: $($field.InternalName)" -ForegroundColor DarkGreen
                    
                } catch {
                    Write-Host "    * Warning: Could not add field '$($field.InternalName)': $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        
        # Create custom views
        if ($listSchema.Views -and $listSchema.Views.Count -gt 0) {
            Write-Host "  - Creating custom views..." -ForegroundColor Gray
            
            foreach ($view in $listSchema.Views) {
                try {
                    if ($view.Title -ne "All Items" -and $view.Title -ne "All Documents") {
                        $existingView = Get-PnPView -List $listSchema.Title -Identity $view.Title -ErrorAction SilentlyContinue
                        
                        if (-not $existingView) {
                            $viewFields = $view.ViewFields | Where-Object { $_ -ne $null }
                            
                            Add-PnPView -List $listSchema.Title `
                                -Title $view.Title `
                                -Fields $viewFields `
                                -Query $view.ViewQuery `
                                -RowLimit $view.RowLimit `
                                -SetAsDefault:$view.DefaultView `
                                -ErrorAction SilentlyContinue
                            
                            Write-Host "    * Added view: $($view.Title)" -ForegroundColor DarkGreen
                        }
                    }
                } catch {
                    Write-Host "    * Warning: Could not add view '$($view.Title)': $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        
        # Import list data if requested
        if ($IncludeData) {
            $listDataPath = Join-Path $ExportPath "List_$($listInfo.Title -replace '[^\w\-]', '_')_Data.json"
            
            if (Test-Path $listDataPath) {
                Write-Host "  - Importing list data..." -ForegroundColor Gray
                $listData = Get-Content $listDataPath -Raw | ConvertFrom-Json
                
                if ($listData.Count -gt 0) {
                    $importedCount = 0
                    foreach ($itemData in $listData) {
                        try {
                            # Convert PSCustomObject to hashtable
                            $values = @{}
                            $itemData.PSObject.Properties | ForEach-Object {
                                if ($_.Name -ne "ID" -and 
                                    $_.Name -ne "Created" -and 
                                    $_.Name -ne "Modified" -and
                                    $_.Name -ne "Author" -and
                                    $_.Name -ne "Editor" -and
                                    $_.Value -ne $null) {
                                    $values[$_.Name] = $_.Value
                                }
                            }
                            
                            if ($values.Count -gt 0) {
                                Add-PnPListItem -List $listSchema.Title -Values $values -ErrorAction Stop | Out-Null
                                $importedCount++
                            }
                        } catch {
                            Write-Host "    * Warning: Could not import item: $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }
                    
                    Write-Host "  - Imported $importedCount of $($listData.Count) items" -ForegroundColor Green
                } else {
                    Write-Host "  - No data to import" -ForegroundColor Gray
                }
            }
        }
        
            $deployedLists += $listSchema.Title
            Write-Host "  - List deployed successfully" -ForegroundColor Green
        }
    }
    
    # Create deployment log
    $logPath = Join-Path $ExportPath "DeploymentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $log = @"
SharePoint Site Deployment Log
==============================
Deployment Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Target Site URL: $TargetSiteUrl
Source Export: $ExportPath

Lists Deployed: $($deployedLists.Count)

List Details:
"@
    
    foreach ($listName in $deployedLists) {
        $log += "`n  - $listName"
    }
    
    $log | Out-File $logPath -Encoding UTF8
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Deployment completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Target site: $TargetSiteUrl" -ForegroundColor White
    Write-Host "Total lists deployed: $($deployedLists.Count)" -ForegroundColor White
    Write-Host "`nDeployed lists:" -ForegroundColor Yellow
    $deployedLists | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
    
} catch {
    Write-Host "`nError during deployment: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
} finally {
    Disconnect-PnPOnline
    Write-Host "`nDisconnected from SharePoint" -ForegroundColor Gray
}

Write-Host "`nDeployment script completed." -ForegroundColor Cyan
