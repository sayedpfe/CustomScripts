<#
.SYNOPSIS
    Bulk import multiple device identifiers to Intune using Microsoft Graph API

.DESCRIPTION
    This script imports multiple corporate device identifiers (IMEI, Serial Number, or Manufacturer+Model+Serial)
    to Microsoft Intune using the Microsoft Graph API beta endpoint in a single bulk operation.

.PARAMETER CsvPath
    Path to CSV file containing device identifiers. CSV should have columns: ImportedDeviceIdentifier, ImportedDeviceIdentityType, Description

.PARAMETER Devices
    Array of device objects to import. Each object should have: ImportedDeviceIdentifier, ImportedDeviceIdentityType, Description

.PARAMETER OverwriteExisting
    If true, overwrites existing device identities with the same identifier

.PARAMETER TenantId
    Azure AD Tenant ID (required for app registration authentication)

.PARAMETER ClientId
    App Registration Client ID (Application ID)

.PARAMETER ClientSecret
    App Registration Client Secret

.PARAMETER UseInteractive
    Use interactive authentication instead of app registration

.EXAMPLE
    # Import from CSV with app registration
    .\Import-BulkDeviceIdentities.ps1 -CsvPath "devices.csv" -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret"

.EXAMPLE
    # Import array of devices
    $devices = @(
        @{ ImportedDeviceIdentifier = "123456789012345"; ImportedDeviceIdentityType = "imei"; Description = "iPhone 15" },
        @{ ImportedDeviceIdentifier = "SN-ABC123"; ImportedDeviceIdentityType = "serialNumber"; Description = "MacBook Pro" }
    )
    .\Import-BulkDeviceIdentities.ps1 -Devices $devices -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret"

.EXAMPLE
    # Interactive authentication
    .\Import-BulkDeviceIdentities.ps1 -CsvPath "devices.csv" -UseInteractive

.NOTES
    Requires Microsoft.Graph.Authentication module
    
    Required API Permissions:
    - DeviceManagementServiceConfig.ReadWrite.All (Application permission)
#>

[CmdletBinding(DefaultParameterSetName = 'InteractiveCsv')]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'InteractiveCsv')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationCsv')]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'InteractiveArray')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationArray')]
    [array]$Devices,
    
    [Parameter(Mandatory = $false)]
    [switch]$OverwriteExisting,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationCsv')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationArray')]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationCsv')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationArray')]
    [string]$ClientId,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationCsv')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistrationArray')]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false, ParameterSetName = 'InteractiveCsv')]
    [Parameter(Mandatory = $false, ParameterSetName = 'InteractiveArray')]
    [switch]$UseInteractive
)

# Function to check if Microsoft.Graph module is installed
function Test-GraphModule {
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
    if (-not $module) {
        Write-Warning "Microsoft.Graph.Authentication module is not installed."
        Write-Host "Install it using: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Yellow
        return $false
    }
    return $true
}

# Main script execution
try {
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Intune Bulk Device Identity Import" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    
    # Check if Graph module is available
    if (-not (Test-GraphModule)) {
        throw "Required module not found. Please install Microsoft.Graph.Authentication module."
    }
    
    # Import the module
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    
    # Load devices from CSV or use provided array
    $deviceList = @()
    
    if ($PSCmdlet.ParameterSetName -like '*Csv*') {
        Write-Host "`n[1/4] Loading devices from CSV..." -ForegroundColor Yellow
        
        if (-not (Test-Path $CsvPath)) {
            throw "CSV file not found: $CsvPath"
        }
        
        $csvDevices = Import-Csv $CsvPath
        Write-Host "  ✓ Found $($csvDevices.Count) devices in CSV" -ForegroundColor Green
        
        # Validate CSV columns
        $requiredColumns = @('ImportedDeviceIdentifier', 'ImportedDeviceIdentityType')
        $csvColumns = $csvDevices[0].PSObject.Properties.Name
        
        foreach ($col in $requiredColumns) {
            if ($col -notin $csvColumns) {
                throw "CSV is missing required column: $col"
            }
        }
        
        $deviceList = $csvDevices
    }
    else {
        Write-Host "`n[1/4] Loading devices from array..." -ForegroundColor Yellow
        Write-Host "  ✓ Received $($Devices.Count) devices" -ForegroundColor Green
        $deviceList = $Devices
    }
    
    # Connect to Microsoft Graph
    Write-Host "`n[2/4] Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    if ($PSCmdlet.ParameterSetName -like 'AppRegistration*') {
        # App Registration (Service Principal) Authentication
        Write-Host "  Using App Registration authentication..." -ForegroundColor Cyan
        
        # Convert client secret to secure string if it's a plain string
        if ($ClientSecret -is [string]) {
            $secureClientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        } else {
            $secureClientSecret = $ClientSecret
        }
        
        # Create credential object
        $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secureClientSecret)
        
        # Connect using client credentials
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop
        
        Write-Host "  ✓ Connected using App Registration" -ForegroundColor Green
        Write-Host "    Tenant ID: $TenantId" -ForegroundColor Gray
        Write-Host "    Client ID: $ClientId" -ForegroundColor Gray
    }
    else {
        # Interactive (Delegated) Authentication
        Write-Host "  Using Interactive authentication..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "DeviceManagementServiceConfig.ReadWrite.All" -ErrorAction Stop
        
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to connect to Microsoft Graph"
        }
        Write-Host "  ✓ Connected as: $($context.Account)" -ForegroundColor Green
    }
    
    # Prepare the request body
    Write-Host "`n[3/4] Preparing device identities..." -ForegroundColor Yellow
    
    $importedDeviceIdentities = @()
    $validTypes = @('imei', 'serialNumber', 'manufacturerModelSerial')
    
    foreach ($device in $deviceList) {
        # Validate identity type
        if ($device.ImportedDeviceIdentityType -notin $validTypes) {
            Write-Warning "  Skipping device '$($device.ImportedDeviceIdentifier)' - Invalid type: $($device.ImportedDeviceIdentityType)"
            continue
        }
        
        $deviceObj = @{
            "@odata.type" = "#microsoft.graph.importedDeviceIdentity"
            "importedDeviceIdentifier" = $device.ImportedDeviceIdentifier
            "importedDeviceIdentityType" = $device.ImportedDeviceIdentityType
        }
        
        # Add description if provided
        if ($device.Description -and -not [string]::IsNullOrWhiteSpace($device.Description)) {
            $deviceObj["description"] = $device.Description
        }
        
        $importedDeviceIdentities += $deviceObj
    }
    
    if ($importedDeviceIdentities.Count -eq 0) {
        throw "No valid devices to import"
    }
    
    Write-Host "  ✓ Prepared $($importedDeviceIdentities.Count) device identities for import" -ForegroundColor Green
    
    # Create request body - ensure proper serialization
    $bodyObj = @{
        importedDeviceIdentities = $importedDeviceIdentities
        overwriteImportedDeviceIdentities = [bool]$OverwriteExisting.IsPresent
    }
    
    $body = $bodyObj | ConvertTo-Json -Depth 10 -Compress:$false
    
    Write-Host "`nRequest Details:" -ForegroundColor Cyan
    Write-Host "  Endpoint: POST https://graph.microsoft.com/beta/deviceManagement/importedDeviceIdentities/importDeviceIdentityList"
    Write-Host "  Device Count: $($importedDeviceIdentities.Count)" -ForegroundColor White
    Write-Host "  Overwrite Existing: $($OverwriteExisting.IsPresent)" -ForegroundColor White
    Write-Host "`nRequest Body:" -ForegroundColor Gray
    Write-Host $body -ForegroundColor DarkGray
    
    # Make the API request
    Write-Host "`n[4/4] Submitting bulk import request..." -ForegroundColor Yellow
    
    $uri = "https://graph.microsoft.com/beta/deviceManagement/importedDeviceIdentities/importDeviceIdentityList"
    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
    
    # Display success message
    Write-Host "`n✓ Bulk import completed successfully!" -ForegroundColor Green
    
    # Parse and display results
    if ($response.value) {
        $successCount = ($response.value | Where-Object { $_.status -eq $true }).Count
        $failCount = ($response.value | Where-Object { $_.status -ne $true }).Count
        
        Write-Host "`n═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  Import Summary" -ForegroundColor Cyan
        Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  Total Devices: $($response.value.Count)" -ForegroundColor White
        Write-Host "  ✓ Successful: $successCount" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  ✗ Failed: $failCount" -ForegroundColor Red
        }
        
        Write-Host "`nDetailed Results:" -ForegroundColor Cyan
        $response.value | ForEach-Object {
            $statusIcon = if ($_.status) { "✓" } else { "✗" }
            $statusColor = if ($_.status) { "Green" } else { "Red" }
            
            Write-Host "  $statusIcon " -ForegroundColor $statusColor -NoNewline
            Write-Host "$($_.importedDeviceIdentifier) " -NoNewline
            Write-Host "($($_.importedDeviceIdentityType))" -ForegroundColor Gray -NoNewline
            
            if ($_.description) {
                Write-Host " - $($_.description)" -ForegroundColor Gray
            } else {
                Write-Host ""
            }
        }
    }
    
    Write-Host "`n═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    
    return $response
}
catch {
    Write-Host "`n✗ Error occurred during bulk import:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($_.Exception.Response) {
        Write-Host "`nHTTP Status Code: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
    }
    
    # Provide troubleshooting tips
    Write-Host "`nTroubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have the required permission: DeviceManagementServiceConfig.ReadWrite.All"
    Write-Host "2. Verify all device identifiers are in the correct format"
    Write-Host "3. Check CSV file format (columns: ImportedDeviceIdentifier, ImportedDeviceIdentityType, Description)"
    Write-Host "4. Ensure you're connected to the correct tenant"
    
    throw
}
finally {
    # Optional: Disconnect from Graph
    # Uncomment the line below if you want to disconnect after execution
    # Disconnect-MgGraph
}
