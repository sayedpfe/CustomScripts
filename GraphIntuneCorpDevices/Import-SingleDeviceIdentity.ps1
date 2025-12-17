<#
.SYNOPSIS
    Import a single device identifier to Intune using Microsoft Graph API

.DESCRIPTION
    This script imports a single corporate device identifier (IMEI, Serial Number, or Manufacturer+Model+Serial)
    to Microsoft Intune using the Microsoft Graph API beta endpoint.
    
    Supports two authentication methods:
    1. Interactive (delegated) authentication
    2. App Registration (service principal) authentication

.PARAMETER ImportedDeviceIdentifier
    The device identifier value (IMEI, Serial Number, etc.)

.PARAMETER ImportedDeviceIdentityType
    The type of identifier. Valid values: imei, serialNumber, manufacturerModelSerial

.PARAMETER Description
    Optional description for the device

.PARAMETER TenantId
    Azure AD Tenant ID (required for app registration authentication)

.PARAMETER ClientId
    App Registration Client ID (Application ID)

.PARAMETER ClientSecret
    App Registration Client Secret

.PARAMETER UseInteractive
    Use interactive authentication instead of app registration

.EXAMPLE
    # Interactive authentication
    .\Import-SingleDeviceIdentity.ps1 -ImportedDeviceIdentifier "123456789012345" -ImportedDeviceIdentityType "imei" -Description "Test device" -UseInteractive

.EXAMPLE
    # App registration with client secret
    .\Import-SingleDeviceIdentity.ps1 -ImportedDeviceIdentifier "123456789012345" -ImportedDeviceIdentityType "imei" -Description "Test device" -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret"

.EXAMPLE
    # App registration with stored credentials
    $secureSecret = ConvertTo-SecureString "your-client-secret" -AsPlainText -Force
    .\Import-SingleDeviceIdentity.ps1 -ImportedDeviceIdentifier "123456789012345" -ImportedDeviceIdentityType "imei" -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret $secureSecret

.NOTES
    Requires Microsoft.Graph.Authentication module
    
    Required API Permissions (Application permissions for app registration):
    - DeviceManagementServiceConfig.ReadWrite.All (Required)
      OR
    - DeviceManagementServiceConfiguration.ReadWrite.All (Alternative)
    
    See APP_REGISTRATION_SETUP.md for detailed setup instructions
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(Mandatory = $true)]
    [string]$ImportedDeviceIdentifier,
    
    [Parameter(Mandatory = $true)]
    [ValidateSet("imei", "serialNumber", "manufacturerModelSerial")]
    [string]$ImportedDeviceIdentityType,
    
    [Parameter(Mandatory = $false)]
    [string]$Description = "",
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]$ClientId,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
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
    Write-Host "Starting device identity import process..." -ForegroundColor Cyan
    
    # Check if Graph module is available
    if (-not (Test-GraphModule)) {
        throw "Required module not found. Please install Microsoft.Graph.Authentication module."
    }
    
    # Import the module
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    
    # Connect to Microsoft Graph based on authentication method
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    if ($PSCmdlet.ParameterSetName -eq 'AppRegistration') {
        # App Registration (Service Principal) Authentication
        Write-Host "Using App Registration authentication..." -ForegroundColor Cyan
        
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
        
        Write-Host "✓ Connected to Microsoft Graph using App Registration" -ForegroundColor Green
        Write-Host "  Tenant ID: $TenantId" -ForegroundColor Gray
        Write-Host "  Client ID: $ClientId" -ForegroundColor Gray
    }
    else {
        # Interactive (Delegated) Authentication
        Write-Host "Using Interactive authentication..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All" -ErrorAction Stop
        
        # Verify connection
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to connect to Microsoft Graph"
        }
        Write-Host "✓ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
    }
    
    # Prepare the request body
    $body = @{
        "@odata.type" = "#microsoft.graph.importedDeviceIdentity"
        "importedDeviceIdentifier" = $ImportedDeviceIdentifier
        "importedDeviceIdentityType" = $ImportedDeviceIdentityType
    }
    
    # Add description if provided
    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        $body["description"] = $Description
    }
    
    # Convert to JSON
    $jsonBody = $body | ConvertTo-Json -Depth 10
    
    Write-Host "`nRequest Details:" -ForegroundColor Cyan
    Write-Host "Endpoint: POST https://graph.microsoft.com/beta/deviceManagement/importedDeviceIdentities"
    Write-Host "Body:" -ForegroundColor Cyan
    Write-Host $jsonBody -ForegroundColor Gray
    
    # Make the API request
    Write-Host "`nSubmitting device identity import..." -ForegroundColor Yellow
    
    # Note: This endpoint requires the importDeviceIdentityList action
    # For single device import, we need to use the bulk import endpoint with a single item
    $uri = "https://graph.microsoft.com/beta/deviceManagement/importedDeviceIdentities/importDeviceIdentityList"
    
    # Wrap in array for bulk import endpoint with required parameter
    $bulkBody = @{
        "importedDeviceIdentities" = @($body)
        "overwriteImportedDeviceIdentities" = $false  # Don't overwrite existing devices
    } | ConvertTo-Json -Depth 10
    
    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $bulkBody -ContentType "application/json" -ErrorAction Stop
    
    # Display success message
    Write-Host "`n✓ Device identity imported successfully!" -ForegroundColor Green
    Write-Host "`nResponse Details:" -ForegroundColor Cyan
    $response | ConvertTo-Json -Depth 10 | Write-Host -ForegroundColor Gray
    
    # Display key information
    if ($response.id) {
        Write-Host "`nImported Device ID: $($response.id)" -ForegroundColor Green
    }
    if ($response.importedDeviceIdentifier) {
        Write-Host "Device Identifier: $($response.importedDeviceIdentifier)" -ForegroundColor Green
    }
    if ($response.importedDeviceIdentityType) {
        Write-Host "Identity Type: $($response.importedDeviceIdentityType)" -ForegroundColor Green
    }
    
    return $response
}
catch {
    Write-Host "`n✗ Error occurred during device import:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($_.Exception.Response) {
        Write-Host "`nHTTP Status Code: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
    }
    
    # Provide troubleshooting tips
    Write-Host "`nTroubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have the required permission: DeviceManagementManagedDevices.ReadWrite.All"
    Write-Host "2. Verify the device identifier format is correct for the identity type"
    Write-Host "3. Check if the device is already imported"
    Write-Host "4. Ensure you're connected to the correct tenant"
    
    throw
}
finally {
    # Optional: Disconnect from Graph
    # Uncomment the line below if you want to disconnect after execution
    # Disconnect-MgGraph
}
