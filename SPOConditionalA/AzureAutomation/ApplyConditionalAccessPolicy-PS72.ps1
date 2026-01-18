# PowerShell 7.2 Runbook for Conditional Access Policy
# Uses PnP.PowerShell with Managed Identity - NO CERTIFICATE NEEDED!

param(
    [Parameter(Mandatory=$false)]
    [object]$WebhookData
)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "[$timestamp] [$Level] $Message"
}

Write-Log "========================================" "INFO"
Write-Log "Conditional Access Policy Application" "INFO"
Write-Log "========================================" "INFO"

try {
    # Import PnP.PowerShell module
    Write-Log "Importing PnP.PowerShell module..." "INFO"
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Log "✅ Module imported successfully" "SUCCESS"

    # Parse webhook data
    if ($WebhookData) {
        Write-Log "Webhook triggered - parsing input data..." "INFO"
        
        $requestBody = $WebhookData.RequestBody
        if ($requestBody) {
            $inputData = $requestBody | ConvertFrom-Json
        } else {
            throw "No request body received from webhook"
        }
    } else {
        throw "No webhook data received. This runbook must be triggered via webhook."
    }

    # Extract parameters
    $siteUrl = $inputData.SiteUrl
    $authContextName = $inputData.AuthenticationContextName
    $requestorEmail = $inputData.RequestorEmail
    $requestId = $inputData.RequestId

    Write-Log "Input Parameters:" "INFO"
    Write-Log "  Site URL: $siteUrl" "INFO"
    Write-Log "  Auth Context: $authContextName" "INFO"
    Write-Log "  Requestor: $requestorEmail" "INFO"
    Write-Log "  Request ID: $requestId" "INFO"

    # Validate inputs
    if (-not $siteUrl -or -not $authContextName) {
        throw "Missing required parameters: SiteUrl and AuthenticationContextName are required"
    }

    # Get SharePoint admin URL from site URL
    if ($siteUrl -match "https://([^.]+)\.sharepoint\.com") {
        $tenantName = $Matches[1]
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
    } else {
        throw "Invalid SharePoint site URL format: $siteUrl"
    }

    Write-Log "Admin URL: $adminUrl" "INFO"
    
    # Connect to SharePoint using Managed Identity
    Write-Log "Connecting to SharePoint with Managed Identity..." "INFO"
    Connect-PnPOnline -Url $adminUrl -ManagedIdentity
    Write-Log "✅ Connected to SharePoint" "SUCCESS"

    # Apply conditional access policy
    Write-Log "Applying conditional access policy..." "INFO"
    
    # Use Set-PnPTenantSite - the PnP equivalent of Set-SPOSite
    Set-PnPTenantSite -Identity $siteUrl `
                      -ConditionalAccessPolicy AuthenticationContext `
                      -ErrorAction Stop

    Write-Log "✅ Policy set to AuthenticationContext" "INFO"
    
    # Now set the authentication context name using CSOM
    Write-Log "Setting authentication context name: $authContextName" "INFO"
    $site = Get-PnPSite -Identity $siteUrl
    $site.AuthenticationContextName = $authContextName
    $site.Update()
    $site.Context.ExecuteQuery()

    Write-Log "✅ Policy application command executed" "SUCCESS"

    # Wait a moment for changes to propagate
    Start-Sleep -Seconds 3

    # Verify the change
    Write-Log "Verifying policy application..." "INFO"
    $site = Get-PnPSite -Identity $siteUrl
    
    if ($site.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Log "✅ SUCCESS! Conditional Access Policy verified" "SUCCESS"
        Write-Log "  Policy Type: $($site.ConditionalAccessPolicy)" "INFO"
        
        # Disconnect
        Disconnect-PnPOnline
        
        # Return success result
        $result = @{
            Success = $true
            Message = "Conditional Access Policy successfully applied to $siteUrl"
            SiteUrl = $siteUrl
            Policy = "AuthenticationContext"
            AuthContext = $authContextName
            RequestId = $requestId
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        Write-Log "========================================" "INFO"
        Write-Output ($result | ConvertTo-Json)
        
    } else {
        throw "Policy application verification failed. Expected 'AuthenticationContext', got '$($site.ConditionalAccessPolicy)'"
    }

} catch {
    Write-Log "❌ ERROR occurred" "ERROR"
    Write-Log "Error Message: $($_.Exception.Message)" "ERROR"
    Write-Log "Error Details: $($_.Exception.ToString())" "ERROR"
    
    # Try to disconnect if connected
    try { Disconnect-PnPOnline } catch {}
    
    # Return error result
    $errorResult = @{
        Success = $false
        Message = "Failed to apply Conditional Access Policy"
        Error = $_.Exception.Message
        SiteUrl = $siteUrl
        RequestId = $requestId
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
    
    Write-Log "========================================" "INFO"
    Write-Output ($errorResult | ConvertTo-Json)
    
    throw
}
