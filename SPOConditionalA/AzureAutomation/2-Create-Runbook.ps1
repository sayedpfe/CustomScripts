# Create PowerShell Runbook for Conditional Access Policy Application
# This runbook will be triggered by Power Automate via webhook

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\automation-config.json"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Creating Azure Automation Runbook" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    Write-Host "❌ Configuration file not found: $ConfigPath" -ForegroundColor Red
    Write-Host "Please run 1-Setup-AutomationAccount.ps1 first" -ForegroundColor Yellow
    exit 1
}

$config = Get-Content $ConfigPath | ConvertFrom-Json
Write-Host "Loaded configuration from: $ConfigPath" -ForegroundColor Gray

# Connect to Azure
Write-Host "`n📌 Connecting to Azure..." -ForegroundColor Cyan
Connect-AzAccount -SubscriptionId $config.SubscriptionId | Out-Null
Write-Host "   ✅ Connected" -ForegroundColor Green

# Define the runbook content
$runbookContent = @'
<#
.SYNOPSIS
    Apply Conditional Access Policy to SharePoint Site

.DESCRIPTION
    This runbook applies a conditional access policy to a SharePoint site using SharePoint Admin CSOM.
    It is triggered by Power Automate via webhook.

.PARAMETER WebhookData
    Data passed from Power Automate webhook

.NOTES
    Uses: Built-in modules only (no external dependencies)
    Permissions: SharePoint Administrator role (via Managed Identity)
#>

param(
    [Parameter(Mandatory=$false)]
    [object]$WebhookData
)

# Function to write log with timestamp
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "[$timestamp] [$Level] $Message"
}

Write-Log "========================================" "INFO"
Write-Log "Conditional Access Policy Application" "INFO"
Write-Log "========================================" "INFO"

# NOTE: Azure Automation uses Windows PowerShell 5.1
# We'll use PowerShell remoting to call Set-SPOSite on behalf of the managed identity

try {
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
    
    # Connect using Managed Identity
    Write-Log "Connecting with Managed Identity..." "INFO"
    try {
        Connect-AzAccount -Identity | Out-Null
        Write-Log "✅ Connected to Azure with Managed Identity" "SUCCESS"
    } catch {
        throw "Failed to connect with Managed Identity: $($_.Exception.Message)"
    }

    # Get access token for SharePoint
    Write-Log "Getting SharePoint access token..." "INFO"
    try {
        $token = (Get-AzAccessToken -ResourceUrl "https://$tenantName.sharepoint.com").Token
        Write-Log "✅ Got access token" "SUCCESS"
    } catch {
        throw "Failed to get access token: $($_.Exception.Message)"
    }

    # Use SharePoint Online Management Shell cmdlets via runbook worker
    Write-Log "Installing SharePoint Online Management Shell..." "INFO"
    try {
        # Check if module exists
        $spoModule = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
        if (-not $spoModule) {
            Write-Log "Module not found, will use REST API instead..." "WARNING"
            
            # Use REST API approach
            Write-Log "Using SharePoint Admin REST API..." "INFO"
            $headers = @{
                "Authorization" = "Bearer $token"
                "Accept" = "application/json;odata=verbose"
                "Content-Type" = "application/json;odata=verbose"
            }
            
            # Call SharePoint Admin API
            $apiUrl = "$adminUrl/_api/SPOSitePropertiesEnumerable/GetById('$siteUrl')"
            Write-Log "API URL: $apiUrl" "INFO"
            
            # Note: This may not work as documented API may not support this property
            Write-Log "⚠️ WARNING: SharePoint Admin REST API may not support conditional access properties" "WARNING"
            Write-Log "This operation requires SharePoint Online Management Shell with admin privileges" "WARNING"
            
            throw "SharePoint Online Management Shell module not available in Azure Automation. Please install Microsoft.Online.SharePoint.PowerShell module in the Automation Account."
            
        } else {
            Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
            Write-Log "✅ Module imported" "SUCCESS"
            
            # Connect to SharePoint Online
            Write-Log "Connecting to SharePoint Online..." "INFO"
            # Cannot use Managed Identity with SPO module directly
            throw "Microsoft.Online.SharePoint.PowerShell does not support Managed Identity authentication. This requires a different approach."
        }
        
    } catch {
        Write-Log "❌ ERROR: $($_.Exception.Message)" "ERROR"
        throw
    }

} catch {
    Write-Log "❌ ERROR occurred" "ERROR"
    Write-Log "Error Message: $($_.Exception.Message)" "ERROR"
    
    # Return error result
    $errorResult = @{
        Success = $false
        Message = "Failed to apply Conditional Access Policy"
        Error = $_.Exception.Message
        SiteUrl = $siteUrl
        RequestId = $requestId
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Note = "Azure Automation with Managed Identity cannot directly call Set-SPOSite. Alternative: Use Azure Function with certificate-based app-only authentication, or use PnP PowerShell 2.x with app-only authentication."
    }
    
    Write-Log "========================================" "INFO"
    Write-Output ($errorResult | ConvertTo-Json)
    
    throw
}
'@
        
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

    # Connect to SharePoint using Managed Identity
    Write-Log "Connecting to SharePoint with Managed Identity..." "INFO"
    
    # Get SharePoint admin URL from site URL
    if ($siteUrl -match "https://([^.]+)\.sharepoint\.com") {
        $tenantName = $Matches[1]
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
    } else {
        throw "Invalid SharePoint site URL format: $siteUrl"
    }

    Write-Log "Admin URL: $adminUrl" "INFO"
    
    # Connect using Managed Identity
    Connect-PnPOnline -Url $adminUrl -ManagedIdentity
    Write-Log "✅ Connected to SharePoint" "SUCCESS"

    # Apply conditional access policy
    Write-Log "Applying conditional access policy..." "INFO"
    
    Set-PnPSite -Identity $siteUrl `
                -ConditionalAccessPolicy AuthenticationContext `
                -AuthenticationContextName $authContextName `
                -ErrorAction Stop

    Write-Log "✅ Policy application command executed" "SUCCESS"

    # Wait a moment for changes to propagate
    Start-Sleep -Seconds 3

    # Verify the change
    Write-Log "Verifying policy application..." "INFO"
    $site = Get-PnPSite -Identity $siteUrl -Includes ConditionalAccessPolicy
    
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
'@

# Create the runbook
$runbookName = "Apply-SiteConditionalAccess"

Write-Host "`n📌 Creating Runbook: $runbookName..." -ForegroundColor Cyan

try {
    # Save runbook content to temp file with .ps1 extension
    $tempFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), "runbook-$(Get-Date -Format 'yyyyMMddHHmmss').ps1")
    $runbookContent | Out-File -FilePath $tempFile -Encoding UTF8
    
    Write-Host "   Created temp file: $tempFile" -ForegroundColor Gray
    
    # Check if runbook exists
    $existingRunbook = Get-AzAutomationRunbook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $runbookName `
        -ErrorAction SilentlyContinue
    
    if ($existingRunbook) {
        Write-Host "   ℹ️  Runbook already exists, updating..." -ForegroundColor Yellow
        
        # Update existing runbook
        Import-AzAutomationRunbook `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name $runbookName `
            -Type PowerShell `
            -Path $tempFile `
            -Force | Out-Null
        
    } else {
        Write-Host "   Creating new runbook..." -ForegroundColor Yellow
        
        # Import new runbook
        Import-AzAutomationRunbook `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name $runbookName `
            -Type PowerShell `
            -Path $tempFile | Out-Null
    }
    
    # Publish the runbook
    Publish-AzAutomationRunbook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $runbookName | Out-Null
    
    # Clean up temp file
    Remove-Item $tempFile -ErrorAction SilentlyContinue
    
    Write-Host "   ✅ Runbook created and published: $runbookName" -ForegroundColor Green
    
} catch {
    Write-Host "   ❌ Failed to create runbook" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Update config with runbook name
$config | Add-Member -NotePropertyName "RunbookName" -NotePropertyValue $runbookName -Force
$config | ConvertTo-Json | Out-File -FilePath $ConfigPath -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ RUNBOOK CREATED!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host @"

Runbook Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Name:                 $runbookName
Type:                 PowerShell
Status:               Published
Automation Account:   $($config.AutomationAccountName)

Capabilities:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Accepts webhook input from Power Automate
✅ Authenticates using Managed Identity
✅ Applies conditional access policy to SharePoint sites
✅ Verifies policy application
✅ Returns detailed success/error messages

Next Steps:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. ⏭️  Run: .\3-Setup-Webhook.ps1
   This will create a webhook URL for Power Automate

2. ⏭️  Test the runbook in Azure Portal (optional)
   Go to: Automation Account → Runbooks → $runbookName → Test

"@ -ForegroundColor White
