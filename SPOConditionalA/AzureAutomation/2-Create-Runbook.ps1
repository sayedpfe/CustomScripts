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
    This runbook applies a conditional access policy to a SharePoint site.
    It is triggered by Power Automate via webhook.

.PARAMETER WebhookData
    Data passed from Power Automate webhook

.NOTES
    Requires: PnP.PowerShell module
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
    # Check if runbook exists
    $existingRunbook = Get-AzAutomationRunbook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $runbookName `
        -ErrorAction SilentlyContinue
    
    if ($existingRunbook) {
        Write-Host "   ℹ️  Runbook already exists, updating..." -ForegroundColor Yellow
        
        # Import updated content
        Import-AzAutomationRunbook `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name $runbookName `
            -Type PowerShell `
            -Path ([IO.Path]::GetTempFileName()) `
            -Force `
            -Published | Out-Null
        
    } else {
        Write-Host "   Creating new runbook..." -ForegroundColor Yellow
        
        # Save runbook content to temp file
        $tempFile = [IO.Path]::GetTempFileName()
        $runbookContent | Out-File -FilePath $tempFile -Encoding UTF8
        
        # Import runbook
        Import-AzAutomationRunbook `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name $runbookName `
            -Type PowerShell `
            -Path $tempFile `
            -Published | Out-Null
        
        Remove-Item $tempFile
    }
    
    # Update runbook content
    $tempFile = [IO.Path]::GetTempFileName()
    $runbookContent | Out-File -FilePath $tempFile -Encoding UTF8
    
    Set-AzAutomationRunbookContent `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $runbookName `
        -Path $tempFile `
        -Force | Out-Null
    
    Remove-Item $tempFile
    
    # Publish the runbook
    Publish-AzAutomationRunbook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $runbookName | Out-Null
    
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
