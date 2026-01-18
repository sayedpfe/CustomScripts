# Create Runbook with Microsoft.Online.SharePoint.PowerShell
# This version uses certificate-based app authentication

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\automation-config.json"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Creating Updated Runbook (SPO Module)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Load configuration
$config = Get-Content $ConfigPath | ConvertFrom-Json
Write-Host "Loaded configuration" -ForegroundColor Gray

# Connect to Azure
Write-Host "`nConnecting to Azure..." -ForegroundColor Cyan
Connect-AzAccount -SubscriptionId $config.SubscriptionId | Out-Null
Write-Host "✅ Connected" -ForegroundColor Green

# Define the runbook content using SPO module with certificate
$runbookContent = @'
<#
.SYNOPSIS
    Apply Conditional Access Policy to SharePoint Site

.DESCRIPTION
    Uses Microsoft.Online.SharePoint.PowerShell with certificate authentication.
    Certificate must be uploaded to Automation Account.

.PARAMETER WebhookData
    Data from Power Automate webhook

.NOTES
    Requires: Microsoft.Online.SharePoint.PowerShell module
    Requires: Certificate uploaded to Automation Account as "SPOAppCertificate"
    Permissions: SharePoint Administrator role on Service Principal
#>

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
    # Import module
    Write-Log "Importing Microsoft.Online.SharePoint.PowerShell module..." "INFO"
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
    Write-Log "✅ Module imported" "SUCCESS"

    # Parse webhook data
    if ($WebhookData) {
        Write-Log "Parsing webhook data..." "INFO"
        $requestBody = $WebhookData.RequestBody
        if ($requestBody) {
            $inputData = $requestBody | ConvertFrom-Json
        } else {
            throw "No request body received"
        }
    } else {
        throw "No webhook data received"
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
        throw "Missing required parameters"
    }

    # Get admin URL
    if ($siteUrl -match "https://([^.]+)\.sharepoint\.com") {
        $tenantName = $Matches[1]
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
    } else {
        throw "Invalid SharePoint site URL format: $siteUrl"
    }

    Write-Log "Admin URL: $adminUrl" "INFO"

    # Get certificate from Automation Account
    Write-Log "Retrieving certificate..." "INFO"
    $cert = Get-AutomationCertificate -Name "SPOAppCertificate"
    if (-not $cert) {
        throw "Certificate 'SPOAppCertificate' not found in Automation Account"
    }
    Write-Log "✅ Certificate retrieved" "SUCCESS"

    # Get App ID from Automation Variable
    $appId = Get-AutomationVariable -Name "SPOAppId"
    if (-not $appId) {
        throw "Variable 'SPOAppId' not found in Automation Account"
    }
    Write-Log "App ID: $appId" "INFO"

    # Connect to SharePoint Online using certificate
    Write-Log "Connecting to SharePoint Online with app-only authentication..." "INFO"
    Connect-SPOService -Url $adminUrl -ClientId $appId -Certificate $cert
    Write-Log "✅ Connected to SharePoint Online" "SUCCESS"

    # Apply conditional access policy
    Write-Log "Applying conditional access policy..." "INFO"
    Set-SPOSite -Identity $siteUrl `
                -ConditionalAccessPolicy AuthenticationContext `
                -AuthenticationContextName $authContextName `
                -ErrorAction Stop

    Write-Log "✅ Policy application command executed" "SUCCESS"

    # Wait for propagation
    Start-Sleep -Seconds 3

    # Verify
    Write-Log "Verifying policy application..." "INFO"
    $site = Get-SPOSite -Identity $siteUrl
    
    if ($site.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Log "✅ SUCCESS! Policy verified" "SUCCESS"
        
        # Disconnect
        Disconnect-SPOService
        
        # Return success
        $result = @{
            Success = $true
            Message = "Conditional Access Policy successfully applied"
            SiteUrl = $siteUrl
            Policy = "AuthenticationContext"
            AuthContext = $authContextName
            RequestId = $requestId
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        Write-Log "========================================" "INFO"
        Write-Output ($result | ConvertTo-Json)
        
    } else {
        throw "Policy verification failed"
    }

} catch {
    Write-Log "❌ ERROR: $($_.Exception.Message)" "ERROR"
    
    # Try to disconnect
    try { Disconnect-SPOService } catch {}
    
    # Return error
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

# Save and import runbook
$runbookName = "Apply-SiteConditionalAccess"
$tempFile = [IO.Path]::Combine([IO.Path]::GetTempPath(), "runbook-spo-$(Get-Date -Format 'yyyyMMddHHmmss').ps1")
$runbookContent | Out-File -FilePath $tempFile -Encoding UTF8

Write-Host "`nImporting runbook..." -ForegroundColor Cyan

Import-AzAutomationRunbook `
    -ResourceGroupName $config.ResourceGroupName `
    -AutomationAccountName $config.AutomationAccountName `
    -Name $runbookName `
    -Type PowerShell `
    -Path $tempFile `
    -Force | Out-Null

Publish-AzAutomationRunbook `
    -ResourceGroupName $config.ResourceGroupName `
    -AutomationAccountName $config.AutomationAccountName `
    -Name $runbookName | Out-Null

Remove-Item $tempFile -ErrorAction SilentlyContinue

Write-Host "✅ Runbook updated and published" -ForegroundColor Green

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "NEXT STEPS REQUIRED:" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host @"

The runbook is ready but needs certificate-based authentication setup:

1. ⏭️  Run: .\Setup-CertificateAuth.ps1
   This will:
   - Create Azure AD App Registration
   - Generate self-signed certificate
   - Grant SharePoint permissions
   - Upload certificate to Automation Account
   - Store App ID as Automation Variable

2. ⏭️  Test the runbook again

Note: This is the ONLY way to make it work with Azure Automation + PS 5.1

"@ -ForegroundColor White
