# Azure Automation Runbook - Apply SharePoint Conditional Access Policy
# Uses Microsoft.Online.SharePoint.PowerShell with Set-SPOSite cmdlet
# Runtime: PowerShell 7.2
# 
# Prerequisites:
# 1. Create Automation Variable: SPOTenantName (e.g., "M365CPI90282478")
# 2. Create Automation Credential: SPOAdminCredential (SharePoint Admin username/password)
# 3. Import Module: Microsoft.Online.SharePoint.PowerShell

param(
    [Parameter(Mandatory=$false)]
    [object]$WebhookData,
    
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$AuthenticationContextName
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Type = "INFO")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Type) {
        "SUCCESS" { "Green" }
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        default { "White" }
    }
    
    Write-Output "[$timestamp] [$Type] $Message"
    Write-Host $Message -ForegroundColor $color
}

try {
    Write-Log "========================================" "INFO"
    Write-Log "Apply Conditional Access Policy - Set-SPOSite" "INFO"
    Write-Log "========================================" "INFO"
    
    # Parse webhook data if provided, otherwise use direct parameters
    if ($WebhookData) {
        Write-Log "Webhook triggered - parsing input data..." "INFO"
        
        $requestBody = $WebhookData.RequestBody
        if ($requestBody) {
            $inputData = $requestBody | ConvertFrom-Json
            $SiteUrl = $inputData.SiteUrl
            $AuthenticationContextName = $inputData.AuthenticationContextName
        } else {
            throw "No request body received from webhook"
        }
    }

    Write-Log "Parameters:" "INFO"
    Write-Log "  Site URL: $SiteUrl" "INFO"
    Write-Log "  Auth Context: $AuthenticationContextName" "INFO"

    # Validate inputs
    if (-not $SiteUrl -or -not $AuthenticationContextName) {
        throw "Missing required parameters: SiteUrl and AuthenticationContextName are required"
    }

    # Get automation variables
    Write-Log "`nRetrieving configuration..." "INFO"
    $tenantName = Get-AutomationVariable -Name 'SPOTenantName'
    
    if (-not $tenantName) {
        throw "Automation variable 'SPOTenantName' not found. Please create it with value like 'M365CPI90282478' (without .onmicrosoft.com)"
    }
    
    $adminUrl = "https://$tenantName-admin.sharepoint.com"
    Write-Log "✅ Admin URL: $adminUrl" "SUCCESS"
    
    # Get stored credential
    Write-Log "`nRetrieving stored credentials..." "INFO"
    $credential = Get-AutomationPSCredential -Name 'SPOAdminCredential'
    
    if (-not $credential) {
        throw "Credential 'SPOAdminCredential' not found in Automation Account. Please create a credential asset with SharePoint Admin username and password."
    }
    
    Write-Log "✅ Credential retrieved for: $($credential.UserName)" "SUCCESS"
    
    # Import SharePoint Online module
    Write-Log "`nImporting Microsoft.Online.SharePoint.PowerShell module..." "INFO"
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
    Write-Log "✅ Module imported successfully" "SUCCESS"
    
    # Connect to SharePoint Online Admin Center
    Write-Log "`nConnecting to SharePoint Online Admin Center..." "INFO"
    
    try {
        Connect-SPOService -Url $adminUrl -Credential $credential
        Write-Log "✅ Connected successfully" "SUCCESS"
    } catch {
        throw "Failed to connect to SharePoint: $($_.Exception.Message)"
    }
    
    # Apply conditional access policy
    Write-Log "`nApplying conditional access policy..." "INFO"
    
    try {
        # Use Set-SPOSite with both ConditionalAccessPolicy and AuthenticationContextName parameters
        # This is the official Microsoft-documented method
        Set-SPOSite -Identity $SiteUrl `
                    -ConditionalAccessPolicy AuthenticationContext `
                    -AuthenticationContextName $AuthenticationContextName
        
        Write-Log "✅ Conditional access policy applied successfully!" "SUCCESS"
        Write-Log "   Policy Type: AuthenticationContext" "SUCCESS"
        Write-Log "   Context Name: $AuthenticationContextName" "SUCCESS"
        
    } catch {
        throw "Failed to apply policy: $($_.Exception.Message)"
    }
    
    # Wait for propagation
    Write-Log "`nWaiting for changes to propagate..." "INFO"
    Start-Sleep -Seconds 3
    
    # Verify the changes
    Write-Log "`nVerifying changes..." "INFO"
    
    try {
        $verifiedSite = Get-SPOSite -Identity $SiteUrl -Detailed
        
        Write-Log "`nVerification Results:" "INFO"
        Write-Log "  Site URL: $($verifiedSite.Url)" "INFO"
        Write-Log "  Conditional Access Policy: $($verifiedSite.ConditionalAccessPolicy)" "INFO"
        
        if ($verifiedSite.ConditionalAccessPolicy -eq "AuthenticationContext") {
            Write-Log "`n🎉 SUCCESS! Conditional access policy is active!" "SUCCESS"
            
            $result = @{
                Success = $true
                SiteUrl = $SiteUrl
                PolicyType = $verifiedSite.ConditionalAccessPolicy
                AuthenticationContextName = $AuthenticationContextName
                Message = "Conditional access policy applied successfully"
                Timestamp = (Get-Date).ToString("o")
            }
        } else {
            Write-Log "⚠️ Policy may not have been applied correctly" "WARNING"
            
            $result = @{
                Success = $false
                SiteUrl = $SiteUrl
                PolicyType = $verifiedSite.ConditionalAccessPolicy
                Message = "Policy applied but verification shows unexpected value"
                Timestamp = (Get-Date).ToString("o")
            }
        }
        
    } catch {
        Write-Log "Warning: Could not verify changes: $($_.Exception.Message)" "WARNING"
        
        $result = @{
            Success = $true
            SiteUrl = $SiteUrl
            Message = "Policy applied but verification failed"
            Timestamp = (Get-Date).ToString("o")
        }
    }
    
    Write-Log "`n========================================" "INFO"
    Write-Log "Runbook execution completed successfully" "SUCCESS"
    Write-Log "========================================" "INFO"
    
    # Disconnect
    Disconnect-SPOService
    
    # Return result as JSON
    return ($result | ConvertTo-Json)
    
} catch {
    Write-Log "`n========================================" "ERROR"
    Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
    Write-Log "========================================" "ERROR"
    
    # Disconnect if connected
    try { Disconnect-SPOService } catch {}
    
    $errorResult = @{
        Success = $false
        SiteUrl = $SiteUrl
        Error = $_.Exception.Message
        ErrorDetails = $_.Exception.ToString()
        Timestamp = (Get-Date).ToString("o")
    }
    
    throw ($errorResult | ConvertTo-Json)
}
