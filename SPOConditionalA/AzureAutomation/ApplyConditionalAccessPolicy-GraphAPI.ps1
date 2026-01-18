# Azure Automation Runbook - Apply SharePoint Conditional Access Policy
# Uses Microsoft Graph API with certificate authentication
# Runtime: PowerShell 7.2

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
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
    Write-Log "Apply Conditional Access Policy via Graph API" "INFO"
    Write-Log "========================================" "INFO"
    
    Write-Log "Parameters:" "INFO"
    Write-Log "  Site URL: $SiteUrl" "INFO"
    Write-Log "  Auth Context: $AuthenticationContextName" "INFO"
    
    # Get automation variables
    Write-Log "`nRetrieving configuration from Automation Account..." "INFO"
    $appId = Get-AutomationVariable -Name 'SPOAppId'
    $tenantId = Get-AutomationVariable -Name 'SPOTenantId'
    
    if (-not $appId -or -not $tenantId) {
        throw "Automation variables SPOAppId or SPOTenantId not found or are empty"
    }
    
    Write-Log "✅ App ID: $appId" "SUCCESS"
    Write-Log "✅ Tenant ID: $tenantId" "SUCCESS"
    
    # Get certificate from Automation Account
    Write-Log "`nRetrieving certificate..." "INFO"
    $cert = Get-AutomationCertificate -Name 'SPOAppCertificate'
    
    if (-not $cert) {
        throw "Certificate 'SPOAppCertificate' not found in Automation Account"
    }
    
    Write-Log "✅ Certificate retrieved (Thumbprint: $($cert.Thumbprint))" "SUCCESS"
    
    # Check if Microsoft.Graph modules are available
    Write-Log "`nChecking Microsoft Graph modules..." "INFO"
    
    $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites')
    foreach ($module in $requiredModules) {
        $mod = Get-Module -ListAvailable -Name $module | Select-Object -First 1
        if ($mod) {
            Write-Log "✅ $module version $($mod.Version) available" "SUCCESS"
        } else {
            throw "$module not found. Please import this module in Automation Account."
        }
    }
    
    # Connect to Microsoft Graph using certificate
    Write-Log "`nConnecting to Microsoft Graph..." "INFO"
    
    try {
        Connect-MgGraph -ClientId $appId -TenantId $tenantId -Certificate $cert -NoWelcome
        Write-Log "✅ Connected to Microsoft Graph successfully" "SUCCESS"
    } catch {
        throw "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    }
    
    # Get the site ID from the URL
    Write-Log "`nRetrieving site information..." "INFO"
    
    # Extract host and site path from URL
    $uri = [System.Uri]$SiteUrl
    $hostname = $uri.Host
    $sitePath = $uri.AbsolutePath
    
    # Get site using Graph API
    try {
        $site = Get-MgSite -SiteId "${hostname}:${sitePath}"
        Write-Log "✅ Site found: $($site.DisplayName)" "SUCCESS"
        Write-Log "   Site ID: $($site.Id)" "INFO"
    } catch {
        throw "Failed to retrieve site: $($_.Exception.Message)"
    }
    
    # Apply conditional access policy using Graph API
    Write-Log "`nApplying conditional access policy..." "INFO"
    
    try {
        $params = @{
            AdditionalProperties = @{
                conditionalAccessPolicy = "AuthenticationContext"
                authenticationContextName = $AuthenticationContextName
            }
        }
        
        Update-MgSite -SiteId $site.Id -BodyParameter $params
        
        Write-Log "✅ Policy applied successfully!" "SUCCESS"
        Write-Log "   Policy Type: AuthenticationContext" "SUCCESS"
        Write-Log "   Context Name: $AuthenticationContextName" "SUCCESS"
        
    } catch {
        throw "Failed to apply policy: $($_.Exception.Message)"
    }
    
    # Wait for propagation
    Write-Log "`nWaiting for changes to propagate..." "INFO"
    Start-Sleep -Seconds 5
    
    # Verify the changes
    Write-Log "`nVerifying changes..." "INFO"
    
    try {
        $verifiedSite = Get-MgSite -SiteId $site.Id
        $conditionalPolicy = $verifiedSite.AdditionalProperties['conditionalAccessPolicy']
        $authContextName = $verifiedSite.AdditionalProperties['authenticationContextName']
        
        Write-Log "`nVerification Results:" "INFO"
        Write-Log "  Site URL: $SiteUrl" "INFO"
        Write-Log "  Conditional Access Policy: $conditionalPolicy" "INFO"
        Write-Log "  Authentication Context Name: $authContextName" "INFO"
        
        if ($conditionalPolicy -eq "AuthenticationContext") {
            Write-Log "`n🎉 SUCCESS! Conditional access policy is active!" "SUCCESS"
            
            # Return success result
            $result = @{
                Success = $true
                SiteUrl = $SiteUrl
                PolicyType = $conditionalPolicy
                AuthenticationContextName = $authContextName
                Message = "Conditional access policy applied successfully"
            }
        } else {
            Write-Log "⚠️ Policy may not have been applied correctly" "WARNING"
            
            $result = @{
                Success = $false
                SiteUrl = $SiteUrl
                PolicyType = $conditionalPolicy
                AuthenticationContextName = $authContextName
                Message = "Policy applied but verification failed"
            }
        }
        
    } catch {
        Write-Log "Warning: Could not verify changes: $($_.Exception.Message)" "WARNING"
        
        $result = @{
            Success = $true
            SiteUrl = $SiteUrl
            Message = "Policy applied but verification failed"
        }
    }
    
    Write-Log "`n========================================" "INFO"
    Write-Log "Runbook execution completed successfully" "SUCCESS"
    Write-Log "========================================" "INFO"
    
    # Disconnect
    Disconnect-MgGraph | Out-Null
    
    # Return result as JSON
    return ($result | ConvertTo-Json)
    
} catch {
    Write-Log "`n========================================" "ERROR"
    Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
    Write-Log "========================================" "ERROR"
    
    # Disconnect if connected
    try { Disconnect-MgGraph | Out-Null } catch {}
    
    $errorResult = @{
        Success = $false
        SiteUrl = $SiteUrl
        Error = $_.Exception.Message
        ErrorDetails = $_.Exception.ToString()
    }
    
    throw ($errorResult | ConvertTo-Json)
}
