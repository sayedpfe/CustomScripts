# SharePoint Conditional Access Policy Setup - Site Owner Methods
# This script provides multiple approaches for Site Owners to handle conditional access policies
# without requiring Global Administrator privileges

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AuthenticationContextName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("PnP", "REST", "PowerPlatform", "RequestWorkflow", "All")]
    [string]$Method = "All"
)

# Function to check current user permissions
function Test-SitePermissions {
    param([string]$Url)
    
    try {
        Connect-PnPOnline -Url $Url -Interactive -WarningAction SilentlyContinue
        $currentUser = Get-PnPCurrentUser
        $userRole = Get-PnPUserPermissions -User $currentUser.Email
        
        Write-Host "Current user: $($currentUser.Title)" -ForegroundColor Green
        Write-Host "User permissions: $($userRole -join ', ')" -ForegroundColor Green
        
        return $userRole -contains "FullControl" -or $userRole -contains "SiteCollectionAdministrator"
    }
    catch {
        Write-Warning "Could not verify permissions: $($_.Exception.Message)"
        return $false
    }
}

# Method 1: PnP PowerShell with Site Collection Admin rights
function Set-ConditionalAccessPnP {
    param([string]$Url, [string]$ContextName)
    
    Write-Host "`n=== Method 1: PnP PowerShell ===" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $Url -Interactive
        
        # Check if the command is available
        if (Get-Command Set-PnPSite -ErrorAction SilentlyContinue) {
            Set-PnPSite -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName $ContextName
            Write-Host "✅ Conditional access policy applied successfully using PnP PowerShell" -ForegroundColor Green
            return $true
        } else {
            Write-Warning "Set-PnPSite with conditional access parameters not available in current PnP version"
            return $false
        }
    }
    catch {
        Write-Error "PnP method failed: $($_.Exception.Message)"
        return $false
    }
}

# Method 2: SharePoint REST API
function Set-ConditionalAccessREST {
    param([string]$Url, [string]$ContextName)
    
    Write-Host "`n=== Method 2: REST API ===" -ForegroundColor Cyan
    
    try {
        # Get access token (requires PnP connection)
        $accessToken = Get-PnPAccessToken
        
        $headers = @{
            "Authorization" = "Bearer $accessToken"
            "Accept" = "application/json;odata=verbose"
            "Content-Type" = "application/json;odata=verbose"
            "X-RequestDigest" = (Get-PnPRequestDigest).GetAwaiter().GetResult()
        }

        $body = @{
            "__metadata" = @{ "type" = "SP.Site" }
            "ConditionalAccessPolicy" = "AuthenticationContext"
            "AuthenticationContextName" = $ContextName
        } | ConvertTo-Json -Depth 3

        $restUrl = "$Url/_api/site"
        
        $response = Invoke-RestMethod -Uri $restUrl -Method PATCH -Headers $headers -Body $body
        Write-Host "✅ Conditional access policy applied via REST API" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "REST API method failed: $($_.Exception.Message)"
        return $false
    }
}

# Method 3: Power Platform / Power Automate approach
function New-PowerAutomateRequest {
    param([string]$Url, [string]$ContextName)
    
    Write-Host "`n=== Method 3: Power Platform Integration ===" -ForegroundColor Cyan
    
    $flowDetails = @{
        FlowName = "SharePoint-ConditionalAccess-Request"
        SiteUrl = $Url
        AuthContextName = $ContextName
        RequestedBy = $env:USERNAME
        RequestDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Status = "Pending Admin Approval"
    }
    
    Write-Host "Power Automate Flow Details:" -ForegroundColor Yellow
    $flowDetails | Format-Table -AutoSize
    
    Write-Host @"
To implement this method:
1. Create a Power Automate flow triggered by SharePoint list item creation
2. The flow should call the SharePoint Admin API with service account credentials
3. Include approval workflow before executing the change
4. Send notification to site owner upon completion

Sample Power Automate HTTP action:
- Method: POST
- URI: https://graph.microsoft.com/v1.0/sites/{site-id}
- Headers: Authorization: Bearer {admin-token}
- Body: {"conditionalAccessPolicy": "AuthenticationContext", "authenticationContextName": "$ContextName"}
"@ -ForegroundColor Cyan

    return $false # This method requires setup
}

# Method 4: Create structured request for admin
function New-AdminRequest {
    param([string]$Url, [string]$ContextName)
    
    Write-Host "`n=== Method 4: Admin Request Workflow ===" -ForegroundColor Cyan
    
    $requestId = "CAP-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    
    $requestDetails = [PSCustomObject]@{
        RequestId = $requestId
        SiteUrl = $Url
        RequestedAction = "Set Conditional Access Policy"
        PolicyType = "AuthenticationContext"
        PolicyName = $ContextName
        RequestedBy = $env:USERNAME
        UserEmail = (whoami /upn 2>$null) -replace '.*\\', ''
        RequestDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        BusinessJustification = "Required for compliance and security requirements"
        ExpectedCompletionDate = (Get-Date).AddDays(3).ToString("yyyy-MM-dd")
        PowerShellCommand = "Set-SPOSite -Identity '$Url' -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName '$ContextName'"
    }
    
    Write-Host "=== SharePoint Admin Request ===" -ForegroundColor Yellow
    $requestDetails | Format-List
    
    # Export to CSV for easy tracking
    $csvPath = ".\Admin_Requests_$(Get-Date -Format 'yyyyMMdd').csv"
    $requestDetails | Export-Csv -Path $csvPath -Append -NoTypeInformation
    
    Write-Host "✅ Request saved to: $csvPath" -ForegroundColor Green
    Write-Host "Please forward this request to your SharePoint Administrator" -ForegroundColor Yellow
    
    return $true
}

# Method 5: Site Collection Features approach
function Enable-SiteCollectionFeatures {
    param([string]$Url, [string]$ContextName)
    
    Write-Host "`n=== Method 5: Site Collection Features ===" -ForegroundColor Cyan
    
    try {
        Connect-PnPOnline -Url $Url -Interactive
        
        # Check for available site collection features related to conditional access
        $features = Get-PnPFeature -Scope Site
        $caFeatures = $features | Where-Object { $_.DisplayName -like "*Conditional*" -or $_.DisplayName -like "*Access*" }
        
        if ($caFeatures) {
            Write-Host "Available Conditional Access features:" -ForegroundColor Green
            $caFeatures | Format-Table DisplayName, Id, DefinitionId
            
            # This would be feature-specific implementation
            Write-Host "Feature-based configuration may be available for your scenario" -ForegroundColor Yellow
        } else {
            Write-Host "No conditional access features found at site collection level" -ForegroundColor Yellow
        }
        
        return $false
    }
    catch {
        Write-Error "Feature check failed: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Host "SharePoint Conditional Access Policy Setup - Site Owner Edition" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green

# Validate parameters
if (-not $SiteUrl.StartsWith("https://")) {
    Write-Error "SiteUrl must be a valid SharePoint site URL starting with https://"
    exit 1
}

# Check current permissions
Write-Host "Checking current user permissions..." -ForegroundColor Yellow
$hasPermissions = Test-SitePermissions -Url $SiteUrl

if (-not $hasPermissions) {
    Write-Warning "You may not have sufficient permissions. Proceeding with available methods..."
}

$success = $false

# Execute based on selected method
switch ($Method) {
    "PnP" { $success = Set-ConditionalAccessPnP -Url $SiteUrl -ContextName $AuthenticationContextName }
    "REST" { $success = Set-ConditionalAccessREST -Url $SiteUrl -ContextName $AuthenticationContextName }
    "PowerPlatform" { $success = New-PowerAutomateRequest -Url $SiteUrl -ContextName $AuthenticationContextName }
    "RequestWorkflow" { $success = New-AdminRequest -Url $SiteUrl -ContextName $AuthenticationContextName }
    "All" {
        $success = Set-ConditionalAccessPnP -Url $SiteUrl -ContextName $AuthenticationContextName
        if (-not $success) {
            $success = Set-ConditionalAccessREST -Url $SiteUrl -ContextName $AuthenticationContextName
        }
        if (-not $success) {
            Enable-SiteCollectionFeatures -Url $SiteUrl -ContextName $AuthenticationContextName
        }
        if (-not $success) {
            New-PowerAutomateRequest -Url $SiteUrl -ContextName $AuthenticationContextName
            New-AdminRequest -Url $SiteUrl -ContextName $AuthenticationContextName
        }
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Green
if ($success) {
    Write-Host "✅ Conditional access policy configuration completed successfully" -ForegroundColor Green
} else {
    Write-Host "⚠️  Direct configuration not possible with current permissions" -ForegroundColor Yellow
    Write-Host "📋 Admin request has been generated for approval workflow" -ForegroundColor Cyan
}

Write-Host "`nFor ongoing management, consider:" -ForegroundColor Cyan
Write-Host "1. Requesting Site Collection Administrator role" -ForegroundColor White
Write-Host "2. Setting up automated Power Platform workflows" -ForegroundColor White
Write-Host "3. Coordinating with SharePoint Administrator for bulk changes" -ForegroundColor White