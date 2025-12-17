# Alternative 1: Using PnP PowerShell with Site Collection Admin rights
# This requires the user to be a Site Collection Administrator (not Global Admin)
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/research" -Interactive

# Set conditional access policy using PnP (if available in your PnP version)
try {
    Set-PnPSite -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName "Sensitive information - guest terms of use"
    Write-Host "Conditional access policy applied successfully using PnP PowerShell" -ForegroundColor Green
} catch {
    Write-Warning "PnP method failed: $($_.Exception.Message)"
    Write-Host "Falling back to alternative methods..." -ForegroundColor Yellow
}

# Alternative 2: Using SharePoint REST API with Site Owner permissions
$siteUrl = "https://contoso.sharepoint.com/sites/research"
$authContextName = "Sensitive information - guest terms of use"

# REST API approach (requires Site Owner permissions)
$headers = @{
    "Accept" = "application/json;odata=verbose"
    "Content-Type" = "application/json;odata=verbose"
}

$body = @{
    ConditionalAccessPolicy = "AuthenticationContext"
    AuthenticationContextName = $authContextName
} | ConvertTo-Json

try {
    $response = Invoke-RestMethod -Uri "$siteUrl/_api/site" -Method POST -Headers $headers -Body $body
    Write-Host "Conditional access policy applied via REST API" -ForegroundColor Green
} catch {
    Write-Warning "REST API method failed: $($_.Exception.Message)"
}

# Alternative 3: Create a request workflow for admin approval
$requestDetails = @{
    SiteUrl = $siteUrl
    RequestedAction = "Set Conditional Access Policy"
    PolicyType = "AuthenticationContext"
    PolicyName = $authContextName
    RequestedBy = $env:USERNAME
    RequestDate = Get-Date
}

Write-Host "=== Admin Request Details ===" -ForegroundColor Cyan
$requestDetails | Format-Table -AutoSize
Write-Host "Please forward this request to your SharePoint Administrator" -ForegroundColor Yellow