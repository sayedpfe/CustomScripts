# Test SharePoint REST API for Conditional Access Policy
# This will determine if SharePoint REST API supports this operation

param(
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs",
    [string]$AuthContextName = "Sensitive Information - Guest terms of Use"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Testing SharePoint REST API Capabilities" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Connect to Microsoft Graph to get token
Write-Host "`n1. Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "Sites.FullControl.All" -NoWelcome

# Get SharePoint access token
Write-Host "`n2. Getting SharePoint access token..." -ForegroundColor Yellow
$context = Get-MgContext
$tenantId = $context.TenantId

# Get SharePoint token using delegated permissions
$sharepointToken = (Get-MgContext).Token
if (-not $sharepointToken) {
    Write-Host "   Unable to get SharePoint token from current context" -ForegroundColor Red
    Write-Host "   Trying alternative method..." -ForegroundColor Yellow
}

# Test 1: Get current site properties via REST API
Write-Host "`n3. Testing SharePoint REST API - GET site properties..." -ForegroundColor Yellow
$restUrl = "$SiteUrl/_api/site"
$headers = @{
    "Authorization" = "Bearer $sharepointToken"
    "Accept" = "application/json;odata=verbose"
}

try {
    $response = Invoke-RestMethod -Uri $restUrl -Headers $headers -Method Get
    Write-Host "   ✅ SUCCESS: Can read site via REST API" -ForegroundColor Green
    Write-Host "   Site Title: $($response.d.Title)" -ForegroundColor White
    
    # Check if conditional access properties exist
    if ($response.d.PSObject.Properties.Name -contains "ConditionalAccessPolicy") {
        Write-Host "   ✅ ConditionalAccessPolicy property exists!" -ForegroundColor Green
        Write-Host "      Current Value: $($response.d.ConditionalAccessPolicy)" -ForegroundColor White
    } else {
        Write-Host "   ❌ ConditionalAccessPolicy property NOT found in response" -ForegroundColor Red
    }
    
    if ($response.d.PSObject.Properties.Name -contains "AuthenticationContextName") {
        Write-Host "   ✅ AuthenticationContextName property exists!" -ForegroundColor Green
        Write-Host "      Current Value: $($response.d.AuthenticationContextName)" -ForegroundColor White
    } else {
        Write-Host "   ❌ AuthenticationContextName property NOT found in response" -ForegroundColor Red
    }
    
    Write-Host "`n   Available properties:" -ForegroundColor Cyan
    $response.d.PSObject.Properties.Name | ForEach-Object { Write-Host "      - $_" -ForegroundColor Gray }
    
} catch {
    Write-Host "   ❌ FAILED to read site via REST API" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Try to update site properties via REST API
Write-Host "`n4. Testing SharePoint REST API - UPDATE site properties..." -ForegroundColor Yellow
$updateUrl = "$SiteUrl/_api/site"
$updateHeaders = @{
    "Authorization" = "Bearer $sharepointToken"
    "Accept" = "application/json;odata=verbose"
    "Content-Type" = "application/json;odata=verbose"
    "X-HTTP-Method" = "MERGE"
    "IF-MATCH" = "*"
}

$body = @{
    "__metadata" = @{ "type" = "SP.Site" }
    "ConditionalAccessPolicy" = "AuthenticationContext"
    "AuthenticationContextName" = $AuthContextName
} | ConvertTo-Json

try {
    $updateResponse = Invoke-RestMethod -Uri $updateUrl -Headers $updateHeaders -Method Post -Body $body
    Write-Host "   ✅ UPDATE request completed without error" -ForegroundColor Green
    Write-Host "   Response: $updateResponse" -ForegroundColor White
    
    # Verify the change
    Write-Host "`n5. Verifying the change..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    $verifyResponse = Invoke-RestMethod -Uri $restUrl -Headers $headers -Method Get
    
    if ($verifyResponse.d.ConditionalAccessPolicy -eq "AuthenticationContext") {
        Write-Host "   ✅ SUCCESS! Conditional Access Policy was applied!" -ForegroundColor Green
        Write-Host "   Policy: $($verifyResponse.d.ConditionalAccessPolicy)" -ForegroundColor White
        Write-Host "   Auth Context: $($verifyResponse.d.AuthenticationContextName)" -ForegroundColor White
    } else {
        Write-Host "   ❌ FAILED: Policy was NOT applied" -ForegroundColor Red
        Write-Host "   Policy is still: $($verifyResponse.d.ConditionalAccessPolicy)" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "   ❌ FAILED to update site via REST API" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $responseBody = $reader.ReadToEnd()
        Write-Host "   Response Body: $responseBody" -ForegroundColor Red
    }
}

# Test 3: Check SharePoint Admin API
Write-Host "`n6. Testing SharePoint Admin API..." -ForegroundColor Yellow
$adminUrl = "https://m365cpi90282478-admin.sharepoint.com/_api/SPO.Tenant/sites/getbyurl('$SiteUrl')"

try {
    $adminResponse = Invoke-RestMethod -Uri $adminUrl -Headers $headers -Method Get
    Write-Host "   ✅ Can access SharePoint Admin API" -ForegroundColor Green
    Write-Host "   Site properties available via Admin API" -ForegroundColor White
} catch {
    Write-Host "   ❌ Cannot access SharePoint Admin API" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   This likely requires SharePoint Administrator role" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CONCLUSION:" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host @"

Based on the tests above:
1. If ConditionalAccessPolicy property appears in GET response: SharePoint REST API MAY support it
2. If UPDATE succeeds and verification confirms: SharePoint REST API DOES support it
3. If UPDATE fails or doesn't persist: Only PowerShell cmdlets work (they use SharePoint Admin API)

For Power Automate:
- If REST API works: Use "Send an HTTP request to SharePoint" action
- If REST API doesn't work: Must use Azure Automation Runbook with PowerShell

"@ -ForegroundColor White

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Gray
