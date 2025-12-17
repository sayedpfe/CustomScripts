# Test if SharePoint REST API supports Conditional Access using PnP
# This will give us the definitive answer

param(
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs",
    [string]$AuthContextName = "Sensitive Information - Guest terms of Use"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Definitive SharePoint REST API Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check if PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Write-Host "`n1. Connecting to SharePoint site..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "`n2. Testing REST API - GET site properties..." -ForegroundColor Yellow
try {
    # Use PnP's REST API access
    $response = Invoke-PnPSPRestMethod -Url "/_api/site" -Method Get
    
    Write-Host "   ✅ Can read site via REST API" -ForegroundColor Green
    Write-Host "   Site ID: $($response.Id)" -ForegroundColor White
    
    # Check for conditional access properties
    $hasConditionalAccessProp = $response.PSObject.Properties.Name -contains "ConditionalAccessPolicy"
    $hasAuthContextProp = $response.PSObject.Properties.Name -contains "AuthenticationContextName"
    
    Write-Host "`n   📋 Checking for Conditional Access properties:" -ForegroundColor Cyan
    Write-Host "      - ConditionalAccessPolicy: $(if($hasConditionalAccessProp){'✅ EXISTS'}else{'❌ NOT FOUND'})" -ForegroundColor $(if($hasConditionalAccessProp){'Green'}else{'Red'})
    Write-Host "      - AuthenticationContextName: $(if($hasAuthContextProp){'✅ EXISTS'}else{'❌ NOT FOUND'})" -ForegroundColor $(if($hasAuthContextProp){'Green'}else{'Red'})
    
    if ($hasConditionalAccessProp) {
        Write-Host "      Current value: $($response.ConditionalAccessPolicy)" -ForegroundColor White
    }
    if ($hasAuthContextProp) {
        Write-Host "      Current value: $($response.AuthenticationContextName)" -ForegroundColor White
    }
    
    if (-not $hasConditionalAccessProp -and -not $hasAuthContextProp) {
        Write-Host "`n   ❌ CRITICAL FINDING: Conditional Access properties NOT exposed in REST API" -ForegroundColor Red
        Write-Host "   Available properties:" -ForegroundColor Yellow
        $response.PSObject.Properties.Name | Select-Object -First 20 | ForEach-Object { 
            Write-Host "      - $_" -ForegroundColor Gray 
        }
    }
    
} catch {
    Write-Host "   ❌ FAILED to read site" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test if we can update using REST API
if ($hasConditionalAccessProp) {
    Write-Host "`n3. Testing REST API - UPDATE site properties..." -ForegroundColor Yellow
    
    $updateBody = @{
        "__metadata" = @{ "type" = "SP.Site" }
        "ConditionalAccessPolicy" = "AuthenticationContext"
        "AuthenticationContextName" = $AuthContextName
    }
    
    try {
        $updateResponse = Invoke-PnPSPRestMethod -Url "/_api/site" -Method Post -Content $updateBody
        Write-Host "   ✅ UPDATE request accepted" -ForegroundColor Green
        
        # Verify
        Start-Sleep -Seconds 2
        $verifyResponse = Invoke-PnPSPRestMethod -Url "/_api/site" -Method Get
        
        if ($verifyResponse.ConditionalAccessPolicy -eq "AuthenticationContext") {
            Write-Host "   ✅ SUCCESS! REST API CAN set conditional access!" -ForegroundColor Green
            Write-Host "   Policy: $($verifyResponse.ConditionalAccessPolicy)" -ForegroundColor White
            Write-Host "   Context: $($verifyResponse.AuthenticationContextName)" -ForegroundColor White
        } else {
            Write-Host "   ❌ UPDATE accepted but NOT applied" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "   ❌ FAILED to update site" -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "`n3. Skipping UPDATE test (properties not available)" -ForegroundColor Gray
}

# Test using PowerShell cmdlet
Write-Host "`n4. Testing PnP PowerShell cmdlet Set-PnPSite..." -ForegroundColor Yellow
try {
    Set-PnPSite -Identity $SiteUrl -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName $AuthContextName
    Write-Host "   ✅ PowerShell cmdlet executed successfully" -ForegroundColor Green
    
    # Verify
    Start-Sleep -Seconds 2
    $site = Get-PnPSite -Includes ConditionalAccessPolicy
    Write-Host "   Verification: Policy = $($site.ConditionalAccessPolicy)" -ForegroundColor White
    
} catch {
    Write-Host "   ❌ PowerShell cmdlet failed" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "🎯 FINAL ANSWER" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

if ($hasConditionalAccessProp) {
    Write-Host @"
✅ SharePoint REST API DOES expose Conditional Access properties
   → Power Automate CAN use "Send an HTTP request to SharePoint" action
   → No need for Azure Automation or Azure Functions
   
"@ -ForegroundColor Green
} else {
    Write-Host @"
❌ SharePoint REST API does NOT expose Conditional Access properties
   → These properties are only available through SharePoint Admin API
   → PowerShell cmdlets (Set-SPOSite, Set-PnPSite) work because they call Admin API
   → Power Automate CANNOT do this with REST API alone
   
   ✅ SOLUTION: Use Azure Automation Runbook with PowerShell
      - Power Automate triggers webhook
      - Azure Automation runs Set-PnPSite cmdlet
      - Managed Identity provides admin permissions
   
"@ -ForegroundColor Yellow
}

Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
