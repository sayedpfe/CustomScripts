# Create Authentication Context for SharePoint Conditional Access
# Based on: https://learn.microsoft.com/en-us/sharepoint/authentication-context-example

param(
    [Parameter(Mandatory=$false)]
    [string]$DisplayName = "Test Policy - Graph API Demo",
    
    [Parameter(Mandatory=$false)]
    [string]$Description = "Authentication context for SharePoint conditional access demo",
    
    [Parameter(Mandatory=$false)]
    [string]$Id = "c1"  # Custom ID for the auth context (c1-c25 are allowed)
)

Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     Create Authentication Context for SharePoint Online      ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Function to check if Microsoft Graph module is installed
function Test-GraphModule {
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Identity.SignIns
    if (-not $module) {
        Write-Host "Microsoft.Graph.Identity.SignIns module not found." -ForegroundColor Yellow
        Write-Host "Installing module..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
        Write-Host "✅ Module installed successfully" -ForegroundColor Green
    } else {
        Write-Host "✅ Microsoft.Graph.Identity.SignIns module found (v$($module.Version))" -ForegroundColor Green
    }
}

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    Write-Host "`n🔐 Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    try {
        # Connect with required permissions for authentication context
        Connect-MgGraph -Scopes "Policy.Read.All", "Policy.ReadWrite.ConditionalAccess", "Application.Read.All" -NoWelcome
        
        $context = Get-MgContext
        Write-Host "✅ Connected to Microsoft Graph" -ForegroundColor Green
        Write-Host "   Tenant: $($context.TenantId)" -ForegroundColor Gray
        Write-Host "   Account: $($context.Account)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        return $false
    }
}

# Function to check if authentication context already exists
function Get-ExistingAuthContext {
    param([string]$DisplayName, [string]$Id)
    
    Write-Host "`n🔍 Checking for existing authentication contexts..." -ForegroundColor Yellow
    
    try {
        # Get all authentication contexts
        $authContexts = Get-MgIdentityConditionalAccessAuthenticationContextClassReference
        
        if ($authContexts) {
            Write-Host "Found $($authContexts.Count) existing authentication context(s):" -ForegroundColor Cyan
            $authContexts | ForEach-Object {
                Write-Host "   - $($_.DisplayName) (ID: $($_.Id))" -ForegroundColor Gray
            }
            
            # Check if our context already exists
            $existing = $authContexts | Where-Object { $_.DisplayName -eq $DisplayName -or $_.Id -eq $Id }
            
            if ($existing) {
                Write-Host "`n⚠️  Authentication context already exists!" -ForegroundColor Yellow
                Write-Host "   Display Name: $($existing.DisplayName)" -ForegroundColor Yellow
                Write-Host "   ID: $($existing.Id)" -ForegroundColor Yellow
                Write-Host "   Description: $($existing.Description)" -ForegroundColor Yellow
                return $existing
            }
        } else {
            Write-Host "No existing authentication contexts found." -ForegroundColor Gray
        }
        
        return $null
    }
    catch {
        Write-Warning "Could not retrieve existing authentication contexts: $($_.Exception.Message)"
        return $null
    }
}

# Function to create new authentication context
function New-AuthenticationContext {
    param(
        [string]$DisplayName,
        [string]$Description,
        [string]$Id
    )
    
    Write-Host "`n🚀 Creating new authentication context..." -ForegroundColor Yellow
    Write-Host "════════════════════════════════════════" -ForegroundColor Yellow
    
    try {
        # Create the authentication context class reference
        $params = @{
            Id = $Id
            DisplayName = $DisplayName
            Description = $Description
            IsAvailable = $true
        }
        
        Write-Host "Parameters:" -ForegroundColor Cyan
        Write-Host "   ID: $Id" -ForegroundColor White
        Write-Host "   Display Name: $DisplayName" -ForegroundColor White
        Write-Host "   Description: $Description" -ForegroundColor White
        Write-Host "   Is Available: True" -ForegroundColor White
        
        Write-Host "`n⏳ Creating authentication context in Azure AD..." -ForegroundColor Yellow
        
        $newContext = New-MgIdentityConditionalAccessAuthenticationContextClassReference -BodyParameter $params
        
        Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║         ✅ Authentication Context Created!            ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Green
        
        Write-Host "`nAuthentication Context Details:" -ForegroundColor Green
        Write-Host "   ID: $($newContext.Id)" -ForegroundColor Green
        Write-Host "   Display Name: $($newContext.DisplayName)" -ForegroundColor Green
        Write-Host "   Description: $($newContext.Description)" -ForegroundColor Green
        Write-Host "   Is Available: $($newContext.IsAvailable)" -ForegroundColor Green
        
        return $newContext
    }
    catch {
        Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "║                    ❌ FAILED!                         ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Red
        
        Write-Error "Failed to create authentication context: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "`n💡 Tip: Try using a different ID (c1-c25 are valid)" -ForegroundColor Yellow
        }
        
        return $null
    }
}

# Function to show how to use this in SharePoint
function Show-SharePointUsageInstructions {
    param($AuthContext)
    
    Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║              How to Use This Authentication Context          ║" -ForegroundColor Magenta
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    
    Write-Host "`n📋 Method 1: PowerShell (Admin Required)" -ForegroundColor Cyan
    Write-Host "Set-SPOSite -Identity 'https://tenant.sharepoint.com/sites/sitename' ``" -ForegroundColor White
    Write-Host "             -ConditionalAccessPolicy AuthenticationContext ``" -ForegroundColor White
    Write-Host "             -AuthenticationContextName '$($AuthContext.DisplayName)'" -ForegroundColor White
    
    Write-Host "`n📋 Method 2: Microsoft Graph API (What Power Automate Uses)" -ForegroundColor Cyan
    Write-Host "PATCH https://graph.microsoft.com/v1.0/sites/{site-id}" -ForegroundColor White
    Write-Host "Content-Type: application/json" -ForegroundColor Gray
    Write-Host "{" -ForegroundColor White
    Write-Host '  "conditionalAccessPolicy": "AuthenticationContext",' -ForegroundColor White
    Write-Host "  `"authenticationContextName`": `"$($AuthContext.DisplayName)`"" -ForegroundColor White
    Write-Host "}" -ForegroundColor White
    
    Write-Host "`n📋 Method 3: Run Our Demo Script" -ForegroundColor Cyan
    Write-Host ".\Demo-GraphAPI-ConditionalAccess.ps1 ``" -ForegroundColor White
    Write-Host "    -SiteUrl 'https://m365cpi90282478.sharepoint.com/sites/DeutschKurs' ``" -ForegroundColor White
    Write-Host "    -AuthenticationContextName '$($AuthContext.DisplayName)' ``" -ForegroundColor White
    Write-Host "    -UseInteractiveAuth ``" -ForegroundColor White
    Write-Host "    -TenantId 'M365CPI90282478.onmicrosoft.com'" -ForegroundColor White
}

# Function to create a simple conditional access policy (optional)
function New-ConditionalAccessPolicy {
    param(
        [string]$PolicyName,
        [string]$AuthContextId
    )
    
    Write-Host "`n📋 Would you like to create a conditional access policy using this context? (Y/N): " -ForegroundColor Yellow -NoNewline
    $createPolicy = Read-Host
    
    if ($createPolicy -notlike "Y*") {
        Write-Host "Skipping conditional access policy creation." -ForegroundColor Gray
        return $null
    }
    
    Write-Host "`n🚀 Creating conditional access policy..." -ForegroundColor Yellow
    
    try {
        # This is a basic example - you can customize this
        $params = @{
            DisplayName = "$PolicyName - CA Policy"
            State = "disabled"  # Start disabled for safety
            Conditions = @{
                Applications = @{
                    IncludeApplications = @("All")
                }
                Users = @{
                    IncludeUsers = @("All")
                    ExcludeUsers = @()  # Add break-glass accounts here
                }
                AuthenticationContextClassReferences = @($AuthContextId)
            }
            GrantControls = @{
                Operator = "OR"
                BuiltInControls = @("mfa")  # Require MFA
            }
        }
        
        Write-Host "⚠️  Note: This is a basic example policy that requires MFA" -ForegroundColor Yellow
        Write-Host "   Customize it in Azure Portal before enabling!" -ForegroundColor Yellow
        
        $policy = New-MgIdentityConditionalAccessPolicy -BodyParameter $params
        
        Write-Host "✅ Conditional Access Policy created (disabled)" -ForegroundColor Green
        Write-Host "   Policy Name: $($policy.DisplayName)" -ForegroundColor Green
        Write-Host "   Policy ID: $($policy.Id)" -ForegroundColor Green
        Write-Host "   State: $($policy.State)" -ForegroundColor Yellow
        
        Write-Host "`n⚠️  Remember to enable this policy in Azure Portal after reviewing!" -ForegroundColor Yellow
        
        return $policy
    }
    catch {
        Write-Warning "Could not create conditional access policy: $($_.Exception.Message)"
        return $null
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Host "This script will create an authentication context for SharePoint Online" -ForegroundColor Cyan
Write-Host "based on Microsoft's official documentation." -ForegroundColor Cyan

# Step 1: Check and install required modules
Write-Host "`nSTEP 1: Checking Prerequisites" -ForegroundColor Yellow
Write-Host "═══════════════════════════" -ForegroundColor Yellow
Test-GraphModule

# Step 2: Connect to Microsoft Graph
Write-Host "`nSTEP 2: Authentication" -ForegroundColor Yellow
Write-Host "═══════════════════════" -ForegroundColor Yellow

$connected = Connect-ToGraph
if (-not $connected) {
    Write-Error "Cannot proceed without Graph connection."
    exit 1
}

# Step 3: Check for existing authentication context
Write-Host "`nSTEP 3: Check Existing Contexts" -ForegroundColor Yellow
Write-Host "═══════════════════════════════" -ForegroundColor Yellow

$existing = Get-ExistingAuthContext -DisplayName $DisplayName -Id $Id

if ($existing) {
    Write-Host "`nAuthentication context already exists. Would you like to:" -ForegroundColor Yellow
    Write-Host "1. Use the existing context" -ForegroundColor White
    Write-Host "2. Create a new one with different name/ID" -ForegroundColor White
    Write-Host "3. Exit" -ForegroundColor White
    
    $choice = Read-Host "`nEnter your choice (1-3)"
    
    switch ($choice) {
        "1" {
            Write-Host "✅ Using existing authentication context" -ForegroundColor Green
            $authContext = $existing
        }
        "2" {
            $DisplayName = Read-Host "Enter new display name"
            $Id = Read-Host "Enter new ID (c1-c25)"
            $authContext = New-AuthenticationContext -DisplayName $DisplayName -Description $Description -Id $Id
        }
        "3" {
            Write-Host "Exiting..." -ForegroundColor Gray
            exit 0
        }
        default {
            Write-Host "Invalid choice. Using existing context." -ForegroundColor Yellow
            $authContext = $existing
        }
    }
} else {
    # Step 4: Create new authentication context
    Write-Host "`nSTEP 4: Create Authentication Context" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════" -ForegroundColor Yellow
    
    $authContext = New-AuthenticationContext -DisplayName $DisplayName -Description $Description -Id $Id
}

if (-not $authContext) {
    Write-Error "Failed to create or retrieve authentication context."
    exit 1
}

# Step 5: Show usage instructions
Write-Host "`nSTEP 5: Usage Instructions" -ForegroundColor Yellow
Write-Host "═══════════════════════════" -ForegroundColor Yellow

Show-SharePointUsageInstructions -AuthContext $authContext

# Step 6: Optionally create a conditional access policy
Write-Host "`nSTEP 6: Conditional Access Policy (Optional)" -ForegroundColor Yellow
Write-Host "════════════════════════════════════════════" -ForegroundColor Yellow

New-ConditionalAccessPolicy -PolicyName $DisplayName -AuthContextId $authContext.Id

# Final summary
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                    Setup Complete! ✅                          ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green

Write-Host "`n🎯 NEXT STEPS:" -ForegroundColor Cyan
Write-Host "1. The authentication context is now created in Azure AD" -ForegroundColor White
Write-Host "2. You can now run the Graph API demo script" -ForegroundColor White
Write-Host "3. The demo will apply this context to your SharePoint site" -ForegroundColor White

Write-Host "`n📋 Ready to run the demo? Execute:" -ForegroundColor Yellow
Write-Host ".\Demo-GraphAPI-ConditionalAccess.ps1 ``" -ForegroundColor Cyan
Write-Host "    -SiteUrl 'https://m365cpi90282478.sharepoint.com/sites/DeutschKurs' ``" -ForegroundColor Cyan
Write-Host "    -AuthenticationContextName '$($authContext.DisplayName)' ``" -ForegroundColor Cyan
Write-Host "    -UseInteractiveAuth ``" -ForegroundColor Cyan
Write-Host "    -TenantId 'M365CPI90282478.onmicrosoft.com'" -ForegroundColor Cyan

Write-Host "`n✅ Authentication context is ready to use!" -ForegroundColor Green