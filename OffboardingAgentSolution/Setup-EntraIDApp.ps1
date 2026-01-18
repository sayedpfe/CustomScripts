<#
.SYNOPSIS
    Register an Entra ID App for SharePoint PnP authentication

.DESCRIPTION
    This script creates an Entra ID app registration that can be used with PnP PowerShell
    for SharePoint authentication. Run this once to set up the app.

.EXAMPLE
    .\Setup-EntraIDApp.ps1
#>

[CmdletBinding()]
param()

Write-Host "=== Entra ID App Registration Setup ===" -ForegroundColor Cyan
Write-Host "This script will register an app for SharePoint authentication`n" -ForegroundColor White

# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Write-Host "Installing Microsoft.Graph modules..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

try {
    # Connect to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "Please sign in with admin credentials..." -ForegroundColor Gray
    Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome
    
    Write-Host "Connected successfully!`n" -ForegroundColor Green
    
    # Create the app registration
    $appName = "SharePoint Site Export Tool"
    $redirectUri = "http://localhost"
    
    Write-Host "Creating app registration: $appName" -ForegroundColor Yellow
    
    # Define required resource access for SharePoint
    $requiredResourceAccess = @{
        ResourceAppId = "00000003-0000-0ff1-ce00-000000000000" # SharePoint
        ResourceAccess = @(
            @{
                Id = "678536fe-1083-478a-9c59-b99265e6b0d3" # AllSites.FullControl
                Type = "Scope"
            },
            @{
                Id = "fbcd29d2-fcca-4405-afa2-2d1de87f44e0" # AllSites.Read
                Type = "Scope"
            }
        )
    }
    
    $appParams = @{
        DisplayName = $appName
        SignInAudience = "AzureADMyOrg"
        PublicClient = @{
            RedirectUris = @($redirectUri)
        }
        RequiredResourceAccess = @($requiredResourceAccess)
    }
    
    $app = New-MgApplication @appParams
    
    Write-Host "`nApp registration created successfully!" -ForegroundColor Green
    Write-Host "`n=== IMPORTANT: Save these details ===" -ForegroundColor Cyan
    Write-Host "App Name: $($app.DisplayName)" -ForegroundColor White
    Write-Host "Application (client) ID: $($app.AppId)" -ForegroundColor Yellow
    Write-Host "Redirect URI: $redirectUri" -ForegroundColor White
    
    # Save to config file
    $config = @{
        AppName = $app.DisplayName
        ClientId = $app.AppId
        TenantId = (Get-MgContext).TenantId
        RedirectUri = $redirectUri
        CreatedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    $configPath = Join-Path $PSScriptRoot "AppConfig.json"
    $config | ConvertTo-Json | Out-File $configPath -Encoding UTF8
    
    Write-Host "`nConfiguration saved to: $configPath" -ForegroundColor Green
    
    Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
    Write-Host "1. Admin must grant consent for the app permissions" -ForegroundColor White
    Write-Host "2. Go to: https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps" -ForegroundColor Gray
    Write-Host "3. Find '$appName' and click 'Grant admin consent'" -ForegroundColor Gray
    Write-Host "4. Once consent is granted, you can use the export/deploy scripts" -ForegroundColor White
    
    Write-Host "`nYour Client ID is: $($app.AppId)" -ForegroundColor Yellow
    Write-Host "Use this when running the export/deploy scripts with -ClientId parameter" -ForegroundColor Gray
    
    Disconnect-MgGraph | Out-Null
    
} catch {
    Write-Host "`nError during app registration: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}

Write-Host "`nSetup completed." -ForegroundColor Cyan
