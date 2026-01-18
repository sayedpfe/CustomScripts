# Parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/Bosch",
    
    [Parameter(Mandatory=$false)]
    [string]$AuthContextName = "Sensitive information - guest terms of use"
)

# Get variables from Automation Account
$SPOAdminUrl = Get-AutomationVariable -Name 'SPOAdminUrl'
$AppId = Get-AutomationVariable -Name 'AppId'
$TenantId = Get-AutomationVariable -Name 'TenantId'
$CertThumbprint = Get-AutomationVariable -Name 'CertThumbprint'

Write-Output "Starting SPO Authentication Context Configuration"
Write-Output "================================================"
Write-Output "Target Site: $SiteUrl"
Write-Output "Auth Context: $AuthContextName"
Write-Output ""

try {
    # Connect to SharePoint Online
    Write-Output "Connecting to SharePoint Online Admin..."
    
    Connect-SPOService -Url $SPOAdminUrl `
        -ClientId $AppId `
        -Tenant $TenantId `
        -CertificateThumbprint $CertThumbprint
    
    Write-Output "Successfully connected to SPO Admin"
    
    # Set the Authentication Context
    Write-Output "Setting Authentication Context on site..."
    
    Set-SPOSite -Identity $SiteUrl `
        -ConditionalAccessPolicy AuthenticationContext `
        -AuthenticationContextName $AuthContextName
    
    Write-Output "Successfully configured Authentication Context!"
    
    # Verify the change
    $site = Get-SPOSite -Identity $SiteUrl | Select-Object Url, ConditionalAccessPolicy, AuthenticationContextName
    Write-Output ""
    Write-Output "Verification:"
    Write-Output "  URL: $($site.Url)"
    Write-Output "  Conditional Access Policy: $($site.ConditionalAccessPolicy)"
    Write-Output "  Authentication Context: $($site.AuthenticationContextName)"
    
} catch {
    Write-Error "Error occurred: $_"
    throw $_
} finally {
    # Disconnect
    try {
        Disconnect-SPOService -ErrorAction SilentlyContinue
    } catch {
        # Ignore disconnect errors
    }
}

Write-Output ""
Write-Output "Script completed successfully!"