# Launcher script - Runs the SPO script in a fresh PowerShell session to avoid module conflicts
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AuthenticationContextName
)

Write-Host "`nLaunching conditional access script in a fresh PowerShell session..." -ForegroundColor Yellow
Write-Host "This avoids conflicts between PnP.PowerShell and Microsoft.Online.SharePoint.PowerShell modules.`n" -ForegroundColor Gray

$scriptPath = Join-Path $PSScriptRoot "Apply-ConditionalAccess-Local-SPO.ps1"

# Build the command to run in new session
$command = "& '$scriptPath' -SiteUrl '$SiteUrl' -AuthenticationContextName '$AuthenticationContextName'; Read-Host 'Press Enter to close this window'"

# Start new PowerShell process
Start-Process pwsh -ArgumentList "-NoProfile", "-Command", $command -Wait

Write-Host "`nScript execution completed. Check the output in the window that opened.`n" -ForegroundColor Green
