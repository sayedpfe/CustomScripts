# SharePoint List Schema for Conditional Access Requests
# Create this list in SharePoint to track requests

@"
List Name: ConditionalAccessRequests
List Type: Custom List

Columns:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Column Name                  Type            Required    Notes
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Title                        Single line     Yes         Request title/description
SiteURL                      Hyperlink       Yes         Full URL to SharePoint site
AuthenticationContextName    Single line     Yes         Name of auth context
RequestorEmail               Single line     Yes         Email of person requesting
BusinessJustification        Multiple lines  Yes         Why policy is needed
Status                       Choice          Yes         Pending/Approved/Rejected/Processing/Completed
RequestDate                  Date/Time       No          Auto-populated with [Today]
ProcessingStarted            Date/Time       No          When automation started
ProcessingCompleted          Date/Time       No          When automation finished
ApprovedBy                   Person/Group    No          Who approved the request
ApprovalDate                 Date/Time       No          When it was approved
Result                       Multiple lines  No          Success/error message
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Status Choice Values:
- Pending (default)
- Approved
- Rejected
- Processing
- Completed
- Failed

"@ | Out-File -FilePath ".\SharePoint-List-Schema.txt" -Encoding UTF8

# PowerShell commands to create the list
$createListScript = @'
# Create SharePoint List for Conditional Access Requests
# Run this after connecting to your SharePoint site

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

# Install PnP PowerShell if needed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

# Connect to site
Connect-PnPOnline -Url $SiteUrl -Interactive

# Create list
$list = New-PnPList -Title "ConditionalAccessRequests" -Template GenericList -OnQuickLaunch

Write-Host "List created: ConditionalAccessRequests" -ForegroundColor Green

# Add custom columns
Write-Host "Adding custom columns..." -ForegroundColor Yellow

# SiteURL (Hyperlink)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "SiteURL" -InternalName "SiteURL" -Type URL -AddToDefaultView -Required

# AuthenticationContextName (Text)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "AuthenticationContextName" -InternalName "AuthenticationContextName" -Type Text -AddToDefaultView -Required

# RequestorEmail (Text)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "RequestorEmail" -InternalName "RequestorEmail" -Type Text -AddToDefaultView -Required

# BusinessJustification (Multi-line)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "BusinessJustification" -InternalName "BusinessJustification" -Type Note -AddToDefaultView -Required

# Status (Choice)
$statusChoices = @("Pending", "Approved", "Rejected", "Processing", "Completed", "Failed")
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices $statusChoices -DefaultValue "Pending" -AddToDefaultView -Required

# RequestDate (DateTime)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "RequestDate" -InternalName "RequestDate" -Type DateTime -AddToDefaultView
Set-PnPField -List "ConditionalAccessRequests" -Identity "RequestDate" -Values @{DefaultValue="[today]"}

# ProcessingStarted (DateTime)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ProcessingStarted" -InternalName "ProcessingStarted" -Type DateTime -AddToDefaultView

# ProcessingCompleted (DateTime)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ProcessingCompleted" -InternalName "ProcessingCompleted" -Type DateTime -AddToDefaultView

# ApprovedBy (Person)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ApprovedBy" -InternalName "ApprovedBy" -Type User -AddToDefaultView

# ApprovalDate (DateTime)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ApprovalDate" -InternalName "ApprovalDate" -Type DateTime -AddToDefaultView

# Result (Multi-line)
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "Result" -InternalName "Result" -Type Note -AddToDefaultView

Write-Host "✅ List created successfully with all columns!" -ForegroundColor Green
Write-Host "`nList URL: $SiteUrl/Lists/ConditionalAccessRequests" -ForegroundColor Cyan

Disconnect-PnPOnline
'@

$createListScript | Out-File -FilePath ".\Create-SharePointList.ps1" -Encoding UTF8

Write-Host "✅ Created SharePoint list schema files:" -ForegroundColor Green
Write-Host "   - SharePoint-List-Schema.txt" -ForegroundColor White
Write-Host "   - Create-SharePointList.ps1" -ForegroundColor White
