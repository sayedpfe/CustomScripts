# SharePoint Conditional Access Request Submitter
# This script helps Site Owners submit conditional access policy requests via Power Platform workflow

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AuthenticationContextName,
    
    [Parameter(Mandatory=$true)]
    [string]$BusinessJustification,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Low", "Medium", "High", "Critical")]
    [string]$Priority = "Medium",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("AuthenticationContext", "BlockAccess", "AllowLimitedAccess")]
    [string]$PolicyType = "AuthenticationContext",
    
    [Parameter(Mandatory=$false)]
    [string]$RequestPortalUrl = "https://yourtenant.sharepoint.com/sites/ITRequests"
)

# Function to validate SharePoint URL
function Test-SharePointUrl {
    param([string]$Url)
    
    if (-not $Url.StartsWith("https://")) {
        return $false
    }
    
    if (-not $Url.Contains(".sharepoint.com")) {
        return $false
    }
    
    return $true
}

# Function to connect to the request portal
function Connect-RequestPortal {
    param([string]$PortalUrl)
    
    try {
        Write-Host "Connecting to SharePoint request portal..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $PortalUrl -Interactive -WarningAction SilentlyContinue
        
        # Test connection
        $web = Get-PnPWeb
        Write-Host "✅ Connected to: $($web.Title)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to connect to request portal: $($_.Exception.Message)"
        return $false
    }
}

# Function to check if request already exists
function Test-ExistingRequest {
    param([string]$SiteUrl, [string]$PolicyName)
    
    try {
        $existingItems = Get-PnPListItem -List "ConditionalAccessRequests" -Query "<View><Query><Where><And><Eq><FieldRef Name='SiteURL'/><Value Type='URL'>$SiteUrl</Value></Eq><Eq><FieldRef Name='AuthenticationContextName'/><Value Type='Text'>$PolicyName</Value></Eq></And></Where></Query></View>"
        
        if ($existingItems.Count -gt 0) {
            $activeItems = $existingItems | Where-Object { $_.FieldValues.Status -in @("Pending", "In Review", "Approved") }
            
            if ($activeItems.Count -gt 0) {
                Write-Warning "Active request already exists for this site and policy combination:"
                $activeItems | ForEach-Object {
                    Write-Host "- Request ID: $($_.FieldValues.Title), Status: $($_.FieldValues.Status), Created: $($_.FieldValues.Created)" -ForegroundColor Yellow
                }
                return $true
            }
        }
        
        return $false
    }
    catch {
        Write-Warning "Could not check for existing requests: $($_.Exception.Message)"
        return $false
    }
}

# Function to generate unique request ID
function New-RequestId {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $random = Get-Random -Minimum 1000 -Maximum 9999
    return "CAP-$timestamp-$random"
}

# Function to submit the request
function Submit-ConditionalAccessRequest {
    param(
        [string]$SiteUrl,
        [string]$AuthContextName,
        [string]$Justification,
        [string]$Priority,
        [string]$PolicyType
    )
    
    try {
        # Generate request ID
        $requestId = New-RequestId
        
        # Get current user
        $currentUser = Get-PnPCurrentUser
        
        # Create list item
        $itemValues = @{
            "Title" = $requestId
            "SiteURL" = @{
                "Url" = $SiteUrl
                "Description" = "Target SharePoint Site"
            }
            "AuthenticationContextName" = $AuthContextName
            "BusinessJustification" = $Justification
            "Priority" = $Priority
            "PolicyType" = $PolicyType
            "RequestedById" = $currentUser.Id
            "Status" = "Pending"
        }
        
        Write-Host "Submitting conditional access request..." -ForegroundColor Yellow
        
        $newItem = Add-PnPListItem -List "ConditionalAccessRequests" -Values $itemValues
        
        Write-Host "✅ Request submitted successfully!" -ForegroundColor Green
        Write-Host "Request ID: $requestId" -ForegroundColor Cyan
        Write-Host "Request URL: $($RequestPortalUrl)/Lists/ConditionalAccessRequests/DispForm.aspx?ID=$($newItem.Id)" -ForegroundColor Cyan
        
        return @{
            Success = $true
            RequestId = $requestId
            ItemId = $newItem.Id
            RequestUrl = "$RequestPortalUrl/Lists/ConditionalAccessRequests/DispForm.aspx?ID=$($newItem.Id)"
        }
    }
    catch {
        Write-Error "Failed to submit request: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

# Function to display request summary
function Show-RequestSummary {
    param($RequestDetails)
    
    Write-Host "`n=== Request Summary ===" -ForegroundColor Green
    Write-Host "Target Site: $SiteUrl" -ForegroundColor White
    Write-Host "Policy Name: $AuthenticationContextName" -ForegroundColor White
    Write-Host "Policy Type: $PolicyType" -ForegroundColor White
    Write-Host "Priority: $Priority" -ForegroundColor White
    Write-Host "Business Justification: $BusinessJustification" -ForegroundColor White
    Write-Host "Requested By: $($env:USERNAME)" -ForegroundColor White
    Write-Host "Request Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
    
    if ($RequestDetails.Success) {
        Write-Host "`n=== Submission Details ===" -ForegroundColor Green
        Write-Host "Request ID: $($RequestDetails.RequestId)" -ForegroundColor Cyan
        Write-Host "Portal Item ID: $($RequestDetails.ItemId)" -ForegroundColor Cyan
        Write-Host "Track Request: $($RequestDetails.RequestUrl)" -ForegroundColor Cyan
    }
}

# Function to provide next steps information
function Show-NextSteps {
    Write-Host "`n=== Next Steps ===" -ForegroundColor Green
    Write-Host "1. 📧 SharePoint administrators have been automatically notified" -ForegroundColor White
    Write-Host "2. 📋 Request status has been set to 'In Review'" -ForegroundColor White
    Write-Host "3. 🔍 Admins will review your business justification and priority" -ForegroundColor White
    Write-Host "4. ✅ Once approved, the policy will be automatically applied" -ForegroundColor White
    Write-Host "5. 📩 You'll receive email notifications for status updates" -ForegroundColor White
    
    Write-Host "`n=== Timeline Expectations ===" -ForegroundColor Cyan
    switch ($Priority) {
        "Critical" { Write-Host "Expected approval time: 2-4 hours" -ForegroundColor Red }
        "High" { Write-Host "Expected approval time: 4-8 hours" -ForegroundColor Yellow }
        "Medium" { Write-Host "Expected approval time: 1-2 business days" -ForegroundColor White }
        "Low" { Write-Host "Expected approval time: 2-3 business days" -ForegroundColor Gray }
    }
    
    Write-Host "`nTo check status: Visit the request portal at $RequestPortalUrl" -ForegroundColor Cyan
}

# Function to validate prerequisites
function Test-Prerequisites {
    Write-Host "Validating prerequisites..." -ForegroundColor Yellow
    
    $issues = @()
    
    # Check PnP PowerShell module
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        $issues += "PnP.PowerShell module not installed. Run: Install-Module PnP.PowerShell"
    }
    
    # Validate SharePoint URL
    if (-not (Test-SharePointUrl -Url $SiteUrl)) {
        $issues += "Invalid SharePoint URL format. Must be https://tenant.sharepoint.com/sites/sitename"
    }
    
    # Check authentication context name
    if ($AuthenticationContextName.Length -lt 5) {
        $issues += "Authentication context name too short. Must be descriptive (minimum 5 characters)"
    }
    
    # Check business justification
    if ($BusinessJustification.Length -lt 20) {
        $issues += "Business justification too brief. Please provide detailed explanation (minimum 20 characters)"
    }
    
    if ($issues.Count -gt 0) {
        Write-Host "❌ Prerequisites check failed:" -ForegroundColor Red
        $issues | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }
        return $false
    }
    
    Write-Host "✅ Prerequisites check passed" -ForegroundColor Green
    return $true
}

# Main execution
Write-Host "SharePoint Conditional Access Request Submitter" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green

# Validate prerequisites
if (-not (Test-Prerequisites)) {
    Write-Host "`nPlease resolve the issues above and try again." -ForegroundColor Red
    exit 1
}

# Show request summary
Show-RequestSummary

# Confirm submission
Write-Host "`nDo you want to submit this request? (Y/N): " -ForegroundColor Yellow -NoNewline
$confirmation = Read-Host

if ($confirmation -notlike "Y*") {
    Write-Host "Request submission cancelled." -ForegroundColor Yellow
    exit 0
}

# Connect to SharePoint
if (-not (Connect-RequestPortal -PortalUrl $RequestPortalUrl)) {
    Write-Host "Cannot proceed without connection to request portal." -ForegroundColor Red
    exit 1
}

# Check for existing requests
if (Test-ExistingRequest -SiteUrl $SiteUrl -PolicyName $AuthenticationContextName) {
    Write-Host "`nDo you want to submit a duplicate request anyway? (Y/N): " -ForegroundColor Yellow -NoNewline
    $duplicateConfirmation = Read-Host
    
    if ($duplicateConfirmation -notlike "Y*") {
        Write-Host "Request submission cancelled to avoid duplicates." -ForegroundColor Yellow
        exit 0
    }
}

# Submit the request
$result = Submit-ConditionalAccessRequest -SiteUrl $SiteUrl -AuthContextName $AuthenticationContextName -Justification $BusinessJustification -Priority $Priority -PolicyType $PolicyType

# Show results
Show-RequestSummary -RequestDetails $result

if ($result.Success) {
    Show-NextSteps
    
    # Optional: Open the request in browser
    Write-Host "`nWould you like to open the request in your browser? (Y/N): " -ForegroundColor Yellow -NoNewline
    $browserConfirmation = Read-Host
    
    if ($browserConfirmation -like "Y*") {
        Start-Process $result.RequestUrl
    }
} else {
    Write-Host "`n❌ Request submission failed. Please contact your SharePoint administrator." -ForegroundColor Red
    Write-Host "Error details: $($result.Error)" -ForegroundColor Yellow
}

Write-Host "`nThank you for using the SharePoint Conditional Access Request Portal!" -ForegroundColor Green