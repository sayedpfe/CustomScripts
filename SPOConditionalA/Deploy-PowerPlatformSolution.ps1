# Power Platform Conditional Access Solution Deployment Script
# This script sets up the complete Power Platform workflow for SharePoint conditional access management

param(
    [Parameter(Mandatory=$true)]
    [string]$TenantUrl,  # e.g., "https://contoso.sharepoint.com"
    
    [Parameter(Mandatory=$true)]
    [string]$RequestPortalSiteUrl,  # e.g., "https://contoso.sharepoint.com/sites/ITRequests"
    
    [Parameter(Mandatory=$false)]
    [string]$AdminEmail = "sharepoint-admins@yourdomain.com",
    
    [Parameter(Mandatory=$false)]
    [string]$TeamsChannelId,  # Optional: Teams channel for notifications
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSiteIfNotExists
)

# Import required modules
function Import-RequiredModules {
    Write-Host "Checking required PowerShell modules..." -ForegroundColor Yellow
    
    $requiredModules = @(
        @{Name="PnP.PowerShell"; MinVersion="1.12.0"},
        @{Name="Microsoft.PowerApps.Administration.PowerShell"; MinVersion="2.0.0"},
        @{Name="AzureAD"; MinVersion="2.0.0"}
    )
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -ListAvailable -Name $module.Name | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $installedModule -or $installedModule.Version -lt [version]$module.MinVersion) {
            Write-Host "Installing/updating $($module.Name)..." -ForegroundColor Yellow
            Install-Module -Name $module.Name -Force -AllowClobber -Scope CurrentUser
            Write-Host "✅ $($module.Name) installed successfully" -ForegroundColor Green
        } else {
            Write-Host "✅ $($module.Name) already installed (v$($installedModule.Version))" -ForegroundColor Green
        }
    }
}

# Function to create SharePoint site for requests (if needed)
function New-RequestPortalSite {
    param([string]$SiteUrl, [string]$TenantUrl)
    
    if ($CreateSiteIfNotExists) {
        try {
            $siteName = $SiteUrl.Split('/')[-1]
            Write-Host "Creating SharePoint site for request portal..." -ForegroundColor Yellow
            
            Connect-PnPOnline -Url $TenantUrl -Interactive
            
            $newSite = New-PnPSite -Type TeamSite -Title "IT Service Requests" -Alias $siteName -Description "Portal for IT service requests including SharePoint conditional access policies"
            
            Write-Host "✅ Request portal site created: $($newSite.Url)" -ForegroundColor Green
            return $newSite.Url
        }
        catch {
            Write-Warning "Could not create site automatically: $($_.Exception.Message)"
            Write-Host "Please create the site manually: $SiteUrl" -ForegroundColor Yellow
            return $null
        }
    }
    
    return $SiteUrl
}

# Function to create the SharePoint list
function New-ConditionalAccessRequestsList {
    param([string]$SiteUrl)
    
    Write-Host "Creating ConditionalAccessRequests list..." -ForegroundColor Yellow
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        # Create the main list
        $list = New-PnPList -Title "ConditionalAccessRequests" -Template GenericList -OnQuickLaunch
        
        Write-Host "✅ Base list created successfully" -ForegroundColor Green
        
        # Add custom columns
        $columns = @(
            @{DisplayName="SiteURL"; InternalName="SiteURL"; Type="URL"; Required=$true},
            @{DisplayName="AuthenticationContextName"; InternalName="AuthenticationContextName"; Type="Text"; Required=$true},
            @{DisplayName="BusinessJustification"; InternalName="BusinessJustification"; Type="Note"; Required=$true},
            @{DisplayName="Priority"; InternalName="Priority"; Type="Choice"; Choices=@("Low","Medium","High","Critical"); Default="Medium"},
            @{DisplayName="Status"; InternalName="Status"; Type="Choice"; Choices=@("Pending","In Review","Approved","Rejected","Completed"); Default="Pending"},
            @{DisplayName="PolicyType"; InternalName="PolicyType"; Type="Choice"; Choices=@("AuthenticationContext","BlockAccess","AllowLimitedAccess"); Default="AuthenticationContext"},
            @{DisplayName="ApprovedBy"; InternalName="ApprovedBy"; Type="User"},
            @{DisplayName="ApprovalDate"; InternalName="ApprovalDate"; Type="DateTime"},
            @{DisplayName="CompletedDate"; InternalName="CompletedDate"; Type="DateTime"},
            @{DisplayName="Comments"; InternalName="Comments"; Type="Note"}
        )
        
        foreach ($column in $columns) {
            try {
                switch ($column.Type) {
                    "URL" {
                        Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type -Required:$column.Required
                    }
                    "Text" {
                        Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type -Required:$column.Required
                    }
                    "Note" {
                        Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type -Required:$column.Required
                    }
                    "Choice" {
                        $choiceField = Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type -Choices $column.Choices
                        if ($column.Default) {
                            Set-PnPField -List $list -Identity $column.InternalName -Values @{DefaultValue=$column.Default}
                        }
                    }
                    "User" {
                        Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type
                    }
                    "DateTime" {
                        Add-PnPField -List $list -DisplayName $column.DisplayName -InternalName $column.InternalName -Type $column.Type
                    }
                }
                Write-Host "  ✅ Added column: $($column.DisplayName)" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not add column $($column.DisplayName): $($_.Exception.Message)"
            }
        }
        
        # Create custom views
        Write-Host "Creating custom views..." -ForegroundColor Yellow
        
        # My Requests view
        $myRequestsFields = @("Title", "SiteURL", "AuthenticationContextName", "Priority", "Status", "Created", "ApprovalDate")
        Add-PnPView -List $list -Title "My Requests" -Fields $myRequestsFields -Query "<Where><Eq><FieldRef Name='Author'/><Value Type='Integer'><UserID/></Value></Eq></Where>"
        
        # Pending Approvals view (for admins)
        $pendingFields = @("Title", "SiteURL", "AuthenticationContextName", "Priority", "BusinessJustification", "Author", "Created")
        Add-PnPView -List $list -Title "Pending Approvals" -Fields $pendingFields -Query "<Where><Eq><FieldRef Name='Status'/><Value Type='Choice'>Pending</Value></Eq></Where>"
        
        # All Active Requests view
        $activeFields = @("Title", "SiteURL", "AuthenticationContextName", "Priority", "Status", "Author", "Created", "ApprovalDate")
        Add-PnPView -List $list -Title "All Active Requests" -Fields $activeFields -Query "<Where><Neq><FieldRef Name='Status'/><Value Type='Choice'>Completed</Value></Neq></Where>"
        
        Write-Host "✅ Custom views created successfully" -ForegroundColor Green
        
        # Set list permissions
        Write-Host "Setting list permissions..." -ForegroundColor Yellow
        
        # Allow all authenticated users to contribute (submit requests)
        Set-PnPListPermission -Identity $list -User "Everyone except external users" -AddRole "Contribute"
        
        Write-Host "✅ SharePoint list setup completed successfully" -ForegroundColor Green
        
        return $list
    }
    catch {
        Write-Error "Failed to create SharePoint list: $($_.Exception.Message)"
        return $null
    }
}

# Function to export Power Automate flow templates
function Export-FlowTemplates {
    param([string]$SiteUrl, [string]$AdminEmail, [string]$TeamsChannelId)
    
    Write-Host "Exporting Power Automate flow templates..." -ForegroundColor Yellow
    
    $flowsFolder = ".\PowerAutomate_Templates"
    if (-not (Test-Path $flowsFolder)) {
        New-Item -ItemType Directory -Path $flowsFolder -Force | Out-Null
    }
    
    # Flow 1: Request Notification Handler
    $notificationFlow = @{
        "displayName" = "SharePoint Conditional Access - Request Notification"
        "description" = "Notifies admins when new conditional access requests are submitted"
        "trigger" = @{
            "type" = "SharePoint"
            "operation" = "OnNewItem"
            "siteUrl" = $SiteUrl
            "listName" = "ConditionalAccessRequests"
        }
        "actions" = @(
            @{
                "name" = "Send notification email"
                "type" = "Office365Outlook"
                "operation" = "SendEmailV2"
                "parameters" = @{
                    "To" = $AdminEmail
                    "Subject" = "New SharePoint Conditional Access Request"
                    "Body" = "A new conditional access policy request has been submitted. Please review in the admin portal."
                }
            },
            @{
                "name" = "Update status to In Review"
                "type" = "SharePoint"
                "operation" = "UpdateItem"
                "parameters" = @{
                    "Status" = "In Review"
                }
            }
        )
    }
    
    # Flow 2: Request Executor
    $executorFlow = @{
        "displayName" = "SharePoint Conditional Access - Request Executor"
        "description" = "Executes approved conditional access policy requests using service account"
        "trigger" = @{
            "type" = "SharePoint"
            "operation" = "OnUpdatedItem"
            "siteUrl" = $SiteUrl
            "listName" = "ConditionalAccessRequests"
            "condition" = "Status equals 'Approved'"
        }
        "actions" = @(
            @{
                "name" = "Execute SharePoint command via Graph API"
                "type" = "HTTP"
                "method" = "PATCH"
                "uri" = "https://graph.microsoft.com/v1.0/sites/{site-id}"
                "authentication" = @{
                    "type" = "ServicePrincipal"
                    "clientId" = "{service-account-client-id}"
                    "clientSecret" = "{service-account-secret}"
                    "tenant" = "{tenant-id}"
                }
                "body" = @{
                    "conditionalAccessPolicy" = "AuthenticationContext"
                    "authenticationContextName" = "@{triggerBody()?['AuthenticationContextName']}"
                }
            },
            @{
                "name" = "Mark as completed"
                "type" = "SharePoint"
                "operation" = "UpdateItem"
                "parameters" = @{
                    "Status" = "Completed"
                    "CompletedDate" = "@{utcnow()}"
                }
            },
            @{
                "name" = "Send completion notification"
                "type" = "Office365Outlook"
                "operation" = "SendEmailV2"
                "parameters" = @{
                    "To" = "@{triggerBody()?['Author']?['Email']}"
                    "Subject" = "SharePoint Conditional Access Request Completed"
                    "Body" = "Your conditional access policy request has been implemented successfully."
                }
            }
        )
    }
    
    # Save flow templates
    $notificationFlow | ConvertTo-Json -Depth 10 | Out-File "$flowsFolder\NotificationFlow.json" -Encoding UTF8
    $executorFlow | ConvertTo-Json -Depth 10 | Out-File "$flowsFolder\ExecutorFlow.json" -Encoding UTF8
    
    Write-Host "✅ Flow templates exported to: $flowsFolder" -ForegroundColor Green
    
    # Create deployment instructions
    $instructions = @"
# Power Automate Flow Deployment Instructions

## Prerequisites
1. Power Automate Premium license (for HTTP connector with authentication)
2. Azure AD App Registration with SharePoint admin permissions
3. SharePoint Administrator role for service account

## Deployment Steps

### Step 1: Import Notification Flow
1. Go to Power Automate portal (https://flow.microsoft.com)
2. Click "Create" > "Import package (legacy)"
3. Upload NotificationFlow.json
4. Configure connections:
   - SharePoint: Connect to your tenant
   - Office 365 Outlook: Connect with admin account
5. Update trigger settings:
   - Site URL: $SiteUrl
   - List: ConditionalAccessRequests
6. Update email settings:
   - To: $AdminEmail

### Step 2: Import Executor Flow
1. Import ExecutorFlow.json following same process
2. Configure HTTP action authentication:
   - Client ID: {Your Azure AD App Client ID}
   - Client Secret: {Your Azure AD App Secret}
   - Tenant ID: {Your Azure AD Tenant ID}
3. Test with a sample approved request

### Step 3: Enable Flows
1. Turn on both flows
2. Test end-to-end workflow
3. Monitor flow run history for errors

## Service Account Setup Required
- Create Azure AD App Registration
- Assign SharePoint Administrator role
- Generate client secret
- Configure API permissions: Sites.FullControl.All

## Support
Contact your Power Platform administrator for deployment assistance.
"@
    
    $instructions | Out-File "$flowsFolder\DEPLOYMENT_INSTRUCTIONS.md" -Encoding UTF8
    
    Write-Host "✅ Deployment instructions created" -ForegroundColor Green
}

# Function to create sample data for testing
function New-SampleRequests {
    param([string]$SiteUrl)
    
    Write-Host "Creating sample requests for testing..." -ForegroundColor Yellow
    
    $sampleRequests = @(
        @{
            Title = "CAP-TEST-001"
            SiteURL = @{Url = "https://contoso.sharepoint.com/sites/finance"; Description = "Finance Team Site"}
            AuthenticationContextName = "Sensitive financial data - MFA required"
            BusinessJustification = "This site contains sensitive financial information that requires additional security controls to comply with SOX requirements."
            Priority = "High"
            PolicyType = "AuthenticationContext"
            Status = "Pending"
        },
        @{
            Title = "CAP-TEST-002"
            SiteURL = @{Url = "https://contoso.sharepoint.com/sites/hr"; Description = "HR Team Site"}
            AuthenticationContextName = "PII protection - guest access restricted"
            BusinessJustification = "HR site contains personally identifiable information that should not be accessible to external users."
            Priority = "Critical"
            PolicyType = "AuthenticationContext"
            Status = "In Review"
        }
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        
        foreach ($request in $sampleRequests) {
            $newItem = Add-PnPListItem -List "ConditionalAccessRequests" -Values $request
            Write-Host "  ✅ Created sample request: $($request.Title)" -ForegroundColor Green
        }
        
        Write-Host "✅ Sample requests created for testing" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not create sample requests: $($_.Exception.Message)"
    }
}

# Function to display deployment summary
function Show-DeploymentSummary {
    param([string]$SiteUrl, [string]$AdminEmail)
    
    Write-Host "`n" + "="*60 -ForegroundColor Green
    Write-Host "POWER PLATFORM DEPLOYMENT COMPLETED" -ForegroundColor Green
    Write-Host "="*60 -ForegroundColor Green
    
    Write-Host "`n📋 SharePoint List:" -ForegroundColor Cyan
    Write-Host "   URL: $SiteUrl/Lists/ConditionalAccessRequests" -ForegroundColor White
    Write-Host "   Status: Created with custom views and permissions" -ForegroundColor White
    
    Write-Host "`n🔄 Power Automate Flows:" -ForegroundColor Cyan
    Write-Host "   Templates: Exported to .\PowerAutomate_Templates\" -ForegroundColor White
    Write-Host "   Status: Ready for manual import and configuration" -ForegroundColor White
    
    Write-Host "`n👥 User Experience:" -ForegroundColor Cyan
    Write-Host "   Site Owners: Can submit requests via SharePoint list or PowerShell script" -ForegroundColor White
    Write-Host "   Admins: Receive notifications at $AdminEmail" -ForegroundColor White
    
    Write-Host "`n📝 Next Steps:" -ForegroundColor Yellow
    Write-Host "1. Import Power Automate flows using provided templates" -ForegroundColor White
    Write-Host "2. Create Azure AD App Registration for service account" -ForegroundColor White
    Write-Host "3. Configure flow authentication and permissions" -ForegroundColor White
    Write-Host "4. Test end-to-end workflow with sample requests" -ForegroundColor White
    Write-Host "5. Train users on submission process" -ForegroundColor White
    
    Write-Host "`n📚 Documentation:" -ForegroundColor Cyan
    Write-Host "   Detailed guide: PowerPlatform_Detailed_Guide.md" -ForegroundColor White
    Write-Host "   User script: Submit-ConditionalAccessRequest.ps1" -ForegroundColor White
    Write-Host "   Flow templates: .\PowerAutomate_Templates\" -ForegroundColor White
}

# Main deployment process
Write-Host "Power Platform Conditional Access Solution Deployment" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green

try {
    # Step 1: Import required modules
    Import-RequiredModules
    
    # Step 2: Create/verify request portal site
    $portalSite = New-RequestPortalSite -SiteUrl $RequestPortalSiteUrl -TenantUrl $TenantUrl
    if (-not $portalSite) {
        Write-Error "Cannot proceed without request portal site"
        exit 1
    }
    
    # Step 3: Create SharePoint list
    $list = New-ConditionalAccessRequestsList -SiteUrl $RequestPortalSiteUrl
    if (-not $list) {
        Write-Error "Failed to create SharePoint list"
        exit 1
    }
    
    # Step 4: Export Power Automate flow templates
    Export-FlowTemplates -SiteUrl $RequestPortalSiteUrl -AdminEmail $AdminEmail -TeamsChannelId $TeamsChannelId
    
    # Step 5: Create sample data (optional)
    Write-Host "`nWould you like to create sample requests for testing? (Y/N): " -ForegroundColor Yellow -NoNewline
    $createSamples = Read-Host
    
    if ($createSamples -like "Y*") {
        New-SampleRequests -SiteUrl $RequestPortalSiteUrl
    }
    
    # Step 6: Show deployment summary
    Show-DeploymentSummary -SiteUrl $RequestPortalSiteUrl -AdminEmail $AdminEmail
    
    Write-Host "`n✅ Deployment completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host "Please check the error details and retry the deployment." -ForegroundColor Yellow
    exit 1
}