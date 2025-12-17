# Power Platform Workflow for SharePoint Conditional Access Management

## Overview
This solution uses Power Platform (Power Automate + SharePoint Lists) to create a self-service portal for Site Owners to request conditional access policy changes without requiring admin privileges.

## Architecture Components

### 1. SharePoint List (Request Portal)
**List Name**: `ConditionalAccessRequests`

**Columns**:
- **Title** (Single line of text) - Request ID
- **SiteURL** (Hyperlink) - Target SharePoint site
- **AuthenticationContextName** (Single line of text) - Policy name
- **RequestedBy** (Person) - Auto-populated
- **BusinessJustification** (Multiple lines of text) - Why this change is needed
- **Priority** (Choice) - Low, Medium, High, Critical
- **RequestDate** (Date) - Auto-populated
- **Status** (Choice) - Pending, In Review, Approved, Rejected, Completed
- **ApprovedBy** (Person) - SharePoint Admin who approved
- **ApprovalDate** (Date) - When approved
- **CompletedDate** (Date) - When actually implemented
- **Comments** (Multiple lines of text) - Admin notes
- **PolicyType** (Choice) - AuthenticationContext, BlockAccess, AllowLimitedAccess

### 2. Power Automate Flows

#### Flow 1: Request Submission Handler
**Trigger**: When an item is created in ConditionalAccessRequests list

```json
{
  "name": "SharePoint-ConditionalAccess-RequestHandler",
  "trigger": "SharePoint - When an item is created",
  "actions": [
    {
      "name": "Send_notification_to_admins",
      "type": "Office365Outlook.SendEmailV2"
    },
    {
      "name": "Update_status_to_InReview",
      "type": "SharePoint.UpdateItem"
    },
    {
      "name": "Create_Teams_notification",
      "type": "MicrosoftTeams.PostMessageV3"
    }
  ]
}
```

#### Flow 2: Approval and Execution Handler
**Trigger**: When an item is modified in ConditionalAccessRequests list (Status = Approved)

```json
{
  "name": "SharePoint-ConditionalAccess-Executor",
  "trigger": "SharePoint - When an item is modified",
  "condition": "Status equals 'Approved'",
  "actions": [
    {
      "name": "Execute_SPO_Command",
      "type": "HTTP",
      "method": "POST",
      "uri": "https://graph.microsoft.com/v1.0/sites/@{variables('siteId')}"
    },
    {
      "name": "Update_completion_status",
      "type": "SharePoint.UpdateItem"
    },
    {
      "name": "Notify_requestor",
      "type": "Office365Outlook.SendEmailV2"
    }
  ]
}
```

## Implementation Guide

### Phase 1: SharePoint List Setup

1. **Create the Request List**
```powershell
# PowerShell script to create the SharePoint list
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/ITRequests" -Interactive

# Create the list
New-PnPList -Title "ConditionalAccessRequests" -Template GenericList

# Add custom columns
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "SiteURL" -InternalName "SiteURL" -Type URL
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "AuthenticationContextName" -InternalName "AuthenticationContextName" -Type Text
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "BusinessJustification" -InternalName "BusinessJustification" -Type Note
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices @("Low","Medium","High","Critical")
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices @("Pending","In Review","Approved","Rejected","Completed")
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ApprovedBy" -InternalName "ApprovedBy" -Type User
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "ApprovalDate" -InternalName "ApprovalDate" -Type DateTime
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "Comments" -InternalName "Comments" -Type Note
Add-PnPField -List "ConditionalAccessRequests" -DisplayName "PolicyType" -InternalName "PolicyType" -Type Choice -Choices @("AuthenticationContext","BlockAccess","AllowLimitedAccess")

Write-Host "SharePoint list created successfully!" -ForegroundColor Green
```

2. **Set List Permissions**
```powershell
# Give Site Owners contribute access to submit requests
Set-PnPListPermission -Identity "ConditionalAccessRequests" -User "Site Owners" -AddRole "Contribute"

# Give SharePoint Admins full control
Set-PnPListPermission -Identity "ConditionalAccessRequests" -User "SharePoint Admins" -AddRole "Full Control"
```

### Phase 2: Power Automate Flow Creation

#### Flow 1: Request Notification Flow

```json
{
  "definition": {
    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {},
    "triggers": {
      "When_an_item_is_created": {
        "recurrence": {
          "frequency": "Second",
          "interval": 10
        },
        "splitOn": "@triggerBody()?['value']",
        "type": "ApiConnection",
        "inputs": {
          "host": {
            "connection": {
              "name": "@parameters('$connections')['sharepointonline']['connectionId']"
            }
          },
          "method": "get",
          "path": "/datasets/@{encodeURIComponent('https://yourtenant.sharepoint.com/sites/ITRequests')}/tables/@{encodeURIComponent('ConditionalAccessRequests')}/onnewitems"
        }
      }
    },
    "actions": {
      "Send_email_to_admins": {
        "runAfter": {},
        "type": "ApiConnection",
        "inputs": {
          "body": {
            "Body": "<p>New Conditional Access Policy request submitted:<br><br><strong>Site URL:</strong> @{triggerBody()?['SiteURL']?['Url']}<br><strong>Policy Name:</strong> @{triggerBody()?['AuthenticationContextName']}<br><strong>Requested By:</strong> @{triggerBody()?['RequestedBy']?['DisplayName']}<br><strong>Priority:</strong> @{triggerBody()?['Priority']?['Value']}<br><strong>Justification:</strong> @{triggerBody()?['BusinessJustification']}<br><br><a href=\"https://yourtenant.sharepoint.com/sites/ITRequests/Lists/ConditionalAccessRequests\">Review Request</a></p>",
            "Subject": "New SharePoint Conditional Access Request - @{triggerBody()?['Title']}",
            "To": "sharepoint-admins@yourtenant.com"
          },
          "host": {
            "connection": {
              "name": "@parameters('$connections')['office365']['connectionId']"
            }
          },
          "method": "post",
          "path": "/v2/Mail"
        }
      },
      "Update_status_to_InReview": {
        "runAfter": {
          "Send_email_to_admins": [
            "Succeeded"
          ]
        },
        "type": "ApiConnection",
        "inputs": {
          "body": {
            "Status": {
              "Value": "In Review"
            }
          },
          "host": {
            "connection": {
              "name": "@parameters('$connections')['sharepointonline']['connectionId']"
            }
          },
          "method": "patch",
          "path": "/datasets/@{encodeURIComponent('https://yourtenant.sharepoint.com/sites/ITRequests')}/tables/@{encodeURIComponent('ConditionalAccessRequests')}/items/@{triggerBody()?['ID']}"
        }
      },
      "Post_to_Teams": {
        "runAfter": {
          "Update_status_to_InReview": [
            "Succeeded"
          ]
        },
        "type": "ApiConnection",
        "inputs": {
          "body": {
            "messageBody": "📋 **New SharePoint Conditional Access Request**\n\n**Site:** @{triggerBody()?['SiteURL']?['Url']}\n**Policy:** @{triggerBody()?['AuthenticationContextName']}\n**Requested By:** @{triggerBody()?['RequestedBy']?['DisplayName']}\n**Priority:** @{triggerBody()?['Priority']?['Value']}\n\n[Review Request](https://yourtenant.sharepoint.com/sites/ITRequests/Lists/ConditionalAccessRequests/DispForm.aspx?ID=@{triggerBody()?['ID']})",
            "recipient": {
              "channelId": "19:your-channel-id@thread.tacv2",
              "groupId": "your-teams-group-id"
            }
          },
          "host": {
            "connection": {
              "name": "@parameters('$connections')['teams']['connectionId']"
            }
          },
          "method": "post",
          "path": "/v3/beta/teams/@{encodeURIComponent('your-teams-group-id')}/channels/@{encodeURIComponent('19:your-channel-id@thread.tacv2')}/messages"
        }
      }
    }
  }
}
```

#### Flow 2: Execution Flow (Service Account)

```json
{
  "definition": {
    "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {},
    "triggers": {
      "When_an_item_is_modified": {
        "recurrence": {
          "frequency": "Second",
          "interval": 10
        },
        "splitOn": "@triggerBody()?['value']",
        "type": "ApiConnection",
        "inputs": {
          "host": {
            "connection": {
              "name": "@parameters('$connections')['sharepointonline']['connectionId']"
            }
          },
          "method": "get",
          "path": "/datasets/@{encodeURIComponent('https://yourtenant.sharepoint.com/sites/ITRequests')}/tables/@{encodeURIComponent('ConditionalAccessRequests')}/onupdateditems"
        }
      }
    },
    "actions": {
      "Condition_-_Status_is_Approved": {
        "actions": {
          "HTTP_-_Execute_SharePoint_Command": {
            "runAfter": {},
            "type": "Http",
            "inputs": {
              "authentication": {
                "audience": "https://graph.microsoft.com",
                "clientId": "@parameters('ServiceAccountClientId')",
                "secret": "@parameters('ServiceAccountSecret')",
                "tenant": "@parameters('TenantId')",
                "type": "ActiveDirectoryOAuth"
              },
              "body": {
                "conditionalAccessPolicy": "AuthenticationContext",
                "authenticationContextName": "@{triggerBody()?['AuthenticationContextName']}"
              },
              "headers": {
                "Content-Type": "application/json"
              },
              "method": "PATCH",
              "uri": "https://graph.microsoft.com/v1.0/sites/@{variables('SiteId')}"
            }
          },
          "Update_item_-_Mark_as_Completed": {
            "runAfter": {
              "HTTP_-_Execute_SharePoint_Command": [
                "Succeeded"
              ]
            },
            "type": "ApiConnection",
            "inputs": {
              "body": {
                "CompletedDate": "@utcnow()",
                "Status": {
                  "Value": "Completed"
                },
                "Comments": "Policy applied successfully via automated workflow"
              },
              "host": {
                "connection": {
                  "name": "@parameters('$connections')['sharepointonline']['connectionId']"
                }
              },
              "method": "patch",
              "path": "/datasets/@{encodeURIComponent('https://yourtenant.sharepoint.com/sites/ITRequests')}/tables/@{encodeURIComponent('ConditionalAccessRequests')}/items/@{triggerBody()?['ID']}"
            }
          },
          "Send_completion_email": {
            "runAfter": {
              "Update_item_-_Mark_as_Completed": [
                "Succeeded"
              ]
            },
            "type": "ApiConnection",
            "inputs": {
              "body": {
                "Body": "<p>Your SharePoint Conditional Access Policy request has been completed successfully.<br><br><strong>Site URL:</strong> @{triggerBody()?['SiteURL']?['Url']}<br><strong>Policy Applied:</strong> @{triggerBody()?['AuthenticationContextName']}<br><strong>Completed Date:</strong> @{utcnow()}<br><br>The conditional access policy is now active on your site.</p>",
                "Subject": "SharePoint Conditional Access Request Completed - @{triggerBody()?['Title']}",
                "To": "@{triggerBody()?['RequestedBy']?['Email']}"
              },
              "host": {
                "connection": {
                  "name": "@parameters('$connections')['office365']['connectionId']"
                }
              },
              "method": "post",
              "path": "/v2/Mail"
            }
          }
        },
        "runAfter": {},
        "expression": {
          "and": [
            {
              "equals": [
                "@triggerBody()?['Status']?['Value']",
                "Approved"
              ]
            }
          ]
        },
        "type": "If"
      }
    }
  }
}
```

### Phase 3: Service Account Setup

#### Azure AD App Registration for Service Account

```powershell
# PowerShell script to create Azure AD app registration
Connect-AzureAD

$appName = "SharePoint-ConditionalAccess-ServiceAccount"
$app = New-AzureADApplication -DisplayName $appName

# Create service principal
$sp = New-AzureADServicePrincipal -AppId $app.AppId

# Generate client secret
$passwordCred = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -CustomKeyIdentifier "PowerAutomate" -EndDate (Get-Date).AddYears(2)

Write-Host "App Registration Details:" -ForegroundColor Green
Write-Host "Application ID: $($app.AppId)" -ForegroundColor Yellow
Write-Host "Client Secret: $($passwordCred.Value)" -ForegroundColor Yellow
Write-Host "Object ID: $($app.ObjectId)" -ForegroundColor Yellow

# Required API Permissions (to be added manually in Azure Portal)
Write-Host "`nRequired API Permissions:" -ForegroundColor Cyan
Write-Host "- Microsoft Graph: Sites.FullControl.All" -ForegroundColor White
Write-Host "- SharePoint: Sites.FullControl.All" -ForegroundColor White
Write-Host "- Office 365 SharePoint Online: AllSites.FullControl" -ForegroundColor White
```

#### SharePoint Admin Role Assignment

```powershell
# Assign SharePoint Administrator role to service principal
Connect-MsolService

$servicePrincipalObjectId = "your-service-principal-object-id"
$sharePointAdminRole = Get-MsolRole | Where-Object {$_.Name -eq "SharePoint Service Administrator"}

Add-MsolRoleMember -RoleObjectId $sharePointAdminRole.ObjectId -RoleMemberType ServicePrincipal -RoleMemberObjectId $servicePrincipalObjectId

Write-Host "Service principal assigned SharePoint Administrator role" -ForegroundColor Green
```

### Phase 4: Power Apps Request Portal (Optional)

Create a Power Apps canvas app for better user experience:

```json
{
  "name": "SharePoint Conditional Access Request Portal",
  "screens": [
    {
      "name": "RequestForm",
      "controls": [
        {
          "type": "TextInput",
          "property": "SiteURL",
          "displayName": "SharePoint Site URL"
        },
        {
          "type": "Dropdown",
          "property": "AuthenticationContextName",
          "displayName": "Authentication Context",
          "items": [
            "Sensitive information - guest terms of use",
            "Highly confidential - MFA required",
            "External sharing - compliance check"
          ]
        },
        {
          "type": "TextInput",
          "property": "BusinessJustification",
          "displayName": "Business Justification",
          "mode": "Multiline"
        },
        {
          "type": "Dropdown",
          "property": "Priority",
          "displayName": "Priority Level",
          "items": ["Low", "Medium", "High", "Critical"]
        }
      ]
    }
  ]
}
```

## Benefits of Power Platform Approach

### For Site Owners:
- ✅ **Self-service capability** - Submit requests 24/7
- ✅ **Real-time status tracking** - Know where your request stands
- ✅ **Automated notifications** - Get updates via email/Teams
- ✅ **No admin privileges needed** - Works with standard permissions

### For SharePoint Admins:
- ✅ **Centralized request management** - All requests in one place
- ✅ **Automated execution** - Approved requests execute automatically
- ✅ **Audit trail** - Complete history of all changes
- ✅ **Workload reduction** - Less manual PowerShell execution

### For Organizations:
- ✅ **Compliance** - All changes documented and approved
- ✅ **Scalability** - Handles high volumes of requests
- ✅ **Integration** - Works with existing Microsoft 365 workflows
- ✅ **Cost-effective** - Uses existing licenses

## Deployment Timeline

| Week | Tasks | Responsible |
|------|-------|-------------|
| 1 | SharePoint list creation, permissions setup | SharePoint Admin |
| 2 | Azure AD app registration, API permissions | Azure Admin |
| 3 | Power Automate flows creation and testing | Power Platform Developer |
| 4 | User training and rollout | IT Training Team |

## Monitoring and Maintenance

### Key Metrics to Track:
- Request volume and trends
- Average approval time
- Success rate of automated executions
- User satisfaction scores

### Regular Maintenance Tasks:
- Review and renew service account certificates
- Update authentication context names as policies change
- Monitor flow execution history for errors
- Update user permissions as organizational structure changes

This Power Platform solution provides a professional, scalable way to handle SharePoint conditional access policy requests while maintaining proper governance and audit trails.