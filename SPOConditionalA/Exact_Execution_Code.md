# EXACT POWER AUTOMATE ACTION - Conditional Access Policy Execution

## 🎯 **THE PRECISE EXECUTION CODE**

This is the **exact Power Automate action** that applies the conditional access policy:

```json
{
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
  }
}
```

## 🔄 **COMPLETE FLOW EXECUTION SEQUENCE**

### **Flow 2: "SharePoint-ConditionalAccess-Executor"**

**Trigger:** SharePoint list item modified (Status = "Approved")
```json
{
  "When_an_item_is_modified": {
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
}
```

**Condition Check:** Only proceed if status is "Approved"
```json
{
  "Condition_-_Status_is_Approved": {
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
```

**🚀 ACTION 1: APPLY CONDITIONAL ACCESS POLICY** ⭐ **THIS IS IT!**
```json
{
  "HTTP_-_Execute_SharePoint_Command": {
    "runAfter": {},
    "type": "Http",
    "inputs": {
      "authentication": {
        "audience": "https://graph.microsoft.com",
        "clientId": "12345678-abcd-1234-efgh-123456789012",
        "secret": "your-service-account-secret",
        "tenant": "contoso.onmicrosoft.com",
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
  }
}
```

**ACTION 2: Update Status to Completed**
```json
{
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
      }
    }
  }
}
```

**ACTION 3: Send Completion Notification**
```json
{
  "Send_completion_email": {
    "runAfter": {
      "Update_item_-_Mark_as_Completed": [
        "Succeeded"
      ]
    },
    "type": "ApiConnection",
    "inputs": {
      "body": {
        "Body": "<p>Your SharePoint Conditional Access Policy request has been completed successfully.</p>",
        "Subject": "SharePoint Conditional Access Request Completed",
        "To": "@{triggerBody()?['RequestedBy']?['Email']}"
      }
    }
  }
}
```

## 🔍 **RUNTIME VARIABLE RESOLUTION**

When the flow executes, these dynamic values get resolved:

| Variable | Runtime Value | Example |
|----------|---------------|---------|
| `@parameters('ServiceAccountClientId')` | Azure AD App Client ID | `12345678-abcd-1234-efgh-123456789012` |
| `@parameters('ServiceAccountSecret')` | Azure AD App Secret | `abc123def456ghi789...` |
| `@parameters('TenantId')` | Azure AD Tenant ID | `contoso.onmicrosoft.com` |
| `@{triggerBody()?['AuthenticationContextName']}` | Policy name from request | `"Sensitive information - guest terms of use"` |
| `@{variables('SiteId')}` | SharePoint Site Graph ID | `contoso.sharepoint.com,29cc8a03,1a3be1b9` |

## 🌐 **ACTUAL HTTP REQUEST SENT**

When Power Automate executes this action, here's the **actual HTTP request** sent to Microsoft Graph:

```http
PATCH https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,29cc8a03-a8fd-4d2c-ab6b-86b9cb7ad16f,1a3be1b9-7f2e-4c58-9b3d-4a5e6c7f8d9a HTTP/1.1
Host: graph.microsoft.com
Authorization: Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1yNS1BVWliZkFpaTdqOGFiTWxsM0I4Q3ZXTSIsImtpZCI6Ik1yNS1BVWliZkFpaTdqOGFiTWxsM0I4Q3ZXTSJ9...
Content-Type: application/json
Content-Length: 125
User-Agent: Microsoft-PowerAutomate/1.0

{
  "conditionalAccessPolicy": "AuthenticationContext",
  "authenticationContextName": "Sensitive information - guest terms of use"
}
```

## 🔄 **EQUIVALENT POWERSHELL COMMAND**

This Graph API call is equivalent to running this PowerShell command:

```powershell
Set-SPOSite -Identity "https://contoso.sharepoint.com/sites/research" -ConditionalAccessPolicy AuthenticationContext -AuthenticationContextName "Sensitive information - guest terms of use"
```

**But instead of requiring admin PowerShell access, it uses:**
- ✅ **Service Account** with proper permissions
- ✅ **Microsoft Graph API** (modern, supported)
- ✅ **Power Automate** (no server management)
- ✅ **Automated execution** (no manual intervention)

## 🎯 **KEY TECHNICAL POINTS**

1. **Technology**: **Microsoft Graph API** (NOT PowerShell)
2. **Method**: **HTTP PATCH request** to Graph endpoint
3. **Authentication**: **Azure AD OAuth 2.0** with service principal
4. **Execution Platform**: **Power Automate HTTP connector**
5. **Trigger**: **SharePoint list item status change**
6. **Permission Model**: **Application permissions** (Sites.FullControl.All)

## 🔧 **WHY THIS APPROACH IS BETTER**

| Traditional PowerShell | Power Automate + Graph API |
|----------------------|---------------------------|
| Requires admin to run manually | Automated execution |
| Need PowerShell modules installed | No modules needed |
| Risk of human error | Automated, consistent |
| No audit trail | Complete audit trail |
| Manual notification | Automated notifications |
| Single-threaded execution | Parallel processing |
| Server/workstation dependency | Cloud-native |

The conditional access policy is applied through **Microsoft Graph API's HTTP PATCH method** executed by **Power Automate's HTTP connector** using a **service account with SharePoint admin permissions**. This is modern, scalable, and requires no PowerShell execution on any servers or workstations.