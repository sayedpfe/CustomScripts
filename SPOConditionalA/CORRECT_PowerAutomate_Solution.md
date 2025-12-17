# Power Automate - SharePoint REST API Solution for Conditional Access
# This demonstrates the ACTUAL way to set conditional access policies in Power Automate

## ❌ IMPORTANT DISCOVERY: Microsoft Graph API Does NOT Support This!

After testing, we discovered that:
- **Microsoft Graph API `/v1.0/sites/{id}` endpoint does NOT support setting conditional access policies**
- The `Update-MgSite` cmdlet accepts the parameters but doesn't actually apply them
- This is a known limitation of the Graph API sites endpoint

## ✅ ACTUAL Power Automate Solutions

### **Solution 1: SharePoint HTTP Request (Recommended for Power Automate)**

Power Automate Action: **Send an HTTP request to SharePoint**

```json
{
  "siteUrl": "https://yourtenant.sharepoint.com/sites/sitename",
  "method": "POST",
  "uri": "_api/site",
  "headers": {
    "Accept": "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    "X-HTTP-Method": "MERGE",
    "IF-MATCH": "*"
  },
  "body": {
    "ConditionalAccessPolicy": "AuthenticationContext",
    "AuthenticationContextName": "Sensitive Information - Guest terms of Use"
  }
}
```

**Issue**: This may also not be supported in SharePoint REST API.

---

### **Solution 2: Azure Automation Runbook (Most Reliable)**

This is the **production-ready approach** for Power Automate:

#### **Step 1: Create Azure Automation Account**
1. Create an Automation Account in Azure Portal
2. Add PowerShell modules: `Microsoft.Online.SharePoint.PowerShell` or `PnP.PowerShell`
3. Create a Managed Identity with SharePoint Admin permissions

#### **Step 2: Create Runbook**

```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$AuthenticationContextName
)

# Connect using Managed Identity
Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -ManagedIdentity

# Apply conditional access policy
Set-PnPSite -Identity $SiteUrl `
            -ConditionalAccessPolicy AuthenticationContext `
            -AuthenticationContextName $AuthenticationContextName

Write-Output "Successfully applied conditional access policy to $SiteUrl"
```

#### **Step 3: Power Automate Flow**

```json
{
  "trigger": {
    "type": "SharePoint - When item is modified",
    "condition": "Status = Approved"
  },
  "actions": [
    {
      "type": "HTTP - Call Azure Automation Webhook",
      "method": "POST",
      "uri": "https://webhookUrl",
      "body": {
        "SiteUrl": "@{triggerBody()?['SiteURL']}",
        "AuthenticationContextName": "@{triggerBody()?['AuthenticationContextName']}"
      }
    }
  ]
}
```

---

### **Solution 3: Azure Function (Alternative)**

Create an Azure Function with HTTP trigger:

```csharp
[FunctionName("SetConditionalAccess")]
public static async Task<IActionResult> Run(
    [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req,
    ILogger log)
{
    string siteUrl = req.Query["siteUrl"];
    string policyName = req.Query["policyName"];
    
    // Use SharePoint CSOM or PnP Framework
    using (var context = new ClientContext(siteUrl))
    {
        // Authenticate with app-only permissions
        context.ExecutingWebRequest += (s, e) => {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + GetAccessToken();
        };
        
        var site = context.Site;
        site.ConditionalAccessPolicy = SPOConditionalAccessPolicyType.AuthenticationContext;
        site.AuthenticationContextName = policyName;
        site.Update();
        context.ExecuteQuery();
    }
    
    return new OkObjectResult("Policy applied");
}
```

---

## 📊 **Comparison: What Works in Power Automate**

| Method | Works? | Requires | Complexity |
|--------|--------|----------|------------|
| **Microsoft Graph API** | ❌ **NO** | N/A | N/A |
| **SharePoint REST API** | ⚠️ **Unlikely** | SharePoint permissions | Medium |
| **Azure Automation + PowerShell** | ✅ **YES** | Azure Automation Account | Medium |
| **Azure Function + CSOM** | ✅ **YES** | Azure Function App | High |
| **Power Automate → Call PowerShell Script** | ❌ **NO** | Not supported | N/A |

---

## 🎯 **RECOMMENDED PRODUCTION SOLUTION**

### **Azure Automation Runbook Approach**

**Architecture:**
```
Power Automate Flow
    ↓ (HTTP webhook)
Azure Automation Runbook
    ↓ (PowerShell Set-PnPSite)
SharePoint Admin API
    ↓
Site Conditional Access Policy Applied
```

**Why This Works:**
1. ✅ Azure Automation can run PowerShell with admin privileges
2. ✅ Uses Managed Identity (no secrets to manage)
3. ✅ Fully supported by Microsoft
4. ✅ Can be triggered from Power Automate via webhook
5. ✅ This is **exactly what `Set-SPOSite` does** - it calls SharePoint Admin API

---

## 🔧 **Quick Setup Guide for Azure Automation**

### **1. Create Automation Account**
```powershell
az automation account create `
    --name "SPO-ConditionalAccess-Automation" `
    --resource-group "YourResourceGroup" `
    --location "EastUS"
```

### **2. Enable Managed Identity**
```powershell
az automation account update `
    --name "SPO-ConditionalAccess-Automation" `
    --resource-group "YourResourceGroup" `
    --identity "[SystemAssigned]"
```

### **3. Assign SharePoint Admin Role**
```powershell
# Get the Managed Identity Object ID
$identityId = (Get-AzAutomationAccount -Name "SPO-ConditionalAccess-Automation" -ResourceGroupName "YourResourceGroup").Identity.PrincipalId

# Assign SharePoint Administrator role
Connect-AzureAD
$role = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq "SharePoint Administrator"}
Add-AzureADDirectoryRoleMember -ObjectId $role.ObjectId -RefObjectId $identityId
```

### **4. Create and Publish Runbook**
See runbook code above

### **5. Create Webhook**
```powershell
$webhook = New-AzAutomationWebhook `
    -Name "ApplyConditionalAccess" `
    -RunbookName "Set-SiteConditionalAccess" `
    -IsEnabled $true `
    -ExpiryTime (Get-Date).AddYears(1) `
    -AutomationAccountName "SPO-ConditionalAccess-Automation" `
    -ResourceGroupName "YourResourceGroup"

Write-Host "Webhook URL: $($webhook.WebhookURI)"
```

### **6. Use in Power Automate**
Add HTTP action with webhook URL

---

## 💡 **KEY INSIGHT**

The original solution I provided **was incorrect** about Graph API supporting conditional access policies on SharePoint sites. 

**The Truth:**
- ❌ Microsoft Graph API **cannot** set conditional access policies on SharePoint sites
- ✅ **Only SharePoint Admin API** (what Set-SPOSite uses) can do this
- ✅ For Power Automate, you **must** use Azure Automation or Azure Functions
- ✅ This is still automatable and doesn't require manual intervention

**Updated Architecture:**
```
Site Owner Request (SharePoint List)
    ↓
Power Automate (Trigger + Approval)
    ↓
Azure Automation Runbook (PowerShell)
    ↓
SharePoint Admin API
    ↓
Conditional Access Policy Applied
```

This is still a valid self-service solution, but requires Azure Automation as the execution layer.
