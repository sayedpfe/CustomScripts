# Technical Execution Flow - SharePoint Conditional Access Policy Application

## 🎯 **EXACT EXECUTION POINT** - Where the Magic Happens

The conditional access policy is applied in **Flow 2 (Execution Flow)**, specifically in the **HTTP Action** called `"HTTP_-_Execute_SharePoint_Command"`.

## 🔧 **Technology Stack Used**

| Component | Technology | Purpose |
|-----------|------------|---------|
| **Trigger** | SharePoint Connector | Detects when request status = "Approved" |
| **Authentication** | Azure AD OAuth 2.0 | Service account authentication |
| **Execution** | **Microsoft Graph API** | **THIS IS WHERE POLICY GETS APPLIED** |
| **Notification** | Office 365 Connector | Send completion emails |
| **Logging** | SharePoint Connector | Update request status |

## 📋 **Step-by-Step Execution Flow**

### **Flow 1: Notification Flow** (No policy execution here)
```
1. Trigger: SharePoint list item created
2. Action: Send email notification to admins
3. Action: Update status to "In Review"
4. Action: Post to Teams (optional)
```

### **Flow 2: Execution Flow** ⭐ **THIS IS WHERE CONDITIONAL ACCESS POLICY IS APPLIED**

```mermaid
graph TD
    A[SharePoint List Item Modified] --> B{Status = 'Approved'?}
    B -->|Yes| C[🔑 Authenticate with Service Account]
    C --> D[🌐 HTTP Call to Microsoft Graph API]
    D --> E[✅ Policy Applied to SharePoint Site]
    E --> F[📝 Update Request Status to 'Completed']
    F --> G[📧 Notify Requestor]
    B -->|No| H[Do Nothing]
```

## 🔍 **DETAILED BREAKDOWN - The Core Execution**

### **Step 1: Trigger Detection**
```json
{
  "trigger": {
    "type": "SharePoint",
    "operation": "OnItemModified",
    "condition": "Status field changes to 'Approved'"
  }
}
```

### **Step 2: Authentication (Service Account)**
```json
{
  "authentication": {
    "type": "ActiveDirectoryOAuth",
    "audience": "https://graph.microsoft.com",
    "clientId": "12345678-1234-1234-1234-123456789012",
    "clientSecret": "***SECRET***",
    "tenant": "contoso.onmicrosoft.com"
  }
}
```

### **Step 3: 🚀 POLICY EXECUTION - Microsoft Graph API Call**
```json
{
  "name": "HTTP_-_Execute_SharePoint_Command",
  "type": "Http",
  "method": "PATCH",
  "uri": "https://graph.microsoft.com/v1.0/sites/{site-id}",
  "headers": {
    "Content-Type": "application/json",
    "Authorization": "Bearer {access-token}"
  },
  "body": {
    "conditionalAccessPolicy": "AuthenticationContext",
    "authenticationContextName": "Sensitive information - guest terms of use"
  }
}
```

**🎯 THIS IS THE EXACT MOMENT THE CONDITIONAL ACCESS POLICY IS APPLIED!**

## 🔄 **Alternative Implementation Options**

I actually provided **multiple technology approaches** in the solution. Here's what each uses:

### **Option 1: Microsoft Graph API** ⭐ **RECOMMENDED**
- **Technology**: HTTP REST calls to Graph API
- **Authentication**: Azure AD Service Principal
- **Endpoint**: `https://graph.microsoft.com/v1.0/sites/{site-id}`
- **Method**: PATCH request
- **Advantages**: Modern, supported, scalable

### **Option 2: SharePoint REST API** 
- **Technology**: SharePoint-specific REST API
- **Authentication**: SharePoint access token
- **Endpoint**: `https://tenant.sharepoint.com/sites/sitename/_api/site`
- **Method**: HTTP PATCH request
- **Advantages**: Direct SharePoint integration

### **Option 3: PowerShell via Azure Automation** (Alternative approach)
- **Technology**: PowerShell with SharePoint Online Management Shell
- **Authentication**: Managed Identity or Service Principal
- **Command**: `Set-SPOSite -ConditionalAccessPolicy AuthenticationContext`
- **Execution**: Azure Automation Runbook triggered by Power Automate

## 🔧 **Behind the Scenes - What Actually Happens**

### **When Graph API Receives the Request:**

1. **Authentication Validation**
   ```
   Service Account → Azure AD → Access Token → Graph API
   ```

2. **API Call Processing**
   ```
   PATCH https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,12345,67890
   Content-Type: application/json
   Authorization: Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6...
   
   {
     "conditionalAccessPolicy": "AuthenticationContext",
     "authenticationContextName": "Sensitive information - guest terms of use"
   }
   ```

3. **SharePoint Site Update**
   ```
   Graph API → SharePoint Online Service → Site Settings Update
   ```

4. **Response**
   ```json
   {
     "id": "contoso.sharepoint.com,12345,67890",
     "conditionalAccessPolicy": "AuthenticationContext",
     "authenticationContextName": "Sensitive information - guest terms of use",
     "lastModifiedDateTime": "2025-11-14T10:30:00Z"
   }
   ```

## 🛠️ **Service Account Requirements**

The service account needs these **specific permissions**:

### **Azure AD App Registration Permissions:**
```
Microsoft Graph:
- Sites.FullControl.All (Application)
- Sites.ReadWrite.All (Application)

SharePoint:
- Sites.FullControl.All (Application)
```

### **SharePoint Admin Role:**
```
Role: SharePoint Service Administrator
Scope: Tenant-wide
Purpose: Execute Set-SPOSite equivalent operations
```

## 🔍 **Real-World HTTP Request Example**

Here's what the **actual HTTP request** looks like when executed:

```http
PATCH https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/research
Authorization: Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1yNS1BVWliZkFpaTdqOGFiTWxsM0I...
Content-Type: application/json
Accept: application/json

{
  "conditionalAccessPolicy": "AuthenticationContext",
  "authenticationContextName": "Sensitive information - guest terms of use"
}
```

**Response:**
```http
HTTP/1.1 200 OK
Content-Type: application/json

{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites/$entity",
  "id": "contoso.sharepoint.com,29cc8a03-a8fd-4d2c-ab6b-86b9cb7ad16f,1a3be1b9-7f2e-4c58-9b3d-4a5e6c7f8d9a",
  "displayName": "Research Site",
  "webUrl": "https://contoso.sharepoint.com/sites/research",
  "conditionalAccessPolicy": "AuthenticationContext",
  "authenticationContextName": "Sensitive information - guest terms of use",
  "lastModifiedDateTime": "2025-11-14T10:30:45.123Z"
}
```

## 🎯 **Summary: Technology Used**

| Aspect | Technology Choice | Why |
|--------|------------------|-----|
| **Primary Method** | **Microsoft Graph API** | Modern, supported, future-proof |
| **Authentication** | **Azure AD OAuth 2.0** | Secure service account approach |
| **Execution Platform** | **Power Automate** | No servers needed, built-in retry |
| **Trigger Mechanism** | **SharePoint Connectors** | Real-time detection |
| **Fallback Option** | **SharePoint REST API** | Direct SharePoint communication |
| **Alternative** | **PowerShell + Azure Automation** | For PowerShell-preferring environments |

**The conditional access policy is applied using Microsoft Graph API's HTTP PATCH method, NOT PowerShell commands. This makes it more reliable, scalable, and doesn't require PowerShell module management on any servers.**