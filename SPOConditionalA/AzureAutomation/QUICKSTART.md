# Quick Start Guide - Azure Automation for SharePoint Conditional Access

## 🚀 Deploy in 15 Minutes

### Prerequisites
- Azure Subscription with Contributor role
- PowerShell 7.x or Windows PowerShell 5.1

### Step-by-Step Deployment

#### 1️⃣ Setup Azure Automation (10 minutes)
```powershell
cd "D:\OneDrive\OneDrive - Microsoft\Documents\Learning Projects\CustomScripts\SPOConditionalA\AzureAutomation"
.\1-Setup-AutomationAccount.ps1
```
Wait for completion. Note: Module import runs in background (~10 min).

#### 2️⃣ Create Runbook (2 minutes)
```powershell
.\2-Create-Runbook.ps1
```

#### 3️⃣ Setup Webhook (1 minute)
```powershell
.\3-Setup-Webhook.ps1
```
**⚠️ SAVE THE WEBHOOK URL!** (shown only once)

#### 4️⃣ Test (Optional - 2 minutes)
```powershell
.\4-Test-Webhook.ps1
```

#### 5️⃣ Create SharePoint List
```powershell
.\5-Create-SharePointList.ps1
# Then run the generated script:
.\Create-SharePointList.ps1 -SiteUrl "https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE"
```

#### 6️⃣ Setup Power Automate Flow
1. Go to https://make.powerautomate.com
2. Create flow with these actions:
   - **Trigger:** When item modified in ConditionalAccessRequests list
   - **Condition:** If Status = "Approved"
   - **HTTP Action:**
     - Method: POST
     - URI: `<your webhook URL>`
     - Body:
       ```json
       {
         "SiteUrl": "@triggerBody()?['SiteURL']",
         "AuthenticationContextName": "@triggerBody()?['AuthenticationContextName']",
         "RequestorEmail": "@triggerBody()?['RequestorEmail']",
         "RequestId": "@triggerBody()?['ID']"
       }
       ```
   - **Update item:** Status = "Completed"
   - **Send email:** Notify requestor

### ✅ Test End-to-End

1. Create item in SharePoint list:
   - SiteURL: `https://m365cpi90282478.sharepoint.com/sites/DeutschKurs`
   - AuthenticationContextName: `Sensitive Information - Guest terms of Use`
   - RequestorEmail: your-email@domain.com
   - Status: `Pending`

2. Change Status to `Approved`

3. Wait 60 seconds

4. Check Status = `Completed`

5. Verify policy applied to site!

### 📂 Files Created
```
AzureAutomation/
├── 1-Setup-AutomationAccount.ps1
├── 2-Create-Runbook.ps1
├── 3-Setup-Webhook.ps1
├── 4-Test-Webhook.ps1
├── 5-Create-SharePointList.ps1
├── Create-SharePointList.ps1 (generated)
├── PowerAutomate-Flow-Template.json
├── README.md (full documentation)
├── QUICKSTART.md (this file)
├── automation-config.json (generated - config)
└── webhook-url-SECURE.txt (generated - KEEP SECURE!)
```

### 🆘 Troubleshooting

**Runbook fails with module error:**
- Wait 10 minutes for PnP.PowerShell import
- Check: Portal → Automation Account → Modules

**Unauthorized errors:**
- Verify Managed Identity has SharePoint Admin role
- Check: Portal → Entra ID → Roles → SharePoint Administrator

**Webhook 404:**
- Webhook expired or wrong URL
- Run `3-Setup-Webhook.ps1` again

### 📖 Full Documentation
See [README.md](README.md) for complete details, architecture, security, and troubleshooting.

---

**That's it! You now have automated conditional access policy application! 🎉**
