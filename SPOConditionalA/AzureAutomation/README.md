# Azure Automation Solution for SharePoint Conditional Access Policies
## Complete Deployment Guide

---

## 📋 Overview

This solution enables **Site Owners** to apply Conditional Access Policies to SharePoint sites through an automated self-service workflow, without requiring SharePoint Administrator privileges.

**Architecture:**
```
Site Owner → SharePoint List → Power Automate → Azure Automation → SharePoint Admin API
```

**Key Benefits:**
- ✅ Fully automated - no manual admin intervention
- ✅ Self-service for Site Owners
- ✅ Secure - uses Managed Identity
- ✅ Auditable - all requests tracked in SharePoint list
- ✅ Production-ready

---

## 🎯 What Gets Deployed

1. **Azure Automation Account** - Executes PowerShell with admin privileges
2. **Managed Identity** - Secure authentication without passwords
3. **PowerShell Runbook** - Applies conditional access policies
4. **Webhook** - Allows Power Automate to trigger the runbook
5. **SharePoint List** - Tracks requests and approvals
6. **Power Automate Flow** - Orchestrates the entire workflow

---

## 📦 Prerequisites

### Azure Subscription
- Contributor role (to create Automation Account)
- Ability to assign Entra ID roles (for Managed Identity)

### PowerShell Modules (will be installed automatically)
- Az.Accounts
- Az.Automation
- Az.Resources
- Microsoft.Graph.Identity.DirectoryManagement
- PnP.PowerShell

### Permissions Required
- Azure Subscription Contributor
- Entra ID Role Administrator (to assign SharePoint Administrator role)
- SharePoint Site Collection Administrator (to create tracking list)

---

## 🚀 Deployment Steps

### Step 1: Setup Azure Automation Account

Run the first script to create the Azure Automation infrastructure:

```powershell
cd "D:\OneDrive\OneDrive - Microsoft\Documents\Learning Projects\CustomScripts\SPOConditionalA\AzureAutomation"

.\1-Setup-AutomationAccount.ps1
```

**What it does:**
- Creates Resource Group
- Creates Automation Account with System-Assigned Managed Identity
- Imports PnP.PowerShell module (takes 5-10 minutes)
- Assigns SharePoint Administrator role to Managed Identity
- Saves configuration to `automation-config.json`

**Time:** ~10-15 minutes (including module import)

---

### Step 2: Create PowerShell Runbook

Run the second script to create the runbook that applies policies:

```powershell
.\2-Create-Runbook.ps1
```

**What it does:**
- Creates runbook: `Apply-SiteConditionalAccess`
- Publishes the runbook
- Updates configuration file

**Time:** ~2 minutes

**The runbook:**
- Accepts webhook input from Power Automate
- Connects to SharePoint using Managed Identity
- Applies conditional access policy using `Set-PnPSite`
- Verifies the policy was applied
- Returns detailed success/error message

---

### Step 3: Setup Webhook

Run the third script to create a webhook URL:

```powershell
.\3-Setup-Webhook.ps1
```

**What it does:**
- Creates webhook for the runbook
- Displays webhook URL (**save this - shown only once!**)
- Saves URL to `webhook-url-SECURE.txt`
- Webhook expires in 2 years

**Time:** ~1 minute

⚠️ **IMPORTANT:** Copy and save the webhook URL immediately. It cannot be retrieved later.

---

### Step 4: Test the Webhook (Optional)

Test that the webhook works correctly:

```powershell
.\4-Test-Webhook.ps1
```

**What it does:**
- Sends test request to webhook with sample data
- Monitors job execution
- Displays runbook output
- Verifies policy was applied to test site

**Time:** ~1-2 minutes

**Test parameters:**
- Site URL: `https://m365cpi90282478.sharepoint.com/sites/DeutschKurs`
- Auth Context: `Sensitive Information - Guest terms of Use`

---

### Step 5: Create SharePoint List

Run the fifth script to create the tracking list:

```powershell
.\5-Create-SharePointList.ps1
```

This generates:
- `SharePoint-List-Schema.txt` - Column definitions
- `Create-SharePointList.ps1` - Deployment script

To create the list on your site:

```powershell
.\Create-SharePointList.ps1 -SiteUrl "https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE"
```

**List columns:**
- Title, SiteURL, AuthenticationContextName, RequestorEmail
- BusinessJustification, Status, RequestDate
- ProcessingStarted, ProcessingCompleted, ApprovedBy
- ApprovalDate, Result

---

### Step 6: Import Power Automate Flow

#### Option A: Manual Creation

1. Go to https://make.powerautomate.com
2. Create new **Automated cloud flow**
3. Trigger: **SharePoint - When an item is created or modified**
   - Site: Your tracking list site
   - List: `ConditionalAccessRequests`
4. Add condition: **Status = Approved**
5. Add action: **HTTP**
   - Method: `POST`
   - URI: `<paste webhook URL from Step 3>`
   - Headers: `Content-Type: application/json`
   - Body:
     ```json
     {
       "SiteUrl": "@triggerBody()?['SiteURL']",
       "AuthenticationContextName": "@triggerBody()?['AuthenticationContextName']",
       "RequestorEmail": "@triggerBody()?['RequestorEmail']",
       "RequestId": "@triggerBody()?['ID']"
     }
     ```
6. Add action: **Update item** (set Status to Completed)
7. Add action: **Send email** (notify requestor)

#### Option B: Import Template

1. Go to https://make.powerautomate.com
2. **My flows** → **Import** → **Import Package (Legacy)**
3. Select `PowerAutomate-Flow-Template.json`
4. Update these values in the template:
   - Replace `YOUR-TENANT` with your tenant name
   - Replace `YOUR-SITE` with your site name
   - Replace `YOUR-LIST-ID` with your list ID
   - Replace `PASTE-YOUR-WEBHOOK-URL-HERE` with webhook URL
5. Save and test

---

## 🧪 Testing the Complete Solution

### Test Scenario: Apply Policy to DeutschKurs Site

1. **Create a request in SharePoint list:**
   - Title: "Apply Conditional Access to DeutschKurs"
   - SiteURL: `https://m365cpi90282478.sharepoint.com/sites/DeutschKurs`
   - AuthenticationContextName: `Sensitive Information - Guest terms of Use`
   - RequestorEmail: `your-email@domain.com`
   - BusinessJustification: "Contains sensitive training materials"
   - Status: `Pending`

2. **Approve the request:**
   - Change Status to `Approved`
   - Fill in ApprovedBy and ApprovalDate

3. **Power Automate flow triggers:**
   - Detects status change to Approved
   - Updates status to Processing
   - Calls Azure Automation webhook

4. **Azure Automation executes:**
   - Runbook receives webhook data
   - Connects to SharePoint with Managed Identity
   - Applies conditional access policy
   - Returns success/error result

5. **Power Automate completes:**
   - Updates status to Completed
   - Sends email notification to requestor

6. **Verify:**
   - Check Status = Completed in list
   - Check Result column for success message
   - Test site access (should prompt for MFA or block based on policy)

---

## 📊 Monitoring and Troubleshooting

### Check Azure Automation Job Status

**Azure Portal:**
1. Navigate to Automation Account → **Jobs**
2. Find recent job for `Apply-SiteConditionalAccess`
3. Click job to view detailed output
4. Check **Output**, **Errors**, **All Logs**

**PowerShell:**
```powershell
$config = Get-Content ".\automation-config.json" | ConvertFrom-Json

Get-AzAutomationJob `
    -ResourceGroupName $config.ResourceGroupName `
    -AutomationAccountName $config.AutomationAccountName `
    -RunbookName $config.RunbookName |
    Sort-Object StartTime -Descending |
    Select-Object -First 5
```

### Common Issues

#### Module Import Still Running
**Symptom:** Runbook fails with "PnP.PowerShell module not found"
**Solution:** Wait for module import to complete (check Modules blade)

#### Managed Identity Not Assigned Role
**Symptom:** "Access denied" or "Unauthorized" errors
**Solution:** Manually assign SharePoint Administrator role in Entra ID

#### Webhook Expired
**Symptom:** Power Automate gets 404 error
**Solution:** Run `3-Setup-Webhook.ps1` again to create new webhook

#### Wrong Authentication Context Name
**Symptom:** "Authentication context not found"
**Solution:** Verify exact name in Entra ID → Protection → Conditional Access → Authentication contexts

---

## 🔒 Security Considerations

### Webhook URL Security
- Webhook URL contains authentication token
- Anyone with the URL can trigger the runbook
- Store `webhook-url-SECURE.txt` securely
- Don't commit to source control
- Rotate webhook every 6-12 months

### Managed Identity Permissions
- Has SharePoint Administrator role (tenant-wide)
- Can modify ANY SharePoint site
- Consider using custom role with limited permissions
- Monitor Audit Logs for unexpected changes

### SharePoint List Permissions
- Only IT/Admins should approve requests
- Site Owners can submit requests
- Break permission inheritance on list
- Grant approve permission only to specific group

---

## 📈 Scaling and Enhancements

### Multi-Tenant Support
- Deploy separate Automation Account per tenant
- Use tenant-specific configuration files
- Update webhook URLs in respective flows

### Approval Workflow
- Add multi-stage approval (manager → IT admin)
- Implement approval expiration
- Auto-reject after X days

### Notifications
- Email notifications for status changes
- Teams notifications via adaptive cards
- SMS for urgent requests

### Reporting
- Power BI dashboard for request metrics
- Track most common authentication contexts
- Identify sites with policies applied

---

## 💰 Cost Estimate

### Azure Automation
- Automation Account: **Free** (first 500 minutes/month)
- Additional minutes: **$0.002/minute**
- Typical usage: 20 requests/month × 1 minute = **$0.04/month**

### Storage
- Configuration files: **Negligible**

### Power Automate
- Included in Office 365 licenses for standard actions
- HTTP Premium connector may require Power Automate Premium

**Total estimated cost: ~$0.05 - $5/month** (depending on volume)

---

## 📚 Additional Resources

### Microsoft Documentation
- [Azure Automation Overview](https://learn.microsoft.com/azure/automation/automation-intro)
- [SharePoint Conditional Access](https://learn.microsoft.com/sharepoint/authentication-context-conditional-access)
- [Power Automate with Azure Automation](https://learn.microsoft.com/power-automate/desktop-flows/run-desktop-flows-with-azure-automation)

### PowerShell Cmdlets
- [Set-PnPSite](https://pnp.github.io/powershell/cmdlets/Set-PnPSite.html)
- [Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html)

---

## 🆘 Support

### Get Help
- Check Azure Automation job logs
- Review Power Automate run history
- Test webhook using `4-Test-Webhook.ps1`
- Verify Managed Identity role assignments

### Logs Location
- **Azure Automation:** Portal → Automation Account → Jobs → Select job → All Logs
- **Power Automate:** Portal → My flows → Select flow → Run history
- **SharePoint List:** Track Status and Result columns

---

## ✅ Deployment Checklist

- [ ] Azure Subscription with Contributor role
- [ ] Entra ID permissions to assign roles
- [ ] Run `1-Setup-AutomationAccount.ps1`
- [ ] Wait for PnP.PowerShell module import (5-10 min)
- [ ] Verify Managed Identity has SharePoint Admin role
- [ ] Run `2-Create-Runbook.ps1`
- [ ] Run `3-Setup-Webhook.ps1`
- [ ] **SAVE WEBHOOK URL**
- [ ] Test with `4-Test-Webhook.ps1`
- [ ] Create SharePoint list with `5-Create-SharePointList.ps1`
- [ ] Import or create Power Automate flow
- [ ] Update flow with webhook URL
- [ ] Test end-to-end with sample request
- [ ] Document webhook URL location for team
- [ ] Set calendar reminder to rotate webhook (2 years)

---

## 🎉 Success!

You now have a fully automated, self-service solution for applying Conditional Access Policies to SharePoint sites!

**Site Owners can:**
- Submit requests via SharePoint list
- No admin privileges required
- Automated approval and processing

**IT Admins can:**
- Review and approve requests
- Full audit trail
- No manual PowerShell execution

**Next Steps:**
- Train Site Owners on request process
- Document available Authentication Contexts
- Set up approval group in SharePoint list
- Monitor first few requests closely

---

**Created:** December 2025  
**Version:** 1.0  
**Contact:** IT Automation Team
