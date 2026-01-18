# Manual Entra ID App Registration Setup Guide

Since automated app registration had issues, follow these steps to manually create the app:

## Step 1: Create App Registration

1. Go to [Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)

2. Click **"+ New registration"**

3. Fill in the details:
   - **Name**: `SharePoint Site Export Tool`
   - **Supported account types**: Accounts in this organizational directory only (Single tenant)
   - **Redirect URI**: 
     - Platform: **Public client/native (mobile & desktop)**
     - URI: `http://localhost`

4. Click **"Register"**

5. **SAVE THE APPLICATION (CLIENT) ID** - You'll need this!

## Step 2: Configure API Permissions

1. In your new app, go to **"API permissions"** in the left menu

2. Click **"+ Add a permission"**

3. Select **"SharePoint"**

4. Click **"Delegated permissions"**

5. Check these permissions:
   - ☑️ `AllSites.FullControl` - Have full control of all site collections
   - ☑️ `AllSites.Read` - Read items in all site collections

6. Click **"Add permissions"**

7. Click **"Grant admin consent for [Your Organization]"** button

8. Confirm by clicking **"Yes"**

## Step 3: Enable Public Client Flow

1. In your app, go to **"Authentication"** in the left menu

2. Scroll down to **"Advanced settings"**

3. Under **"Allow public client flows"**, set to **"Yes"**

4. Click **"Save"**

## Step 4: Save Your Configuration

Create a file named `AppConfig.json` in the OffboardingAgentSolution folder with this content:

```json
{
  "AppName": "SharePoint Site Export Tool",
  "ClientId": "YOUR-CLIENT-ID-HERE",
  "TenantId": "b22f8675-8375-455b-941a-67bee4cf7747",
  "RedirectUri": "http://localhost",
  "CreatedDate": "2025-12-18"
}
```

**Replace `YOUR-CLIENT-ID-HERE` with the Application (client) ID from Step 1!**

## Step 5: Test the Export Script

Once the app is configured, run:

```powershell
cd "d:\OneDrive\OneDrive - Microsoft\Documents\Learning Projects\CustomScripts\OffboardingAgentSolution"

.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export" `
    -IncludeData
```

The script will automatically load the ClientId from AppConfig.json.

## Alternative: Use ClientId Parameter Directly

If you prefer not to create the config file, you can provide the ClientId directly:

```powershell
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export" `
    -ClientId "your-client-id-here" `
    -IncludeData
```

## Troubleshooting

**Issue**: "AADSTS65001: The user or administrator has not consented"
- **Solution**: Go back to Step 2 and ensure admin consent was granted

**Issue**: "AADSTS700016: Application with identifier was not found"
- **Solution**: Verify the ClientId in AppConfig.json matches the one from Azure Portal

**Issue**: "Invalid redirect URI"
- **Solution**: Ensure redirect URI is exactly `http://localhost` and platform is "Public client"

---

**Your Tenant ID**: `b22f8675-8375-455b-941a-67bee4cf7747`

Once setup is complete, you'll be ready to export and deploy your SharePoint sites!
