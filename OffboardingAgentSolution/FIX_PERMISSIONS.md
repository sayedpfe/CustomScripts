# Fix App Permissions - Add Delegated Permissions

## Problem
Your app has **APPLICATION permissions** but needs **DELEGATED permissions** for interactive login.

## Solution: Add Delegated Permissions

### Step 1: Go to Azure Portal
https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps

### Step 2: Find Your App
- Application ID: `2cdb6dee-9c0d-45ba-a3eb-11aca2cc8ba8`
- Name: SharePoint Site Export Tool

### Step 3: Add Delegated Permissions
1. Click on your app
2. Go to **"API permissions"** (left menu)
3. Click **"+ Add a permission"**
4. Select **"SharePoint"**
5. Click **"Delegated permissions"** (NOT Application)
6. Check these permissions:
   - ☑️ **AllSites.FullControl** - Have full control of all site collections
   - ☑️ **AllSites.Read** - Read items in all site collections
   - ☑️ **AllSites.Write** - Edit items in all site collections

7. Click **"Add permissions"**
8. Click **"Grant admin consent for [Your Organization]"**
9. Confirm **"Yes"**

### Step 4: Verify Permissions
After adding, you should see BOTH:
- **Application permissions** (Sites.FullControl.All, Sites.Manage.All, etc.)
- **Delegated permissions** (AllSites.FullControl, AllSites.Read, AllSites.Write)

### Step 5: Test Again
Once delegated permissions are added and consented:

```powershell
.\NewPnPScript.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -AuthMode Interactive `
    -InteractiveClientId "2cdb6dee-9c0d-45ba-a3eb-11aca2cc8ba8" `
    -IncludeHidden `
    -OutputCsv ".\AllLists.csv"
```

This should now find all your lists including "Offboarding List"!

## Why This Matters

- **APPLICATION permissions** = App-only authentication (no user context)
  - Requires certificate-based auth
  - Used for automated/daemon processes
  
- **DELEGATED permissions** = User + App authentication
  - Works with interactive login
  - User signs in and acts on their behalf
  - This is what we need!

## Expected Result
After fixing permissions, you should see lists like:
- Offboarding List
- Documents (if any)
- Site Pages
- Any other custom lists

---

Let me know once you've added the delegated permissions and we'll test again!
