# App Registration Setup Guide for Device Import Script

This guide provides step-by-step instructions to create and configure an Azure AD App Registration for importing device identities to Intune using the Import-SingleDeviceIdentity.ps1 script.

## Overview

App Registration (Service Principal) authentication is ideal for:
- **Automated/scheduled tasks** - Scripts running unattended
- **Service accounts** - Background processes without user interaction
- **CI/CD pipelines** - Automated deployments
- **Secure credential storage** - Using Azure Key Vault or other secret management

---

## Step 1: Create the App Registration

### Using Azure Portal (Portal.Azure.com)

1. **Navigate to Azure AD**
   - Go to [Azure Portal](https://portal.azure.com)
   - Search for and select **"Azure Active Directory"** or **"Microsoft Entra ID"**

2. **Create New App Registration**
   - Click **"App registrations"** in the left menu
   - Click **"+ New registration"**
   
3. **Configure Basic Settings**
   - **Name**: `Intune-Device-Import-App` (or your preferred name)
   - **Supported account types**: 
     - Select **"Accounts in this organizational directory only (Single tenant)"**
   - **Redirect URI**: Leave blank (not needed for service principal)
   - Click **"Register"**

4. **Note Important Values**
   - After creation, you'll see the app overview page
   - **Copy and save these values**:
     - ✅ **Application (client) ID** - This is your `ClientId`
     - ✅ **Directory (tenant) ID** - This is your `TenantId`

---

## Step 2: Create Client Secret

1. **Navigate to Certificates & Secrets**
   - In your app registration, click **"Certificates & secrets"** in the left menu
   - Click on **"Client secrets"** tab

2. **Add New Client Secret**
   - Click **"+ New client secret"**
   - **Description**: `Device Import Script Secret` (or your preferred description)
   - **Expires**: Choose expiration period
     - ⚠️ Recommended: **180 days or 1 year** (you'll need to rotate it)
     - For production: Set calendar reminder to rotate before expiration
   - Click **"Add"**

3. **Copy the Secret Value**
   - ⚠️ **IMPORTANT**: Copy the **Value** immediately (it's shown only once)
   - This is your `ClientSecret`
   - Store it securely (e.g., Azure Key Vault, Password Manager)
   - ❌ Do NOT commit this to source control

---

## Step 3: Grant API Permissions

1. **Navigate to API Permissions**
   - In your app registration, click **"API permissions"** in the left menu

2. **Add Microsoft Graph Permission**
   - Click **"+ Add a permission"**
   - Select **"Microsoft Graph"**
   - Select **"Application permissions"** (NOT Delegated permissions)

3. **Select Required Permission**
   - Expand **"DeviceManagementServiceConfig"**
   - Check ✅ **"DeviceManagementServiceConfig.ReadWrite.All"**
     - This allows the app to import corporate device identifiers
   - Click **"Add permissions"**

4. **Grant Admin Consent** ⚠️ (CRITICAL STEP)
   - Click **"✓ Grant admin consent for [Your Organization]"**
   - Click **"Yes"** to confirm
   - You should see a green checkmark in the "Status" column
   - ⚠️ **This step requires Global Administrator or Privileged Role Administrator**

---

## Step 4: Verify Permissions

After granting admin consent, verify your setup:

### API Permissions Should Show:
| API / Permissions name | Type | Status |
|---|---|---|
| Microsoft Graph - DeviceManagementServiceConfig.ReadWrite.All | Application | ✓ Granted for [Your Org] |

---

## Step 5: Test the Configuration

### Option A: Test with Script

Create a test script to verify authentication:

```powershell
# Replace with your actual values
$TenantId = "YOUR-TENANT-ID"
$ClientId = "YOUR-CLIENT-ID"
$ClientSecret = "YOUR-CLIENT-SECRET"

# Test import
.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "123456789012345" `
    -ImportedDeviceIdentityType "imei" `
    -Description "Test device from app registration" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret
```

### Option B: Use Secure Credential Storage

```powershell
# Store credentials securely
$TenantId = "YOUR-TENANT-ID"
$ClientId = "YOUR-CLIENT-ID"
$ClientSecret = Get-Secret -Name "IntuneAppSecret" -Vault "MyVault"  # Using SecretManagement module

# Or read from environment variable
$ClientSecret = $env:INTUNE_CLIENT_SECRET

.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "123456789012345" `
    -ImportedDeviceIdentityType "imei" `
    -Description "Test device" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret
```

---

## Security Best Practices

### ✅ DO:
- **Store secrets securely** using Azure Key Vault, HashiCorp Vault, or Windows Credential Manager
- **Rotate client secrets regularly** (before expiration)
- **Use least privilege** - Only grant necessary permissions
- **Monitor app activity** using Azure AD sign-in logs
- **Set secret expiration reminders**
- **Document who has access** to the credentials

### ❌ DON'T:
- **Never commit secrets** to source control (use .gitignore)
- **Don't share secrets** via email or chat
- **Don't use overly broad permissions**
- **Don't forget to rotate** expired secrets
- **Don't hard-code secrets** in scripts

---

## Alternative: Using Certificate-Based Authentication

For enhanced security, consider certificate-based authentication instead of client secrets:

### Create Self-Signed Certificate:
```powershell
# Create certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=IntuneDeviceImport" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddMonths(12)

# Export certificate
$certPassword = ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "IntuneDeviceImport.pfx" -Password $certPassword
Export-Certificate -Cert $cert -FilePath "IntuneDeviceImport.cer"
```

Then upload the .cer file to **App Registration → Certificates & secrets → Certificates**.

---

## Troubleshooting

### Error: "Insufficient privileges to complete the operation"
- **Cause**: Admin consent not granted or insufficient permissions
- **Solution**: Ensure admin consent is granted (Step 3.4)

### Error: "AADSTS7000215: Invalid client secret is provided"
- **Cause**: Wrong client secret or expired secret
- **Solution**: Verify secret value or create new secret

### Error: "AADSTS700016: Application not found"
- **Cause**: Wrong Client ID or Tenant ID
- **Solution**: Double-check Application (client) ID and Directory (tenant) ID

### Error: "User is not authorized to perform this action"
- **Cause**: The app needs Application permission, not Delegated
- **Solution**: Verify you selected "Application permissions" in Step 3

---

## Automation Example: Azure Automation Runbook

To run this script in Azure Automation:

```powershell
# Store credentials in Azure Automation Variables (encrypted)
$TenantId = Get-AutomationVariable -Name 'IntuneTenantId'
$ClientId = Get-AutomationVariable -Name 'IntuneClientId'
$ClientSecret = Get-AutomationVariable -Name 'IntuneClientSecret'

# Import devices from CSV
$devices = Import-Csv "devices.csv"

foreach ($device in $devices) {
    .\Import-SingleDeviceIdentity.ps1 `
        -ImportedDeviceIdentifier $device.IMEI `
        -ImportedDeviceIdentityType "imei" `
        -Description $device.Description `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -ClientSecret $ClientSecret
}
```

---

## Quick Reference

### Required Information Summary

| Item | Where to Find | Parameter Name |
|------|---------------|----------------|
| Tenant ID | Azure AD → App Registration → Overview | `-TenantId` |
| Client ID | Azure AD → App Registration → Overview | `-ClientId` |
| Client Secret | Azure AD → App Registration → Certificates & secrets | `-ClientSecret` |

### Required Permission

| API | Permission | Type | Admin Consent |
|-----|------------|------|---------------|
| Microsoft Graph | DeviceManagementServiceConfig.ReadWrite.All | Application | Required ✅ |

---

## Additional Resources

- [Microsoft Graph API - importedDeviceIdentity](https://learn.microsoft.com/en-us/graph/api/resources/intune-enrollment-importeddeviceidentity?view=graph-rest-beta)
- [Register an application with Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Use Azure Key Vault secrets in pipeline](https://learn.microsoft.com/en-us/azure/devops/pipelines/release/azure-key-vault)

---

## Support

For issues or questions:
1. Verify all steps in this guide
2. Check Azure AD sign-in logs for authentication errors
3. Review script error messages
4. Consult Microsoft Graph API documentation

---

**Document Version**: 1.0  
**Last Updated**: December 17, 2025
