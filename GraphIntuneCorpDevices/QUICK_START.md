# Quick Start Guide

## Prerequisites

1. **Install Microsoft Graph PowerShell Module**
   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
   ```

2. **Choose Authentication Method**:
   - **Interactive**: Quick testing, manual runs
   - **App Registration**: Automation, scheduled tasks, CI/CD

---

## Option 1: Interactive Authentication (Quick Test)

Perfect for testing and one-off imports:

```powershell
.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "123456789012345" `
    -ImportedDeviceIdentityType "imei" `
    -Description "Test device" `
    -UseInteractive
```

**Required Permission**: 
- You need: `Intune Administrator` or `Global Administrator` role
- The script will prompt you to sign in and consent to permissions

---

## Option 2: App Registration (Automation)

For automated/scheduled tasks:

### Step 1: Create App Registration
See detailed steps in [APP_REGISTRATION_SETUP.md](./APP_REGISTRATION_SETUP.md)

**Quick Summary**:
1. Azure Portal → Azure AD → App registrations → New registration
2. Name: `Intune-Device-Import-App`
3. Copy: Application (client) ID and Directory (tenant) ID
4. Create client secret and copy its value
5. Add API Permission: `DeviceManagementManagedDevices.ReadWrite.All` (Application)
6. **Grant admin consent** ✅

### Step 2: Run Script with App Registration

```powershell
# Set your values
$TenantId = "YOUR-TENANT-ID"
$ClientId = "YOUR-CLIENT-ID"
$ClientSecret = "YOUR-CLIENT-SECRET"

# Import device
.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "123456789012345" `
    -ImportedDeviceIdentityType "imei" `
    -Description "Corporate device" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret
```

---

## Examples

### Import IMEI Device
```powershell
.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "353456789012345" `
    -ImportedDeviceIdentityType "imei" `
    -Description "iPhone 15 Pro" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret
```

### Import Serial Number
```powershell
.\Import-SingleDeviceIdentity.ps1 `
    -ImportedDeviceIdentifier "C02ABC123DEF" `
    -ImportedDeviceIdentityType "serialNumber" `
    -Description "MacBook Pro" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret
```

### Batch Import from CSV
```powershell
# Create CSV with columns: Identifier, Type, Description
$devices = Import-Csv "devices.csv"

foreach ($device in $devices) {
    .\Import-SingleDeviceIdentity.ps1 `
        -ImportedDeviceIdentifier $device.Identifier `
        -ImportedDeviceIdentityType $device.Type `
        -Description $device.Description `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -ClientSecret $ClientSecret
    
    Start-Sleep -Seconds 2  # Rate limiting
}
```

---

## Troubleshooting

### "Module not found"
```powershell
Install-Module Microsoft.Graph.Authentication -Force
```

### "Insufficient privileges"
- For Interactive: Ensure you have Intune Administrator role
- For App Registration: Verify admin consent was granted

### "Invalid client secret"
- Secret may be expired - create new one in Azure Portal
- Verify you copied the correct secret value (not the Secret ID)

---

## Next Steps

- **Secure Secrets**: Use Azure Key Vault for production
- **Automate**: Schedule with Azure Automation or Task Scheduler
- **Monitor**: Check Intune portal for imported devices
- **Scale**: Process CSV files for bulk imports

For detailed setup instructions, see [APP_REGISTRATION_SETUP.md](./APP_REGISTRATION_SETUP.md)
