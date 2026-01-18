# Quick Start Guide - Offboarding Copilot Agent Deployment

## 🚀 5-Minute Quick Start

### Step 1: Export Your Current Site (2 minutes)

```powershell
cd "d:\OneDrive\OneDrive - Microsoft\Documents\Learning Projects\CustomScripts\OffboardingAgentSolution"

.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export" `
    -IncludeData
```

**What happens**: Creates a complete backup of your site in `.\Export\SharePointExport_<timestamp>\`

### Step 2: Deploy to New Site (3 minutes)

```powershell
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://YOUR_TENANT.sharepoint.com/sites/YOUR_NEW_SITE" `
    -ExportPath ".\Export\SharePointExport_<timestamp>" `
    -IncludeData
```

**What happens**: Recreates all lists, columns, and data in your new site

### Step 3: Update Your Copilot Agent

Update your agent's SharePoint connection to point to the new site URL.

---

## 📋 Prerequisites Checklist

- [ ] PowerShell 7+ installed
- [ ] SharePoint site URL ready
- [ ] Appropriate permissions (Read for export, Full Control for deploy)
- [ ] Network connectivity to SharePoint

---

## 🎯 Common Scenarios

### Scenario 1: Testing Environment
```powershell
# Export production
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export"

# Deploy to test (without production data)
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess-Test" `
    -ExportPath ".\Export\SharePointExport_<timestamp>"
```

### Scenario 2: Full Backup & Restore
```powershell
# Backup everything
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Backups\$(Get-Date -Format 'yyyy-MM-dd')" `
    -IncludeData

# Restore if needed
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess-Restored" `
    -ExportPath ".\Backups\2024-12-18\SharePointExport_<timestamp>" `
    -IncludeData
```

### Scenario 3: New Region Deployment
```powershell
# Export template only (no data)
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Templates"

# Deploy to multiple regions
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://region1.sharepoint.com/sites/Offboarding" `
    -ExportPath ".\Templates\SharePointExport_<timestamp>"
```

---

## 🔧 What You'll Get

After export, you'll have:
- ✅ Complete site structure
- ✅ All list definitions and schemas
- ✅ Custom columns and field configurations
- ✅ Custom views and queries
- ✅ List data (if -IncludeData used)
- ✅ JSON and CSV exports for easy viewing

---

## ⚡ Tips for Success

1. **Start Small**: Test export on a small site first
2. **Review Output**: Check the ExportSummary.txt file
3. **Verify Permissions**: Ensure you have access before deploying
4. **Test First**: Always deploy to a test site first
5. **Keep Backups**: Save export folders with timestamps

---

## 🆘 Quick Troubleshooting

| Problem | Solution |
|---------|----------|
| Module not found | Run: `Install-Module PnP.PowerShell -Scope CurrentUser` |
| Access denied | Check your SharePoint permissions |
| List already exists | Script will update it automatically |
| Connection fails | Try `-Interactive` authentication |

---

## 📁 File Structure After Running

```
OffboardingAgentSolution/
├── Export-SharePointSiteStructure.ps1  ← Export script
├── Deploy-SharePointSite.ps1           ← Deploy script
├── README.md                           ← Full documentation
├── QUICKSTART.md                       ← This file
└── Export/                             ← Your exports appear here
    └── SharePointExport_20241218_120000/
        ├── SiteInfo.json
        ├── DeploymentManifest.json
        ├── AllLists.json
        ├── ExportSummary.txt
        └── List_*.json files
```

---

## 🎓 Next Steps

1. ✅ Run your first export
2. ✅ Review the export files
3. ✅ Test deployment to a dev site
4. ✅ Update your Copilot Agent configuration
5. ✅ Document your new site URL

---

Need more details? Check [README.md](README.md) for complete documentation.
