# Offboarding Copilot Agent - SharePoint Deployment Solution

This solution provides a complete set of scripts to export and deploy the SharePoint site structure and lists that support your Custom Copilot Agent for offboarding processes.

## 📋 Overview

The solution consists of two main PowerShell scripts that enable you to:
1. **Export** the complete SharePoint site structure including lists, columns, views, and data
2. **Deploy** the structure to a new SharePoint site for replication or disaster recovery

## 🎯 Components

### Scripts

| Script | Purpose |
|--------|---------|
| `Export-SharePointSiteStructure.ps1` | Exports the complete site structure and lists from your SharePoint site |
| `Deploy-SharePointSite.ps1` | Deploys the exported structure to a new or existing SharePoint site |

### Site Information

- **Production Site**: `https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess`
- **Purpose**: Backend database for Custom Copilot Agent

## 🚀 Quick Start

### Prerequisites

1. **PowerShell 7+** (recommended)
2. **PnP.PowerShell Module** (will be auto-installed if missing)
3. **SharePoint Permissions**:
   - Source site: Read access (for export)
   - Target site: Full Control or Site Owner (for deployment)

### Installation

```powershell
# Clone or navigate to the solution folder
cd "d:\OneDrive\OneDrive - Microsoft\Documents\Learning Projects\CustomScripts\OffboardingAgentSolution"

# The PnP.PowerShell module will be automatically installed when you run the scripts
```

## 📤 Exporting Your Site

### Basic Export

Export site structure without data:

```powershell
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export"
```

### Export with Data

Export site structure including all list items:

```powershell
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Export" `
    -IncludeData
```

### Export Output

The export creates a timestamped folder with the following files:

```
Export/
└── SharePointExport_20241218_120000/
    ├── SiteInfo.json                      # Site metadata
    ├── DeploymentManifest.json            # Deployment configuration
    ├── AllLists.json                      # Summary of all lists
    ├── ExportSummary.txt                  # Human-readable summary
    ├── List_<ListName>_Schema.json        # Individual list schemas
    ├── List_<ListName>_Data.json          # List data (if -IncludeData used)
    └── List_<ListName>_Data.csv           # CSV format (if -IncludeData used)
```

## 📥 Deploying to a New Site

### Basic Deployment

Deploy structure without data:

```powershell
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewOffboarding" `
    -ExportPath ".\Export\SharePointExport_20241218_120000"
```

### Deployment with Data

Deploy structure and import all data:

```powershell
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewOffboarding" `
    -ExportPath ".\Export\SharePointExport_20241218_120000" `
    -IncludeData
```

## 🔍 What Gets Exported/Deployed

### Site Information
- Site title and description
- Site URL and relative path
- Creation date and language settings

### List Configuration
- List name, description, and template type
- Versioning settings (major/minor versions)
- Content approval settings
- Attachment settings
- Folder creation settings

### List Schema
- Custom columns/fields
  - Field types (Text, Choice, Number, Date, etc.)
  - Required/optional status
  - Default values
  - Choice options
- Content types
- Custom views with queries and columns

### List Data (Optional)
- All list items
- Field values (excluding system fields)
- Exported in both JSON and CSV formats

## 📝 Common Use Cases

### 1. Environment Replication

Create a copy of your production environment in a test site:

```powershell
# Step 1: Export from production
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Backups" `
    -IncludeData

# Step 2: Deploy to test environment
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess-Test" `
    -ExportPath ".\Backups\SharePointExport_20241218_120000" `
    -IncludeData
```

### 2. Backup and Recovery

Regular backup of your site structure:

```powershell
# Schedule this to run weekly
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Backups\$(Get-Date -Format 'yyyy-MM')" `
    -IncludeData
```

### 3. Cross-Tenant Migration

Move your site to a different Microsoft 365 tenant:

```powershell
# Export from source tenant
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://sourcetenant.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Migration" `
    -IncludeData

# Deploy to target tenant (use different credentials)
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Migration\SharePointExport_20241218_120000" `
    -IncludeData
```

### 4. Template Creation

Create a template without production data:

```powershell
# Export only structure
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "https://m365cpi90282478.sharepoint.com/sites/OffboardingProcess" `
    -ExportPath ".\Templates"

# Deploy as new instances
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/OffboardingProcess-Region1" `
    -ExportPath ".\Templates\SharePointExport_20241218_120000"
```

## 🔐 Authentication

Both scripts support multiple authentication methods:

### Interactive Authentication (Recommended)

```powershell
# Scripts will prompt for credentials
.\Export-SharePointSiteStructure.ps1 -SiteUrl "<url>" -ExportPath ".\Export"
```

### Credential Parameter

```powershell
# Store credentials
$cred = Get-Credential

# Use with export
.\Export-SharePointSiteStructure.ps1 `
    -SiteUrl "<url>" `
    -ExportPath ".\Export" `
    -Credential $cred

# Use with deployment
.\Deploy-SharePointSite.ps1 `
    -TargetSiteUrl "<url>" `
    -ExportPath ".\Export\SharePointExport_20241218_120000" `
    -Credential $cred
```

## ⚙️ Advanced Options

### Export Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `SiteUrl` | String | Yes | URL of the SharePoint site to export |
| `ExportPath` | String | No | Path where export files will be saved (default: `.\Export`) |
| `IncludeData` | Switch | No | Include list item data in export |
| `Credential` | PSCredential | No | Credentials for authentication |

### Deploy Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `TargetSiteUrl` | String | Yes | URL of the target SharePoint site |
| `ExportPath` | String | Yes | Path containing exported configuration files |
| `CreateSite` | Switch | No | Create a new site (requires admin permissions) |
| `IncludeData` | Switch | No | Import list item data from export |
| `Credential` | PSCredential | No | Credentials for authentication |

## 🛠️ Troubleshooting

### Common Issues

**Issue**: "PnP.PowerShell module not found"
```powershell
# Solution: Install manually
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
```

**Issue**: "Access denied" error
```powershell
# Solution: Verify permissions
# For export: Need at least Read access
# For deployment: Need Full Control or Site Owner rights
```

**Issue**: "List already exists"
```powershell
# Solution: The deployment script will update existing lists
# No action needed - script handles this automatically
```

**Issue**: Field type not supported
```powershell
# Solution: Some complex field types may need manual creation
# Check the deployment log for specific fields that failed
```

### Excluded Lists

The export automatically excludes system lists:
- Hidden lists
- Picture libraries
- App catalogs (appdata, appfiles)
- System galleries (Master Page, Solution, Theme, Web Part)

## 📊 Monitoring and Logs

### Export Logs

- Console output shows real-time progress
- `ExportSummary.txt` provides a readable summary
- `DeploymentManifest.json` contains structured metadata

### Deployment Logs

- Console output shows detailed progress
- `DeploymentLog_<timestamp>.txt` records all actions
- Failed operations are logged with reasons

## 🔄 Integration with Copilot Agent

After deployment, update your Copilot Agent configuration:

1. **Update Connection Strings**: Point to the new SharePoint site
2. **Verify Permissions**: Ensure the agent has access to the new site
3. **Test Functionality**: Verify all queries work with the new lists
4. **Update Documentation**: Record the new site URL

## 📅 Best Practices

1. **Regular Backups**: Schedule weekly exports with data
2. **Version Control**: Keep export folders with timestamps
3. **Test Deployments**: Always test in a non-production environment first
4. **Document Changes**: Track any manual modifications to lists
5. **Permission Reviews**: Regularly audit site permissions

## 🔗 Related Resources

- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [SharePoint REST API](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Custom Copilot Agents](https://docs.microsoft.com/en-us/microsoft-365-copilot/extensibility/)

## 📞 Support

For issues or questions:
1. Check the troubleshooting section
2. Review the deployment logs
3. Verify your permissions
4. Check PnP PowerShell documentation

## 📄 License

Internal use only - Microsoft Employee Project

---

**Version**: 1.0.0  
**Last Updated**: December 18, 2024  
**Maintainer**: Your Organization
