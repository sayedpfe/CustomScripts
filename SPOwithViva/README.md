# SharePoint Sites with Viva Engage Community Analysis

This PowerShell script analyzes all SharePoint sites in your tenant to identify which ones are connected to Viva Engage communities.

## Prerequisites

1. **PowerShell Modules Required:**
   ```powershell
   Install-Module Microsoft.Graph -Force
   Install-Module Microsoft.Graph.Beta -Force
   ```

2. **Required Permissions:**
   - `Sites.Read.All` - To read all SharePoint sites
   - `Group.Read.All` - To read Microsoft 365 Groups
   - `Sites.FullControl.All` - For admin access to all sites

## Usage

### Basic Usage (Default settings)
```powershell
.\SPOVivaSite.ps1
```

### Custom Output File
```powershell
.\SPOVivaSite.ps1 -OutputPath "C:\Reports\VivaEngageSites.csv"
```

### Skip CSV Export
```powershell
.\SPOVivaSite.ps1 -ExportToCsv:$false
```

### Custom Batch Size for Performance Tuning
```powershell
.\SPOVivaSite.ps1 -BatchSize 25
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `OutputPath` | String | `VivaEngageSites_[timestamp].csv` | Path for the CSV export file |
| `ExportToCsv` | Switch | `$true` | Whether to export results to CSV |
| `BatchSize` | Integer | `50` | Number of sites to process before adding a small delay |

## Output

The script provides:

1. **Console Output:**
   - Progress indicator during site analysis
   - Real-time discovery of Viva Engage connected sites
   - Summary statistics
   - List of all Viva Engage connected sites

2. **CSV Files:**
   - Complete results for all sites: `[OutputPath]`
   - Viva Engage sites only: `[OutputPath]_VivaEngageOnly.csv`

3. **PowerShell Object:**
   - Summary object with counts and analysis details
   - Collection of Viva Engage connected sites

## CSV Columns

| Column | Description |
|--------|-------------|
| `SiteName` | Display name of the SharePoint site |
| `SiteUrl` | Full URL of the SharePoint site |
| `SiteId` | Unique identifier for the site |
| `GroupId` | Microsoft 365 Group ID (if group-connected) |
| `GroupDisplayName` | Display name of the associated M365 Group |
| `ResourceProvisioningOptions` | Comma-separated list of provisioned resources |
| `IsVivaEngageCommunity` | Boolean indicating Viva Engage connection |
| `LastModified` | Last modified date of the site |
| `CreatedDateTime` | Creation date of the site |

## Example Output

```
VIVA ENGAGE COMMUNITY ANALYSIS SUMMARY
====================================================
Total SharePoint sites analyzed: 245
Sites connected to Viva Engage: 12
Sites NOT connected to Viva Engage: 233
====================================================

SharePoint Sites Connected to Viva Engage Communities:
-------------------------------------------------------
• Marketing Team Hub
  URL: https://contoso.sharepoint.com/sites/marketing
  Group: Marketing Team

• Product Development
  URL: https://contoso.sharepoint.com/sites/productdev
  Group: Product Development Team
```

## Notes

- The script uses Microsoft Graph API to access site and group information
- Large tenants may take several minutes to complete
- The script includes throttling protection to avoid API limits
- Results are automatically timestamped for easy tracking
- Only group-connected sites can have Viva Engage communities