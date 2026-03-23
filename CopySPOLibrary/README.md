# Copy-SPOLibrary.ps1

A PowerShell script that copies a SharePoint Online document library from one site to another, including custom columns, folder structure, files, and metadata.

## What It Copies

| Item | Details |
|------|---------|
| **Custom columns** | All custom fields (Text, Choice, Number, Date, Yes/No, Person, Hyperlink, etc.) with their full schema |
| **Content types** | Any custom content types associated with the library |
| **Folder structure** | Complete folder hierarchy, preserving nested folders |
| **Files** | All files via server-side copy (fast, no local download required) |
| **Metadata** | All custom field values on each file |
| **Default view** | Custom columns are added to the "All Documents" view |

## Prerequisites

### 1. PowerShell 7+

Check your version:

```powershell
$PSVersionTable.PSVersion
```

If you need to install it: [Install PowerShell](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)

### 2. PnP PowerShell Module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 3. Entra ID (Azure AD) App Registration

The script uses interactive authentication with a registered app. You need an **Entra ID App Registration** with the following **delegated** API permissions:

| API | Permission | Type |
|-----|-----------|------|
| Microsoft Graph | `Sites.ReadWrite.All` | Delegated |
| SharePoint | `AllSites.FullControl` | Delegated |

To register the app:

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com) > **App registrations** > **New registration**
2. Name: `PnP-CopyLibrary` (or any name)
3. Redirect URI: Select **Public client/native** and add `http://localhost`
4. Under **API permissions**, add the permissions listed above
5. Click **Grant admin consent**
6. Copy the **Application (client) ID** - you will need it as the `-ClientId` parameter

### 4. Permissions

The user running the script must have:

- **Read access** (Member/Owner) on the source site
- **Full Control** or **Owner** on the target site

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `SourceSiteUrl` | Yes | - | Full URL of the source SharePoint site |
| `TargetSiteUrl` | Yes | - | Full URL of the target SharePoint site |
| `SourceLibraryName` | Yes | - | Display name of the source document library |
| `TargetLibraryName` | No | Same as source | Display name for the target library (created if it doesn't exist) |
| `ClientId` | No | `YOUR-APP-CLIENT-ID` | Entra app registration Client ID |

## Usage

### Basic - Same library name on target

```powershell
.\Copy-SPOLibrary.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectA" `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/ProjectB" `
    -SourceLibraryName "Project Documents"
```

This creates a library called **"Project Documents"** on the target site with all content.

### Different library name on target

```powershell
.\Copy-SPOLibrary.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectA" `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/ProjectB" `
    -SourceLibraryName "Project Documents" `
    -TargetLibraryName "Archived Documents"
```

### Custom Client ID

```powershell
.\Copy-SPOLibrary.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectA" `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/ProjectB" `
    -SourceLibraryName "Project Documents" `
    -ClientId "your-app-client-id-here"
```

## How It Works

The script runs in 7 steps:

```
Step 0  Authenticate to both sites (2 login prompts)
Step 1  Read source library schema (fields, content types)
Step 2  Create target library
Step 3  Copy custom columns to target library + add to default view
Step 4  Copy content types (if any)
Step 5  Create folder structure on target
Step 6  Server-side copy files (Copy-PnPFile - no download/upload)
Step 7  Set custom metadata on each copied file
```

Files are copied **server-side** within SharePoint, meaning they are NOT downloaded to your machine and re-uploaded. This makes the copy fast and efficient, especially for large files.

## Authentication

The script prompts for login **only twice**:

1. Once for the **source** site
2. Once for the **target** site

Both connections are kept open and reused throughout the entire process.

## Sample Output

```
[2026-03-23 20:15:00] INFO: Connecting to source site: https://contoso.sharepoint.com/sites/ProjectA
[2026-03-23 20:15:05] INFO: Connecting to target site: https://contoso.sharepoint.com/sites/ProjectB
[2026-03-23 20:15:10] INFO: Reading source library: Project Documents
[2026-03-23 20:15:12] INFO: Found 8 custom field(s) to copy.
[2026-03-23 20:15:14] INFO: Found 8 folder(s) and 25 file(s).
[2026-03-23 20:15:15] INFO: Found 0 custom content type(s).
[2026-03-23 20:15:16] SUCCESS: Source data collected.
[2026-03-23 20:15:18] SUCCESS: Library created.
[2026-03-23 20:15:20] SUCCESS: Created field: ProjectName
[2026-03-23 20:15:22] SUCCESS: Created field: Department
...
[2026-03-23 20:15:40] SUCCESS: Created folder: Finance/Reports
...
[2026-03-23 20:15:55] SUCCESS: Copied: Finance/Reports/Q1_Report.xlsx
...
[2026-03-23 20:16:10] SUCCESS: Set metadata: Finance/Reports/Q1_Report.xlsx
...
[2026-03-23 20:16:30] ============================================
[2026-03-23 20:16:30] SUCCESS: Copy complete!
[2026-03-23 20:16:30]   Files copied (server-side): 25
[2026-03-23 20:16:30]   Metadata updated: 25
[2026-03-23 20:16:30]   Custom fields copied: 8
[2026-03-23 20:16:30]   Folders created: 8
[2026-03-23 20:16:30] ============================================
```

## Important Notes

### Lookup Columns
If the source library has **lookup columns** pointing to other lists, those lists must already exist on the target site before running the script.

### Managed Metadata
If the source library uses **managed metadata (taxonomy) columns**, the target site must have access to the same term store and term sets.

### Permissions
The script copies **structure, files, and metadata only**. It does not copy item-level or library-level permissions.

### Re-running the Script
The script is safe to re-run. It will:
- Skip fields that already exist on the target
- Skip content types that already exist
- Overwrite files that already exist (using `-Force`)

### Supported Field Types

| Field Type | Supported |
|-----------|-----------|
| Single line of text | Yes |
| Multiple lines of text | Yes |
| Choice (dropdown/radio) | Yes |
| Number | Yes |
| Currency | Yes |
| Date and Time | Yes |
| Yes/No (Boolean) | Yes |
| Person or Group | Yes |
| Hyperlink/Picture | Yes |
| Lookup | Yes* |
| Managed Metadata | Yes* |
| Calculated | Yes |

\* Requires the referenced list or term set to exist on the target site.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `Connect-PnPOnline: AADSTS65001` | Admin consent not granted for the app. Go to Entra > App registrations > API permissions > Grant admin consent |
| `Access denied` | Ensure you have Owner/Full Control on the target site |
| `Field already exists` | Safe to ignore - the script skips existing fields |
| `Content type not found at site level` | The custom content type needs to be published to the target site via the Content Type Hub |
| `Copy-PnPFile failed` | Ensure both sites are in the same Microsoft 365 tenant |
