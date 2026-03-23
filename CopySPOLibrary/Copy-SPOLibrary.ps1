<#
.SYNOPSIS
    Copies a SharePoint Online document library (structure, custom fields, and files) to another site.

.DESCRIPTION
    This script uses PnP PowerShell to:
    1. Recreate the library on the target site with all custom columns and content types
    2. Copy the folder structure
    3. Server-side copy files using Copy-PnPFile (no download/upload)
    4. Set custom metadata on copied files

    Authentication: Logs in only twice (once per site) using -ReturnConnection.

.PARAMETER SourceSiteUrl
    The URL of the source SharePoint site.

.PARAMETER TargetSiteUrl
    The URL of the target SharePoint site.

.PARAMETER SourceLibraryName
    The name of the source document library.

.PARAMETER TargetLibraryName
    The name of the target document library. Defaults to the source library name.

.PARAMETER ClientId
    The Entra app registration Client ID for PnP interactive auth.

.EXAMPLE
    .\Copy-SPOLibrary.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/source" `
                          -TargetSiteUrl "https://contoso.sharepoint.com/sites/target" `
                          -SourceLibraryName "Project Documents"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$TargetSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$SourceLibraryName,

    [Parameter(Mandatory = $false)]
    [string]$TargetLibraryName = $SourceLibraryName,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "YOUR-APP-CLIENT-ID"
)

#Requires -Modules PnP.PowerShell

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# Built-in fields to skip when copying columns
# ============================================================
$BuiltInFieldsToSkip = @(
    "ContentType", "Title", "_ModerationComments", "File_x0020_Type",
    "ContentTypeId", "_HasCopyDestinations", "_CopySource", "_ModerationStatus",
    "FileRef", "FileDirRef", "Last_x0020_Modified", "Created_x0020_Date",
    "FSObjType", "SortBehavior", "PermMask", "FileLeafRef", "UniqueId",
    "SyncClientId", "ProgId", "ScopeId", "HTML_x0020_File_x0020_Type",
    "_EditMenuTableStart", "_EditMenuTableStart2", "_EditMenuTableEnd",
    "LinkFilenameNoMenu", "LinkFilename", "LinkFilename2", "DocIcon",
    "ServerUrl", "EncodedAbsUrl", "BaseName", "MetaInfo", "Level",
    "_Level", "_IsCurrentVersion", "ItemChildCount", "FolderChildCount",
    "Restricted", "OriginatorId", "NoExecute", "ContentVersion",
    "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime",
    "_ComplianceTagUserId", "_IsRecord", "AccessPolicy", "_VirusStatus",
    "_VirusVendorID", "_VirusInfo", "_CommentFlags", "_CommentCount",
    "BSN", "_ListSchemaVersion", "_Dirty", "_Parsable",
    "ParentUniqueId", "StreamHash", "ParentLeafName",
    "Created", "Modified", "Author", "Editor",
    "_UIVersionString", "_UIVersion", "GUID",
    "WorkflowVersion", "Attachments", "Order", "AppAuthor", "AppEditor",
    "SMTotalSize", "SMLastModifiedDate", "SMTotalFileStreamSize",
    "SMTotalFileCount", "ComplianceAssetId", "CheckoutUser",
    "CheckedOutTitle", "CheckinComment", "LinkCheckedOutTitle",
    "_CheckinComment", "FileSizeDisplay", "SelectTitle",
    "Edit", "owshiddenversion", "_CopyFlags",
    "MediaServiceFastMetadata", "MediaServiceMetadata",
    "MediaServiceAutoTags", "MediaServiceOCR",
    "MediaServiceGenerationTime", "MediaServiceEventHashCode",
    "MediaServiceAutoKeyPoints", "MediaServiceKeyPoints",
    "MediaServiceDateTaken", "MediaServiceLocation",
    "MediaServiceObjectDetectorVersions",
    "MediaLengthInSeconds", "A2ODMountCount", "_IpLabelId",
    "_IpLabelAssignmentMethod", "_DisplayName",
    "_IpLabelHash", "_IsAltChunk"
)

# ============================================================
# Helper: Write progress with timestamp
# ============================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        "ERROR"   { Write-Host "[$timestamp] ERROR: $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "[$timestamp] WARNING: $Message" -ForegroundColor Yellow }
        "SUCCESS" { Write-Host "[$timestamp] SUCCESS: $Message" -ForegroundColor Green }
        default   { Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Cyan }
    }
}

# ============================================================
# Helper: Convert CSOM field values to plain values
# ============================================================
function Convert-FieldValue {
    param($Value)

    if ($null -eq $Value -or $Value -eq "") { return $null }

    if ($Value -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
        return $Value.LookupId
    }
    elseif ($Value -is [Microsoft.SharePoint.Client.FieldUserValue]) {
        return $Value.Email
    }
    elseif ($Value -is [Microsoft.SharePoint.Client.FieldUrlValue]) {
        return "$($Value.Url), $($Value.Description)"
    }
    elseif ($Value -is [array]) {
        $converted = $Value | ForEach-Object {
            if ($_ -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $_.LookupId }
            elseif ($_ -is [Microsoft.SharePoint.Client.FieldUserValue]) { $_.Email }
            else { $_ }
        }
        return $converted
    }
    else {
        return $Value
    }
}

# ============================================================
# STEP 0 — Establish both connections (login only twice)
# ============================================================
Write-Log "Connecting to source site: $SourceSiteUrl"
$sourceConn = Connect-PnPOnline -Url $SourceSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

Write-Log "Connecting to target site: $TargetSiteUrl"
$targetConn = Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

# ============================================================
# STEP 1 — Read source library metadata
# ============================================================
Write-Log "Reading source library: $SourceLibraryName"
$sourceList = Get-PnPList -Identity $SourceLibraryName -Includes ContentTypes, Fields, RootFolder -Connection $sourceConn

if (-not $sourceList) {
    Write-Log "Source library '$SourceLibraryName' not found." "ERROR"
    return
}

# Get custom fields (skip built-in and sealed fields)
$sourceFields = @(Get-PnPField -List $SourceLibraryName -Connection $sourceConn | Where-Object {
    -not $_.Hidden -and
    -not $_.ReadOnlyField -and
    -not $_.FromBaseType -and
    $_.InternalName -notin $BuiltInFieldsToSkip -and
    $_.SchemaXml -notmatch 'SourceID="http://schemas.microsoft.com/'
})

Write-Log "Found $($sourceFields.Count) custom field(s) to copy."

# Get all items (files and folders) from source
Write-Log "Retrieving all items from source library..."
$sourceItems = @(Get-PnPListItem -List $SourceLibraryName -PageSize 2000 -Connection $sourceConn)

$sourceFolders = @($sourceItems | Where-Object { $_.FileSystemObjectType -eq "Folder" })
$sourceFiles   = @($sourceItems | Where-Object { $_.FileSystemObjectType -eq "File" })

Write-Log "Found $($sourceFolders.Count) folder(s) and $($sourceFiles.Count) file(s)."

# Build field internal names list for metadata copy
$customFieldNames = @($sourceFields | ForEach-Object { $_.InternalName })

# Store field schemas for recreation
$fieldSchemas = @($sourceFields | ForEach-Object {
    [PSCustomObject]@{
        InternalName = $_.InternalName
        SchemaXml    = $_.SchemaXml
        TypeAsString = $_.TypeAsString
    }
})

# Store source content types (non-built-in)
$sourceContentTypes = @(Get-PnPContentType -List $SourceLibraryName -Connection $sourceConn | Where-Object {
    $_.Name -ne "Folder" -and $_.Name -ne "Document" -and -not $_.Hidden
})
Write-Log "Found $($sourceContentTypes.Count) custom content type(s)."

# ============================================================
# Collect file and folder data from source
# ============================================================
Write-Log "Collecting file metadata from source..."
$rootFolderUrl = $sourceList.RootFolder.ServerRelativeUrl

# Collect folder relative paths
$folderPaths = @()
foreach ($folder in $sourceFolders) {
    $folderServerRelUrl = $folder.FieldValues["FileRef"]
    $relativePath = $folderServerRelUrl.Substring($rootFolderUrl.Length).TrimStart("/")
    if ($relativePath -and $relativePath -ne "Forms") {
        $folderPaths += $relativePath
    }
}

$folderPaths = @($folderPaths | Sort-Object { ($_ -split "/").Count })

# Collect file server-relative URLs and metadata
$fileDataCollection = @()

foreach ($file in $sourceFiles) {
    $fileServerRelUrl = $file.FieldValues["FileRef"]
    $relativePath = $fileServerRelUrl.Substring($rootFolderUrl.Length).TrimStart("/")

    if ($relativePath -like "Forms/*") { continue }

    $metadata = @{}
    foreach ($fieldName in $customFieldNames) {
        $value = $file.FieldValues[$fieldName]
        $converted = Convert-FieldValue -Value $value
        if ($null -ne $converted -and $converted -ne "") {
            $metadata[$fieldName] = $converted
        }
    }

    $fileDataCollection += [PSCustomObject]@{
        RelativePath    = $relativePath
        SourceServerUrl = $fileServerRelUrl
        Metadata        = $metadata
    }

    Write-Log "Collected metadata: $relativePath"
}

Write-Log "Source data collected." "SUCCESS"

# Parse target site relative path from URL (e.g., /sites/Bosch)
$targetUri = [System.Uri]$TargetSiteUrl
$targetSiteRelativePath = $targetUri.AbsolutePath.TrimEnd("/")

# ============================================================
# STEP 2 — Create library on target
# ============================================================
$existingList = Get-PnPList -Identity $TargetLibraryName -Connection $targetConn -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Log "Library '$TargetLibraryName' already exists on target site. Proceeding to add missing fields and content." "WARNING"
} else {
    Write-Log "Creating library: $TargetLibraryName"
    New-PnPList -Title $TargetLibraryName -Template DocumentLibrary -Connection $targetConn | Out-Null
    Write-Log "Library created." "SUCCESS"
}

# ============================================================
# STEP 3 — Copy custom fields to target library
# ============================================================
Write-Log "Creating custom fields on target library..."

$createdFieldNames = @()

foreach ($field in $fieldSchemas) {
    $existingField = Get-PnPField -List $TargetLibraryName -Identity $field.InternalName -Connection $targetConn -ErrorAction SilentlyContinue
    if ($existingField) {
        Write-Log "Field '$($field.InternalName)' already exists — skipping." "WARNING"
        $createdFieldNames += $field.InternalName
        continue
    }

    try {
        $schema = $field.SchemaXml
        $schema = $schema -replace '\s+List="\{[^"]*\}"', ''
        $schema = $schema -replace '\s+SourceID="[^"]*"', ''
        $schema = $schema -replace '\s+Version="\d+"', ''
        $newGuid = [Guid]::NewGuid().ToString("B")
        $schema = $schema -replace 'ID="\{[^"]*\}"', "ID=""$newGuid"""

        Add-PnPFieldFromXml -List $TargetLibraryName -FieldXml $schema -Connection $targetConn | Out-Null
        $createdFieldNames += $field.InternalName
        Write-Log "Created field: $($field.InternalName)" "SUCCESS"
    }
    catch {
        Write-Log "Failed to create field '$($field.InternalName)': $_" "ERROR"
    }
}

# Add all custom fields to the default view
if ($createdFieldNames.Count -gt 0) {
    try {
        $defaultView = Get-PnPView -List $TargetLibraryName -Identity "All Documents" -Connection $targetConn
        $viewFields = @($defaultView.ViewFields) + $createdFieldNames
        Set-PnPView -List $TargetLibraryName -Identity "All Documents" -Fields $viewFields -Connection $targetConn | Out-Null
        Write-Log "Added $($createdFieldNames.Count) field(s) to default view." "SUCCESS"
    }
    catch {
        Write-Log "Failed to update default view: $_" "ERROR"
    }
}

# ============================================================
# STEP 4 — Copy content types (if any custom ones exist)
# ============================================================
if ($sourceContentTypes.Count -gt 0) {
    Write-Log "Copying custom content types..."

    foreach ($ct in $sourceContentTypes) {
        try {
            $existingCT = Get-PnPContentType -List $TargetLibraryName -Identity $ct.Name -Connection $targetConn -ErrorAction SilentlyContinue
            if ($existingCT) {
                Write-Log "Content type '$($ct.Name)' already exists — skipping." "WARNING"
                continue
            }

            $siteCT = Get-PnPContentType -Identity $ct.Name -Connection $targetConn -ErrorAction SilentlyContinue
            if ($siteCT) {
                Add-PnPContentTypeToList -List $TargetLibraryName -ContentType $ct.Name -Connection $targetConn | Out-Null
                Write-Log "Added content type '$($ct.Name)' to library." "SUCCESS"
            } else {
                Write-Log "Content type '$($ct.Name)' not found at site level on target. You may need to create it manually or deploy it via a content type hub." "WARNING"
            }
        }
        catch {
            Write-Log "Failed to add content type '$($ct.Name)': $_" "ERROR"
        }
    }
}

# ============================================================
# STEP 5 — Create folder structure on target
# ============================================================
Write-Log "Creating folder structure on target..."

foreach ($folderPath in $folderPaths) {
    try {
        Resolve-PnPFolder -SiteRelativePath "$TargetLibraryName/$folderPath" -Connection $targetConn | Out-Null
        Write-Log "Created folder: $folderPath" "SUCCESS"
    }
    catch {
        Write-Log "Failed to create folder '$folderPath': $_" "ERROR"
    }
}

# ============================================================
# STEP 6 — Server-side copy files using Copy-PnPFile
# ============================================================
Write-Log "Copying files server-side (no download/upload)..."

$successCount = 0
$errorCount = 0

foreach ($fileData in $fileDataCollection) {
    $sourceUrl = $fileData.SourceServerUrl
    $targetFolderRelPath = "$targetSiteRelativePath/$TargetLibraryName"
    $parentPath = Split-Path $fileData.RelativePath -Parent
    if ($parentPath) {
        $parentPath = $parentPath -replace "\\", "/"
        $targetFolderRelPath = "$targetFolderRelPath/$parentPath"
    }

    try {
        Copy-PnPFile -SourceUrl $sourceUrl -TargetUrl $targetFolderRelPath -Force -Connection $sourceConn -ErrorAction Stop
        Write-Log "Copied: $($fileData.RelativePath)" "SUCCESS"
        $successCount++
    }
    catch {
        Write-Log "Failed to copy '$($fileData.RelativePath)': $_" "ERROR"
        $errorCount++
    }
}

Write-Log "Server-side copy complete." "SUCCESS"

# ============================================================
# STEP 7 — Set metadata on copied files
# ============================================================
Write-Log "Setting metadata on copied files..."

$metadataSuccess = 0
$metadataError = 0

foreach ($fileData in $fileDataCollection) {
    if ($fileData.Metadata.Count -eq 0) {
        continue
    }

    $targetFileUrl = "$targetSiteRelativePath/$TargetLibraryName/$($fileData.RelativePath)"

    try {
        $targetFile = Get-PnPFile -Url $targetFileUrl -AsListItem -Connection $targetConn -ErrorAction Stop
        Set-PnPListItem -List $TargetLibraryName -Identity $targetFile.Id -Values $fileData.Metadata -Connection $targetConn | Out-Null
        Write-Log "Set metadata: $($fileData.RelativePath)" "SUCCESS"
        $metadataSuccess++
    }
    catch {
        Write-Log "Failed to set metadata for '$($fileData.RelativePath)': $_" "ERROR"
        $metadataError++
    }
}

# ============================================================
# Summary
# ============================================================
Disconnect-PnPOnline -Connection $sourceConn
Disconnect-PnPOnline -Connection $targetConn

Write-Log "============================================"
Write-Log "Copy complete!" "SUCCESS"
Write-Log "  Files copied (server-side): $successCount"
Write-Log "  Metadata updated: $metadataSuccess"
if ($errorCount -gt 0) {
    Write-Log "  File copy errors: $errorCount" "ERROR"
}
if ($metadataError -gt 0) {
    Write-Log "  Metadata errors: $metadataError" "ERROR"
}
Write-Log "  Custom fields copied: $($fieldSchemas.Count)"
Write-Log "  Folders created: $($folderPaths.Count)"
Write-Log "============================================"
