<#
.SYNOPSIS
    SharePoint Online REST API Audit Log Testing Script
    
.DESCRIPTION
    This script tests and proves which SharePoint Online REST API endpoints are logged in Audit Logs.
    Based on research, endpoints that access file content (/$value) ARE logged, while metadata endpoints are NOT.
    
    Endpoints that SHOULD be logged (File Access activities):
    - GET /_api/web/GetFileByServerRelativeUrl('/...')/$value
    - GET /_api/web/GetFileById(guid')/$/value
    
    Endpoints that are NOT logged:
    - GET /_api/web/allproperties
    - GET /_api/web/lists/GetByTitle('Documents')
    - GET /_api/web/lists/.../items
    
.PARAMETER AuthMethod
    Authentication method: 'Interactive' or 'AppRegistration'
    
.PARAMETER TenantName
    Your tenant name (e.g., 'contoso' from contoso.sharepoint.com)
    
.PARAMETER SiteUrl
    Full SharePoint site URL (e.g., 'https://contoso.sharepoint.com/sites/TestSite')
    
.PARAMETER ClientId
    App Registration Client ID (required for AppRegistration auth)
    
.PARAMETER ClientSecret
    App Registration Client Secret (required for AppRegistration auth)
    
.PARAMETER TenantId
    Azure AD Tenant ID (required for AppRegistration auth)
    
.NOTES
    Author: SPO REST API Audit Testing
    Date: 2025-12-17
    Prerequisites:
    - For Interactive: PnP.PowerShell module or user credentials
    - For AppRegistration: App with Sites.Read.All or Sites.FullControl.All permissions
    - For audit verification: ExchangeOnlineManagement module with Compliance Admin role
    
.EXAMPLE
    # Interactive authentication
    .\FileAccess.ps1 -AuthMethod Interactive -TenantName "contoso" -SiteUrl "https://contoso.sharepoint.com/sites/TestSite"
    
.EXAMPLE
    # App Registration authentication
    .\FileAccess.ps1 -AuthMethod AppRegistration -TenantName "contoso" -SiteUrl "https://contoso.sharepoint.com/sites/TestSite" `
        -ClientId "your-client-id" -ClientSecret "your-client-secret" -TenantId "your-tenant-id"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Interactive', 'AppRegistration')]
    [string]$AuthMethod,
    
    [Parameter(Mandatory = $true)]
    [string]$TenantName,
    
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$TestFileName = "AuditTestFile.txt",
    
    [Parameter(Mandatory = $false)]
    [string]$TestLibrary = "Documents"
)

# -----------------------------
# Global Variables
# -----------------------------
$script:AccessToken = $null
$script:TestResults = @()
$script:TestStartTime = Get-Date

# -----------------------------
# Function: Write-Log
# -----------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info'    { 'Cyan' }
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

# -----------------------------
# Function: Get-AccessToken (App Registration)
# -----------------------------
function Get-AccessTokenAppRegistration {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )
    
    Write-Log "Acquiring access token using App Registration..." -Level Info
    
    try {
        $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $body = @{
            client_id     = $ClientId
            scope         = "$Resource/.default"
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }
        
        $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Log "Access token acquired successfully" -Level Success
        return $tokenResponse.access_token
    }
    catch {
        Write-Log "Failed to acquire access token: $_" -Level Error
        throw
    }
}

# -----------------------------
# Function: Get-AccessToken (Interactive)
# -----------------------------
function Get-AccessTokenInteractive {
    param(
        [string]$TenantName,
        [string]$SiteUrl,
        [string]$ClientId,
        [string]$TenantId = "b22f8675-8375-455b-941a-67bee4cf7747"
    )
    
    Write-Log "Setting up interactive authentication with MSAL.PS..." -Level Info
    
    # Check if MSAL.PS is installed
    if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
        Write-Log "MSAL.PS module not found. Installing..." -Level Warning
        Install-Module -Name MSAL.PS -Scope CurrentUser -Force -AllowClobber
    }
    
    Import-Module MSAL.PS -ErrorAction Stop
    
    try {
        $resource = "https://$TenantName.sharepoint.com"
        # Use standard MSAL redirect URI for public clients
        $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        # Use specific tenant ID for authentication
        $authority = "https://login.microsoftonline.com/$TenantId"
        
        Write-Log "Opening browser for interactive sign-in..." -Level Info
        Write-Host "`n========================================" -ForegroundColor Yellow
        Write-Host "INTERACTIVE SIGN-IN REQUIRED" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "A browser window will open for you to sign in." -ForegroundColor Cyan
        Write-Host "Please complete the sign-in process." -ForegroundColor Cyan
        Write-Host "NOTE: Make sure your app has this redirect URI configured:" -ForegroundColor Yellow
        Write-Host "  $redirectUri" -ForegroundColor Cyan
        Write-Host "========================================`n" -ForegroundColor Yellow
        
        # Get token interactively
        $tokenResponse = Get-MsalToken `
            -ClientId $ClientId `
            -Authority $authority `
            -Scopes "$resource/AllSites.Read", "$resource/AllSites.Write" `
            -Interactive `
            -RedirectUri $redirectUri `
            -ErrorAction Stop
        
        Write-Log "Authentication successful!" -Level Success
        Write-Log "Signed in as: $($tokenResponse.Account.Username)" -Level Success
        
        return $tokenResponse.AccessToken
    }
    catch {
        Write-Log "Failed to authenticate: $_" -Level Error
        throw
    }
}

# -----------------------------
# Function: Ensure-TestFile
# -----------------------------
function Ensure-TestFile {
    param(
        [string]$SiteUrl,
        [string]$Library,
        [string]$FileName
    )
    
    Write-Log "Ensuring test file exists: $FileName" -Level Info
    
    try {
        # Use REST API to check/create file (works for both auth methods)
        $listTitle = $Library
        $headers = @{
            "Authorization" = "Bearer $script:AccessToken"
            "Accept"        = "application/json;odata=verbose"
            "Content-Type"  = "application/json;odata=verbose"
        }
        
        # Get library - URL encode the title
        $encodedListTitle = [System.Web.HttpUtility]::UrlEncode($listTitle)
        $libraryEndpoint = "$SiteUrl/_api/web/lists/GetByTitle('$encodedListTitle')/RootFolder"
        try {
            $libraryInfo = Invoke-RestMethod -Uri $libraryEndpoint -Headers $headers -Method Get
            $serverRelativeUrl = $libraryInfo.d.ServerRelativeUrl
        }
        catch {
            Write-Log "Could not find library '$listTitle'. Error: $_" -Level Error
            throw
        }
        
        # Try to get existing file - need to handle URL encoding
        $fileUrl = "$serverRelativeUrl/$FileName"
        $encodedFileUrl = [System.Uri]::EscapeDataString($fileUrl)
        $fileEndpoint = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedFileUrl')"
        
        try {
            $existingFile = Invoke-RestMethod -Uri $fileEndpoint -Headers $headers -Method Get -ErrorAction Stop
            Write-Log "Test file already exists" -Level Info
            return $fileUrl
        }
        catch {
            # File doesn't exist, but we'll use the existing file provided by user
            Write-Log "File will be tested (assuming it exists)" -Level Info
            return $fileUrl
        }
    }
    catch {
        Write-Log "Error ensuring test file: $_" -Level Error
        throw
    }
}

# -----------------------------
# Function: Test-RestEndpoint
# -----------------------------
function Test-RestEndpoint {
    param(
        [string]$EndpointUrl,
        [string]$Description,
        [bool]$ExpectedToBeLogged,
        [string]$AcceptHeader = "application/json;odata=nometadata",
        [bool]$IsFileDownload = $false
    )
    
    Write-Log "Testing: $Description" -Level Info
    Write-Log "Endpoint: $EndpointUrl" -Level Info
    Write-Log "Expected in Audit Log: $ExpectedToBeLogged" -Level Info
    
    $headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Accept"        = $AcceptHeader
    }
    
    $testResult = [PSCustomObject]@{
        Timestamp          = Get-Date
        Endpoint           = $EndpointUrl
        Description        = $Description
        ExpectedInAuditLog = $ExpectedToBeLogged
        Success            = $false
        StatusCode         = $null
        Error              = $null
    }
    
    try {
        if ($IsFileDownload) {
            $outputPath = "$env:TEMP\$TestFileName"
            Invoke-RestMethod -Uri $EndpointUrl -Headers $headers -Method Get -OutFile $outputPath -ErrorAction Stop
            $testResult.Success = $true
            $testResult.StatusCode = 200
            Write-Log "✓ Request succeeded - File downloaded to: $outputPath" -Level Success
        }
        else {
            $response = Invoke-RestMethod -Uri $EndpointUrl -Headers $headers -Method Get -ErrorAction Stop
            $testResult.Success = $true
            $testResult.StatusCode = 200
            Write-Log "✓ Request succeeded - Response received" -Level Success
        }
    }
    catch {
        $testResult.Success = $false
        $testResult.Error = $_.Exception.Message
        $testResult.StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Log "✗ Request failed: $($_.Exception.Message)" -Level Error
    }
    
    $script:TestResults += $testResult
    Write-Log "---" -Level Info
    
    return $testResult
}

# -----------------------------
# Function: Test-FileUpdate
# -----------------------------
function Test-FileUpdate {
    param(
        [string]$FileServerRelativeUrl,
        [string]$Description,
        [bool]$ExpectedToBeLogged
    )
    
    Write-Log "Testing: $Description" -Level Info
    Write-Log "File: $FileServerRelativeUrl" -Level Info
    
    $encodedFileUrl = [System.Uri]::EscapeDataString($FileServerRelativeUrl)
    $endpointUrl = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedFileUrl')/ListItemAllFields"
    
    $headers = @{
        "Authorization"   = "Bearer $script:AccessToken"
        "Accept"          = "application/json;odata=nometadata"
        "Content-Type"    = "application/json;odata=nometadata"
        "IF-MATCH"        = "*"
        "X-HTTP-Method"   = "MERGE"
    }
    
    $updateBody = @{
        Title = "Audit Test - Updated at $(Get-Date -Format 'HH:mm:ss')"
    } | ConvertTo-Json
    
    $testResult = [PSCustomObject]@{
        Timestamp          = Get-Date
        Endpoint           = $endpointUrl
        Description        = $Description
        ExpectedInAuditLog = $ExpectedToBeLogged
        Success            = $false
        StatusCode         = $null
        Error              = $null
    }
    
    try {
        $response = Invoke-RestMethod -Uri $endpointUrl -Headers $headers -Method Post -Body $updateBody -ErrorAction Stop
        $testResult.Success = $true
        $testResult.StatusCode = 200
        Write-Log "✓ File properties updated successfully" -Level Success
    }
    catch {
        $testResult.Success = $false
        $testResult.Error = $_.Exception.Message
        $testResult.StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Log "✗ Update failed: $($_.Exception.Message)" -Level Error
    }
    
    $script:TestResults += $testResult
    Write-Log "---" -Level Info
    return $testResult
}

# -----------------------------
# Function: Test-FileRename
# -----------------------------
function Test-FileRename {
    param(
        [string]$FileServerRelativeUrl,
        [string]$NewFileName,
        [string]$Description,
        [bool]$ExpectedToBeLogged
    )
    
    Write-Log "Testing: $Description" -Level Info
    Write-Log "Original: $FileServerRelativeUrl" -Level Info
    
    # Calculate new path
    $folderPath = Split-Path $FileServerRelativeUrl -Parent
    $newFileUrl = "$folderPath/$NewFileName"
    
    $encodedFileUrl = [System.Uri]::EscapeDataString($FileServerRelativeUrl)
    $encodedNewFileUrl = [System.Uri]::EscapeDataString($newFileUrl)
    
    # Rename to new name
    $renameEndpoint = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedFileUrl')/moveto(newurl='$encodedNewFileUrl',flags=1)"
    
    $headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Accept"        = "application/json;odata=nometadata"
    }
    
    $testResult = [PSCustomObject]@{
        Timestamp          = Get-Date
        Endpoint           = $renameEndpoint
        Description        = $Description
        ExpectedInAuditLog = $ExpectedToBeLogged
        Success            = $false
        StatusCode         = $null
        Error              = $null
    }
    
    try {
        # Rename file
        Invoke-RestMethod -Uri $renameEndpoint -Headers $headers -Method Post -ErrorAction Stop
        Write-Log "✓ File renamed to: $NewFileName" -Level Success
        
        # Rename back to original
        Start-Sleep -Seconds 2
        $renameBackEndpoint = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedNewFileUrl')/moveto(newurl='$encodedFileUrl',flags=1)"
        Invoke-RestMethod -Uri $renameBackEndpoint -Headers $headers -Method Post -ErrorAction Stop
        Write-Log "✓ File renamed back to original name" -Level Success
        
        $testResult.Success = $true
        $testResult.StatusCode = 200
    }
    catch {
        $testResult.Success = $false
        $testResult.Error = $_.Exception.Message
        $testResult.StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Log "✗ Rename failed: $($_.Exception.Message)" -Level Error
    }
    
    $script:TestResults += $testResult
    Write-Log "---" -Level Info
    return $testResult
}

# -----------------------------
# Function: Test-FileUpload
# -----------------------------
function Test-FileUpload {
    param(
        [string]$FolderServerRelativeUrl,
        [string]$FileName,
        [string]$Description,
        [bool]$ExpectedToBeLogged
    )
    
    Write-Log "Testing: $Description" -Level Info
    Write-Log "Target folder: $FolderServerRelativeUrl" -Level Info
    Write-Log "File name: $FileName" -Level Info
    
    $encodedFolderUrl = [System.Uri]::EscapeDataString($FolderServerRelativeUrl)
    $uploadEndpoint = "$SiteUrl/_api/web/GetFolderByServerRelativeUrl('$encodedFolderUrl')/Files/add(url='$FileName',overwrite=true)"
    
    $fileContent = "Audit Test Content - Created at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $fileBytes = [System.Text.Encoding]::UTF8.GetBytes($fileContent)
    
    $headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Accept"        = "application/json;odata=nometadata"
        "Content-Type"  = "application/octet-stream"
    }
    
    $testResult = [PSCustomObject]@{
        Timestamp          = Get-Date
        Endpoint           = $uploadEndpoint
        Description        = $Description
        ExpectedInAuditLog = $ExpectedToBeLogged
        Success            = $false
        StatusCode         = $null
        Error              = $null
    }
    
    try {
        $response = Invoke-RestMethod -Uri $uploadEndpoint -Headers $headers -Method Post -Body $fileBytes -ErrorAction Stop
        $testResult.Success = $true
        $testResult.StatusCode = 200
        Write-Log "✓ File uploaded successfully: $FileName" -Level Success
        Write-Log "✓ Test file kept in library for verification" -Level Info
        
        # Clean up disabled - keeping file in library for verification
        # Uncomment below to enable automatic cleanup:
        # Start-Sleep -Seconds 2
        # $deleteEndpoint = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedFolderUrl/$FileName')"
        # $deleteHeaders = @{
        #     "Authorization" = "Bearer $script:AccessToken"
        #     "IF-MATCH"      = "*"
        #     "X-HTTP-Method" = "DELETE"
        # }
        # Invoke-RestMethod -Uri $deleteEndpoint -Headers $deleteHeaders -Method Post -ErrorAction Stop
        # Write-Log "✓ Test file cleaned up (deleted)" -Level Success
    }
    catch {
        $testResult.Success = $false
        $testResult.Error = $_.Exception.Message
        $testResult.StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Log "✗ Upload failed: $($_.Exception.Message)" -Level Error
    }
    
    $script:TestResults += $testResult
    Write-Log "---" -Level Info
    return $testResult
}

# -----------------------------
# Function: Show-TestSummary
# -----------------------------
function Show-TestSummary {
    Write-Log "`n========================================" -Level Info
    Write-Log "TEST SUMMARY" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "Test Start Time: $script:TestStartTime" -Level Info
    Write-Log "Test End Time: $(Get-Date)" -Level Info
    Write-Log "Total Tests: $($script:TestResults.Count)" -Level Info
    Write-Log "Successful: $($script:TestResults.Where({$_.Success}).Count)" -Level Success
    Write-Log "Failed: $($script:TestResults.Where({-not $_.Success}).Count)" -Level Error
    
    Write-Log "`n========================================" -Level Info
    Write-Log "DETAILED RESULTS" -Level Info
    Write-Log "========================================" -Level Info
    
    foreach ($result in $script:TestResults) {
        $status = if ($result.Success) { "✓ SUCCESS" } else { "✗ FAILED" }
        $level = if ($result.Success) { "Success" } else { "Error" }
        
        Write-Log "`n$status - $($result.Description)" -Level $level
        Write-Log "  Expected in Audit Log: $($result.ExpectedInAuditLog)" -Level Info
        Write-Log "  Status Code: $($result.StatusCode)" -Level Info
        if ($result.Error) {
            Write-Log "  Error: $($result.Error)" -Level Error
        }
    }
    
    Write-Log "`n========================================" -Level Info
    Write-Log "AUDIT LOG VERIFICATION" -Level Info
    Write-Log "========================================" -Level Warning
    Write-Log "IMPORTANT: Audit logs have a delay of 30-60+ minutes" -Level Warning
    Write-Log "Please verify in Purview Audit Log after waiting:" -Level Info
    Write-Log "1. Go to https://compliance.microsoft.com/auditlogsearch" -Level Info
    Write-Log "2. Filter by:" -Level Info
    Write-Log "   - Activities: FileAccessed, FileDownloaded" -Level Info
    Write-Log "   - Date range: From $script:TestStartTime to $(Get-Date)" -Level Info
    Write-Log "3. Look for requests matching your test file: $TestFileName" -Level Info
    Write-Log "`nExpected Results:" -Level Info
    Write-Log "✓ Should see: File download requests (/$Value endpoints)" -Level Success
    Write-Log "✗ Should NOT see: Metadata/list requests (allproperties, lists, items)" -Level Warning
    
    Write-Log "`n========================================" -Level Info
    Write-Log "POWERSHELL AUDIT VERIFICATION COMMAND" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "After 60+ minutes, run this command to verify:" -Level Info
    
    $verificationScript = @"
# Install and connect to Exchange Online (if not already)
# Install-Module ExchangeOnlineManagement -Scope CurrentUser
# Connect-ExchangeOnline

`$start = [DateTime]::Parse('$($script:TestStartTime)')
`$end = Get-Date
`$records = Search-UnifiedAuditLog -StartDate `$start -EndDate `$end -Operations "FileAccessed","FileDownloaded" -ResultSize 5000
`$records | Where-Object {`$_.ObjectId -like "*$TestFileName*"} | Select-Object UserId, Operations, ObjectId, CreationDate | Format-Table -AutoSize
"@
    
    Write-Host "`n$verificationScript" -ForegroundColor Gray
}

# -----------------------------
# MAIN EXECUTION
# -----------------------------
try {
    Write-Log "`n========================================" -Level Info
    Write-Log "SharePoint REST API Audit Log Testing" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "Authentication Method: $AuthMethod" -Level Info
    Write-Log "Site URL: $SiteUrl" -Level Info
    Write-Log "Test Library: $TestLibrary" -Level Info
    Write-Log "Test File: $TestFileName" -Level Info
    Write-Log "========================================`n" -Level Info
    
    # Validate parameters
    if ($AuthMethod -eq 'AppRegistration') {
        if (-not $ClientId -or -not $ClientSecret -or -not $TenantId) {
            throw "For AppRegistration auth, ClientId, ClientSecret, and TenantId are required"
        }
    }
    
    # Authenticate
    if ($AuthMethod -eq 'Interactive') {
        # Use the same ClientId provided or the app registration created
        if (-not $ClientId) {
            throw "ClientId is required for Interactive authentication. Please provide the Client ID of your app registration."
        }
        $script:AccessToken = Get-AccessTokenInteractive -TenantName $TenantName -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId
    }
    else {
        $resource = "https://$TenantName.sharepoint.com"
        $script:AccessToken = Get-AccessTokenAppRegistration -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $resource
    }
    
    # Build file server relative URL directly
    $siteRelativePath = ([System.Uri]$SiteUrl).AbsolutePath
    $fileServerRelativeUrl = "$siteRelativePath/$TestLibrary/$TestFileName"
    Write-Log "Test file path: $fileServerRelativeUrl`n" -Level Info
    
    # Get file GUID for testing
    $headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Accept"        = "application/json;odata=nometadata"
    }
    
    # URL encode the file path properly
    $encodedFileUrl = [System.Uri]::EscapeDataString($fileServerRelativeUrl)
    $fileInfoEndpoint = "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$encodedFileUrl')?`$select=UniqueId"
    
    Write-Log "Attempting to get file info from: $fileInfoEndpoint" -Level Info
    
    try {
        $fileInfo = Invoke-RestMethod -Uri $fileInfoEndpoint -Headers $headers -Method Get
        $fileGuid = $fileInfo.UniqueId
        Write-Log "File GUID: $fileGuid`n" -Level Success
    }
    catch {
        Write-Log "Failed to get file info. Error: $($_.Exception.Message)" -Level Error
        Write-Log "Response: $($_.ErrorDetails.Message)" -Level Error
        Write-Log "This might be a permissions issue. Make sure your app has Sites.Read.All or Sites.FullControl.All APPLICATION permissions with admin consent." -Level Warning
        throw
    }
    
    # Start Testing
    Write-Log "`n========================================" -Level Info
    Write-Log "TESTING ENDPOINTS" -Level Info
    Write-Log "========================================`n" -Level Info
    
    # TEST 1: File download by server relative URL (SHOULD BE LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$fileServerRelativeUrl')/`$value" `
        -Description "File Download - GetFileByServerRelativeUrl with /`$value" `
        -ExpectedToBeLogged $true `
        -AcceptHeader "application/octet-stream" `
        -IsFileDownload $true
    
    # TEST 2: File download by GUID (SHOULD BE LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/GetFileById('$fileGuid')/`$value" `
        -Description "File Download - GetFileById with /`$value" `
        -ExpectedToBeLogged $true `
        -AcceptHeader "application/octet-stream" `
        -IsFileDownload $true
    
    # TEST 3: Web allproperties (NOT LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/allproperties" `
        -Description "Web Metadata - allproperties" `
        -ExpectedToBeLogged $false
    
    # TEST 4: List metadata (NOT LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/lists/GetByTitle('$TestLibrary')" `
        -Description "List Metadata - GetByTitle" `
        -ExpectedToBeLogged $false
    
    # TEST 5: List items (NOT LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/lists/GetByTitle('$TestLibrary')/items?`$top=5" `
        -Description "List Items - Get items" `
        -ExpectedToBeLogged $false
    
    # TEST 6: File metadata without $value (NOT LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web/GetFileByServerRelativeUrl('$fileServerRelativeUrl')" `
        -Description "File Metadata - GetFileByServerRelativeUrl WITHOUT /`$value" `
        -ExpectedToBeLogged $false
    
    # TEST 7: Web properties (NOT LOGGED)
    Test-RestEndpoint `
        -EndpointUrl "$SiteUrl/_api/web?`$select=Title,Url,Created" `
        -Description "Web Properties - Basic properties" `
        -ExpectedToBeLogged $false
    
    # TEST 8: Update file properties (SHOULD BE LOGGED as FileModified)
    Write-Log "`n--- TESTING WRITE OPERATIONS (SHOULD BE LOGGED) ---`n" -Level Warning
    Test-FileUpdate `
        -FileServerRelativeUrl $fileServerRelativeUrl `
        -Description "File Modification - Update file properties" `
        -ExpectedToBeLogged $true
    
    # TEST 9: Rename file (SHOULD BE LOGGED as FileRenamed)
    Test-FileRename `
        -FileServerRelativeUrl $fileServerRelativeUrl `
        -NewFileName "${TestFileName}.renamed" `
        -Description "File Rename - Rename file back and forth" `
        -ExpectedToBeLogged $true
    
    # TEST 10: Upload new file (SHOULD BE LOGGED as FileUploaded)
    Test-FileUpload `
        -FolderServerRelativeUrl "$siteRelativePath/$TestLibrary" `
        -FileName "AuditTest_Upload_$(Get-Date -Format 'HHmmss').txt" `
        -Description "File Upload - Create new file" `
        -ExpectedToBeLogged $true
    
    # Show summary
    Show-TestSummary
    
    # Export results to CSV
    $csvPath = Join-Path $PSScriptRoot "AuditTest_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $script:TestResults | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Log "`nResults exported to: $csvPath" -Level Success
    
}
catch {
    Write-Log "Script execution failed: $_" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
}
finally {
    Write-Log "`nScript completed" -Level Info
}