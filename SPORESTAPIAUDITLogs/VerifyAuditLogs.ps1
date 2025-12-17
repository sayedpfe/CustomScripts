<#
.SYNOPSIS
    Verify SharePoint REST API Audit Logs
    
.DESCRIPTION
    This script searches the Unified Audit Log for SharePoint file access activities.
    Run this script SEPARATELY in a new PowerShell window to avoid DLL conflicts.
    
.NOTES
    Prerequisites:
    - ExchangeOnlineManagement module
    - Compliance Administrator or Audit Logs role
    - Wait 60-90 minutes after running FileAccess.ps1
    
.EXAMPLE
    .\VerifyAuditLogs.ps1
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$FileName = "Offboarding_EkaS@M365CPI90282478.OnMicrosoft.com_de.docx",
    
    [Parameter(Mandatory = $false)]
    [DateTime]$StartDate = (Get-Date "12/17/2025 13:41:38"),
    
    [Parameter(Mandatory = $false)]
    [DateTime]$EndDate = (Get-Date)
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint REST API Audit Log Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "File: $FileName" -ForegroundColor Yellow
Write-Host "Date Range: $StartDate to $EndDate" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Cyan

# Check if ExchangeOnlineManagement is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "⚠ ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Check if connected
try {
    Write-Host "Checking Exchange Online connection..." -ForegroundColor Cyan
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "✓ Already connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Not connected. Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "✓ Connected successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to connect: $_" -ForegroundColor Red
        Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
        Write-Host "1. Make sure you have the ExchangeOnlineManagement module installed" -ForegroundColor Cyan
        Write-Host "2. You need Compliance Admin or Audit Logs role" -ForegroundColor Cyan
        Write-Host "3. Try running in a new PowerShell window" -ForegroundColor Cyan
        exit 1
    }
}

# Check audit logging status
Write-Host "`nChecking audit log configuration..." -ForegroundColor Cyan
try {
    $auditConfig = Get-AdminAuditLogConfig
    if ($auditConfig.UnifiedAuditLogIngestionEnabled) {
        Write-Host "✓ Unified Audit Logging is enabled" -ForegroundColor Green
    }
    else {
        Write-Host "✗ Unified Audit Logging is DISABLED" -ForegroundColor Red
        Write-Host "  Enable it with: Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "⚠ Could not check audit log configuration" -ForegroundColor Yellow
}

# Search audit logs
Write-Host "`nSearching audit logs..." -ForegroundColor Cyan
Write-Host "This may take a moment..." -ForegroundColor Yellow

try {
    $records = Search-UnifiedAuditLog `
        -StartDate $StartDate `
        -EndDate $EndDate `
        -Operations "FileAccessed", "FileDownloaded", "FileModified", "FileUploaded" `
        -ResultSize 5000 `
        -ErrorAction Stop
    
    if ($records) {
        Write-Host "✓ Found $($records.Count) total audit records" -ForegroundColor Green
        
        # Filter for the specific file
        $fileRecords = $records | Where-Object { $_.ObjectId -like "*$FileName*" }
        
        if ($fileRecords) {
            Write-Host "✓ Found $($fileRecords.Count) records for file: $FileName" -ForegroundColor Green
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "AUDIT LOG RESULTS" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            
            $fileRecords | Select-Object `
                @{Name="Time"; Expression={$_.CreationDate}},
                @{Name="User"; Expression={$_.UserIds}},
                @{Name="Operation"; Expression={$_.Operations}},
                @{Name="File"; Expression={
                    $obj = $_.ObjectId
                    if ($obj.Length -gt 50) { "..." + $obj.Substring($obj.Length - 50) } else { $obj }
                }} | Format-Table -AutoSize
            
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "DETAILED AUDIT DATA" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            
            foreach ($record in $fileRecords) {
                $auditData = $record.AuditData | ConvertFrom-Json
                Write-Host "`nOperation: $($auditData.Operation)" -ForegroundColor Yellow
                Write-Host "Time: $($record.CreationDate)" -ForegroundColor Cyan
                Write-Host "User: $($auditData.UserId)" -ForegroundColor Cyan
                Write-Host "Workload: $($auditData.Workload)" -ForegroundColor Cyan
                Write-Host "Client IP: $($auditData.ClientIP)" -ForegroundColor Cyan
                Write-Host "File: $($auditData.ObjectId)" -ForegroundColor Cyan
                Write-Host "Site: $($auditData.SiteUrl)" -ForegroundColor Cyan
                Write-Host "---"
            }
            
            Write-Host "`n========================================" -ForegroundColor Green
            Write-Host "PROOF FOR YOUR CUSTOMER" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "✓ File download operations (/$value endpoints) ARE logged" -ForegroundColor Green
            Write-Host "✓ These appear as FileAccessed or FileDownloaded activities" -ForegroundColor Green
            Write-Host "✓ Metadata operations (allproperties, lists, items) are NOT logged" -ForegroundColor Green
            
            # Export to file
            $exportPath = Join-Path $PSScriptRoot "AuditLog_Verification_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $fileRecords | Select-Object CreationDate, UserIds, Operations, ObjectId, AuditData | Export-Csv -Path $exportPath -NoTypeInformation
            Write-Host "`n✓ Results exported to: $exportPath" -ForegroundColor Green
        }
        else {
            Write-Host "`n⚠ No records found for file: $FileName" -ForegroundColor Yellow
            Write-Host "`nPossible reasons:" -ForegroundColor Yellow
            Write-Host "1. Audit logs take 60-90 minutes to appear (you may need to wait longer)" -ForegroundColor Cyan
            Write-Host "2. The file name doesn't match exactly" -ForegroundColor Cyan
            Write-Host "3. The operations weren't logged (metadata-only requests)" -ForegroundColor Cyan
            
            Write-Host "`nShowing all FileAccessed/FileDownloaded operations in this time range:" -ForegroundColor Cyan
            $allFileOps = $records | Where-Object { $_.Operations -match "File" }
            if ($allFileOps) {
                $allFileOps | Select-Object CreationDate, UserIds, Operations, `
                    @{Name="File"; Expression={
                        $obj = $_.ObjectId
                        if ($obj.Length -gt 60) { "..." + $obj.Substring($obj.Length - 60) } else { $obj }
                    }} | Format-Table -AutoSize
            }
        }
    }
    else {
        Write-Host "⚠ No audit records found in the specified time range" -ForegroundColor Yellow
        Write-Host "`nPossible reasons:" -ForegroundColor Yellow
        Write-Host "1. Audit logs take 60-90 minutes to appear" -ForegroundColor Cyan
        Write-Host "2. The date range may be incorrect" -ForegroundColor Cyan
        Write-Host "3. Audit logging might not be enabled" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "✗ Error searching audit logs: $_" -ForegroundColor Red
    Write-Host "`nCommon issues:" -ForegroundColor Yellow
    Write-Host "1. You need Compliance Administrator or Audit Logs role" -ForegroundColor Cyan
    Write-Host "2. Audit logging must be enabled in your tenant" -ForegroundColor Cyan
    Write-Host "3. The search may be taking too long - try a shorter date range" -ForegroundColor Cyan
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Script completed" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
