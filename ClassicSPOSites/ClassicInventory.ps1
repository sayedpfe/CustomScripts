<#
.SYNOPSIS
    Inventory Classic SharePoint Publishing Site Collections
    Covers MC714186 and MC1117115 retirement announcements

.DESCRIPTION
    Scans all site collections in the tenant and identifies those using
    classic publishing templates that are affected by Microsoft's retirement
    announcements. Exports results to CSV.

.REQUIREMENTS
    - PnP PowerShell (PnP.PowerShell module)
    - SharePoint Administrator role
    - App Registration with delegated permissions (Sites.FullControl.All)

.NOTES
    Install/update the module if needed:
    Install-Module -Name PnP.PowerShell -Force
#>

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION — update these values before running
# ─────────────────────────────────────────────────────────────────────────────
$TenantName   = "YOURTENANT"           # e.g. "contoso" (not the full URL)
$ClientId     = "YOUR-APP-CLIENT-ID"    # Your App Registration Client ID
$OutputPath   = "C:\Reports\ClassicSiteInventory_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$AdminUrl     = "https://$TenantName-admin.sharepoint.com"

# ─────────────────────────────────────────────────────────────────────────────
# Classic publishing templates affected by MC714186 / MC1117115
# ─────────────────────────────────────────────────────────────────────────────
$ClassicTemplates = @{
    "BLANKINTERNETCONTAINER#0" = "Classic Publishing Portal Site"
    "CMSPUBLISHING#0"          = "Classic Publishing Site"
    "BLANKINTERNET#0"          = "Classic Publishing Site Blank"
    "ENTERWIKI#0"              = "Enterprise Wiki"
    "SRCHCEN#0"                = "Enterprise Search Center"
    "SPSSITES#0"               = "Site Directory"
    "SPSNHOME#0"               = "News Home Site"
    "PRODUCTCATALOG#0"         = "Product Catalog"
    "SPSREPORTCENTER#0"        = "Report Center"
    "SPSTOPIC#0"               = "Topic Area Template"
    "CSPCONTAINER#0"           = "SharePoint Embedded Site"
}

# ─────────────────────────────────────────────────────────────────────────────
# Connect
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`nConnecting to SharePoint Online admin (interactive login)..." -ForegroundColor Cyan
Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId

# ─────────────────────────────────────────────────────────────────────────────
# Retrieve all site collections (batched to handle large tenants)
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "Retrieving all site collections (this may take a while for large tenants)..." -ForegroundColor Cyan

$AllSites = Get-PnPTenantSite -IncludeOneDriveSites:$false

Write-Host "Total site collections retrieved: $($AllSites.Count)" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
# Filter to classic publishing sites only
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`nFiltering for classic publishing site templates..." -ForegroundColor Cyan

$Results = @()

foreach ($Site in $AllSites) {
    if ($ClassicTemplates.ContainsKey($Site.Template)) {

        # Retrieve extra detail per site
        try {
            $SiteDetail = Get-PnPTenantSite -Identity $Site.Url -Detailed -ErrorAction Stop
        } catch {
            $SiteDetail = $Site
        }

        $Results += [PSCustomObject]@{
            "Site Title"              = $SiteDetail.Title
            "URL"                     = $SiteDetail.Url
            "Template ID"             = $SiteDetail.Template
            "Template Description"    = $ClassicTemplates[$SiteDetail.Template]
            "Custom Scripting Status" = if ($SiteDetail.DenyAddAndCustomizePages -eq "Disabled") { "WARNING: ENABLED (at risk)" } else { "OK: Disabled (compliant)" }
            "Storage Used (MB)"       = [math]::Round($SiteDetail.StorageUsageCurrent, 2)
            "Owner"                   = $SiteDetail.Owner
            "Last Content Modified"   = $SiteDetail.LastContentModifiedDate
            "Status"                  = $SiteDetail.Status
            "MC1117115 Impact"        = "Blocked from creation; custom scripting disabled by default"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Output results
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`nClassic Publishing Sites Found: $($Results.Count)" -ForegroundColor Yellow

if ($Results.Count -gt 0) {

    # Console summary grouped by template
    Write-Host "`n-- Breakdown by Template --" -ForegroundColor Cyan
    $Results | Group-Object "Template ID" | ForEach-Object {
        Write-Host ("  {0,-35} -> {1} site(s)" -f $ClassicTemplates[$_.Name], $_.Count) -ForegroundColor White
    }

    # Sites with custom scripting still ENABLED (highest risk)
    $RiskySites = $Results | Where-Object { $_."Custom Scripting Status" -like "*ENABLED*" }
    if ($RiskySites.Count -gt 0) {
        Write-Host "`nWARNING: Sites with Custom Scripting still ENABLED ($($RiskySites.Count)) -- highest priority to action:" -ForegroundColor Red
        $RiskySites | Select-Object "Site Title", "URL", "Template Description" | Format-Table -AutoSize
    }

    # Export to CSV
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nFull report exported to: $OutputPath" -ForegroundColor Green

} else {
    Write-Host "`nNo classic publishing site collections found in this tenant." -ForegroundColor Green
}

# ─────────────────────────────────────────────────────────────────────────────
# Optional: Check tenant-level opt-out flag (MC1117115)
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n-- Tenant-Level MC1117115 Opt-Out Status --" -ForegroundColor Cyan
$TenantConfig = Get-PnPTenant

$OptOutStatus = $TenantConfig.DelayDenyAddAndCustomizePagesEnforcementOnClassicPublishingSites
Write-Host "  DelayDenyAddAndCustomizePagesEnforcementOnClassicPublishingSites : $OptOutStatus" -ForegroundColor $(if ($OptOutStatus) { "Yellow" } else { "Green" })

if ($OptOutStatus) {
    Write-Host "  WARNING: Opt-out is currently ACTIVE. This will stop working on March 15, 2026." -ForegroundColor Yellow
} else {
    Write-Host "  OK: Opt-out not active. Enforcement is running as per MC1117115." -ForegroundColor Green
}

$AllowCreation = $TenantConfig.AllowClassicPublishingSiteCreation
Write-Host "  AllowClassicPublishingSiteCreation                              : $AllowCreation" -ForegroundColor $(if ($AllowCreation) { "Yellow" } else { "Green" })

if ($AllowCreation) {
    Write-Host "  WARNING: Classic publishing site creation is currently RE-ENABLED by admin. Review if still needed." -ForegroundColor Yellow
}

Write-Host "`nInventory complete.`n" -ForegroundColor Cyan