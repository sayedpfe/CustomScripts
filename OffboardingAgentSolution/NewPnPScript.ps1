
<# 
.SYNOPSIS
    Connects to a SharePoint Online site using PnP.PowerShell and exports all lists/libraries.

.DESCRIPTION
    Supports three auth modes:
      - Interactive: Opens a browser window (recommended for admins)
      - DeviceCode: For headless environments (copy-paste code flow)
      - AppCert: App-only with certificate (daemon/non-interactive)
    Exports key properties for each list and prints a summary table.

.PARAMETER SiteUrl
    Full URL of the SharePoint site (e.g., https://contoso.sharepoint.com/sites/ProjectX)

.PARAMETER OutputCsv
    Optional full path to a CSV file to save results. If omitted, shows output only.

.PARAMETER IncludeHidden
    Include hidden lists (like Form Templates, Workflow History, etc.). Default: false.

.PARAMETER AuthMode
    One of: Interactive, DeviceCode, AppCert

.PARAMETER Tenant
    Only for AppCert. Your tenant short name or full domain (e.g., contoso.onmicrosoft.com)

.PARAMETER ClientId
    Only for AppCert. Azure AD App (Entra ID) Application (client) ID with Sites.* permissions.

.PARAMETER CertificatePath
    Only for AppCert. Path to a PFX certificate file.

.PARAMETER CertificatePassword
    Only for AppCert. SecureString password for the PFX.

.EXAMPLE
    .\Get-SPOLists-PnP.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" -AuthMode Interactive

.EXAMPLE
    .\Get-SPOLists-PnP.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" -AuthMode DeviceCode -OutputCsv "C:\Temp\ProjectX_Lists.csv"

.EXAMPLE
    $pwd = Read-Host "PFX password" -AsSecureString
    .\Get-SPOLists-PnP.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" `
        -AuthMode AppCert `
        -Tenant "contoso.onmicrosoft.com" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -CertificatePath "C:\certs\pnp-app.pfx" `
        -CertificatePassword $pwd `
        -OutputCsv "C:\Temp\lists.csv"

.NOTES
    Requires: PnP.PowerShell
    App-only requires Graph/SharePoint app permissions (Sites.Read.All or Sites.FullControl.All) 
    AND site collection admin or sufficient site permissions in many cases.
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeHidden,

    # --- Auth Modes ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [ValidateSet('Interactive')]
    [string]$AuthMode,

    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$InteractiveClientId,

    [Parameter(Mandatory = $true, ParameterSetName = 'DeviceCode')]
    [ValidateSet('DeviceCode')]
    [string]$AuthMode_DeviceCode,

    [Parameter(Mandatory = $true, ParameterSetName = 'AppCert')]
    [ValidateSet('AppCert')]
    [string]$AuthMode_AppCert,

    # --- AppCert params ---
    [Parameter(Mandatory = $true, ParameterSetName = 'AppCert')]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant,

    [Parameter(Mandatory = $true, ParameterSetName = 'AppCert')]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true, ParameterSetName = 'AppCert')]
    [ValidateScript({ Test-Path $_ })]
    [string]$CertificatePath,

    [Parameter(Mandatory = $true, ParameterSetName = 'AppCert')]
    [System.Security.SecureString]$CertificatePassword
)

begin {
    $ErrorActionPreference = 'Stop'
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    function Write-Info($msg)  { Write-Host "[INFO]  $msg" -ForegroundColor Cyan }
    function Write-Warn($msg)  { Write-Host "[WARN]  $msg" -ForegroundColor Yellow }
    function Write-Err ($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

    # Normalize which mode was chosen
    switch ($PSCmdlet.ParameterSetName) {
        'Interactive' { $mode = 'Interactive' }
        'DeviceCode'  { $mode = 'DeviceCode' }
        'AppCert'     { $mode = 'AppCert' }
        default       { throw "Unknown parameter set: $($PSCmdlet.ParameterSetName)" }
    }

    Write-Info "Connecting to: $SiteUrl"
    Write-Info "Auth mode: $mode"
}

process {
    try {
        # --- Connect ---
        switch ($mode) {
            'Interactive' {
                if ($InteractiveClientId) {
                    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $InteractiveClientId
                } else {
                    Connect-PnPOnline -Url $SiteUrl -Interactive
                }
            }
            'DeviceCode' {
                Connect-PnPOnline -Url $SiteUrl -DeviceLogin
            }
            'AppCert' {
                # App-only: requires App Registration with Sites.Read.All (or Sites.FullControl.All) consented
                Connect-PnPOnline -Url $SiteUrl `
                    -Tenant $Tenant `
                    -ClientId $ClientId `
                    -CertificatePath $CertificatePath `
                    -CertificatePassword $CertificatePassword
            }
        }

        # --- Retrieve lists ---
        # Select only needed properties to improve performance.
        $lists = Get-PnPList -Includes RootFolder, Title, ItemCount, BaseTemplate, BaseType, Hidden, Created, LastItemUserModifiedDate, Id, DefaultViewUrl, ContentTypesEnabled

        if (-not $IncludeHidden) {
            $lists = $lists | Where-Object { $_.Hidden -eq $false }
        }

        # Map to a friendly object
        $result = $lists | Select-Object `
            @{n='Title';        e={$_.Title}},
            @{n='Url';          e={ 
                if ($_.RootFolder -and $_.RootFolder.ServerRelativeUrl) {
                    # Build absolute URL from server-relative path
                    $root = (Get-PnPWeb).Url.TrimEnd('/')
                    "$root$($_.RootFolder.ServerRelativeUrl)"
                } else {
                    # Fallback to DefaultViewUrl if available
                    $root = (Get-PnPWeb).Url.TrimEnd('/')
                    if ($_.DefaultViewUrl) { "$root$($_.DefaultViewUrl)" } else { '' }
                }
            }},
            @{n='Id';           e={$_.Id}},
            @{n='Hidden';       e={$_.Hidden}},
            @{n='BaseType';     e={$_.BaseType}},           # GenericList / DocumentLibrary / etc.
            @{n='BaseTemplate'; e={$_.BaseTemplate}},       # Numeric template ID (101=DocLib, 100=Custom List, etc.)
            @{n='ItemCount';    e={$_.ItemCount}},
            @{n='Created';      e={$_.Created}},
            @{n='LastModified'; e={$_.LastItemUserModifiedDate}},
            @{n='HasContentTypes'; e={$_.ContentTypesEnabled}},
            @{n='DefaultViewUrl'; e={$_.DefaultViewUrl}}

        if (-not $result) {
            Write-Warn "No lists found. Check permissions or site URL."
        } else {
            # --- Output table to console ---
            $result | Sort-Object Title | Format-Table -AutoSize Title, BaseType, BaseTemplate, ItemCount, Hidden, LastModified

            # --- Optionally export ---
            if ($OutputCsv) {
                $dir = Split-Path -Path $OutputCsv -Parent
                if ($dir -and -not (Test-Path $dir)) {
                    New-Item -ItemType Directory -Path $dir | Out-Null
                }
                $result | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
                Write-Info "Exported to: $OutputCsv"
            }

            Write-Info ("Lists returned: {0}" -f $result.Count)
        }
    }
    catch {
        Write-Err $_.Exception.Message
        if ($_.InvocationInfo.PositionMessage) {
            Write-Err $_.InvocationInfo.PositionMessage
        }
        throw
    }
}
end {
    $stopwatch.Stop()
    Write-Info ("Completed in {0:n1} seconds." -f $stopwatch.Elapsed.TotalSeconds)
}
