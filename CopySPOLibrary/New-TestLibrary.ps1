<#
.SYNOPSIS
    Creates a test document library with custom columns and dummy files for testing Copy-SPOLibrary.ps1.

.PARAMETER SiteUrl
    The URL of the SharePoint site where the test library will be created.

.PARAMETER LibraryName
    The name of the test library. Defaults to "TestSourceLibrary".

.EXAMPLE
    .\New-TestLibrary.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/source"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = "https://YOURTENANT.sharepoint.com/sites/YourSite/",

    [Parameter(Mandatory = $false)]
    [string]$LibraryName = "TestSourceLibrary",

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "YOUR-APP-CLIENT-ID"
)

#Requires -Modules PnP.PowerShell

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        "ERROR"   { Write-Host "[$ts] ERROR: $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "[$ts] WARNING: $Message" -ForegroundColor Yellow }
        "SUCCESS" { Write-Host "[$ts] SUCCESS: $Message" -ForegroundColor Green }
        default   { Write-Host "[$ts] INFO: $Message" -ForegroundColor Cyan }
    }
}

# ============================================================
# Connect
# ============================================================
Write-Log "Connecting to $SiteUrl"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ============================================================
# Create library
# ============================================================
$existing = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Log "Library '$LibraryName' already exists. Removing it first..." "WARNING"
    Remove-PnPList -Identity $LibraryName -Force
    Write-Log "Removed existing library."
}

Write-Log "Creating library: $LibraryName"
New-PnPList -Title $LibraryName -Template DocumentLibrary | Out-Null
Write-Log "Library created." "SUCCESS"

# ============================================================
# Add custom columns
# ============================================================
Write-Log "Adding custom columns..."

# 1. Single line of text
Add-PnPField -List $LibraryName -DisplayName "Project Name" -InternalName "ProjectName" -Type Text -AddToDefaultView | Out-Null
Write-Log "Added: Project Name (Text)" "SUCCESS"

# 2. Multi-line text
Add-PnPField -List $LibraryName -DisplayName "Description" -InternalName "ProjectDescription" -Type Note -AddToDefaultView | Out-Null
Write-Log "Added: Description (Note)" "SUCCESS"

# 3. Choice column
Add-PnPField -List $LibraryName -DisplayName "Department" -InternalName "Department" -Type Choice -Choices "IT","HR","Finance","Marketing","Operations" -AddToDefaultView | Out-Null
Write-Log "Added: Department (Choice)" "SUCCESS"

# 4. Number column
Add-PnPField -List $LibraryName -DisplayName "Budget" -InternalName "Budget" -Type Number -AddToDefaultView | Out-Null
Write-Log "Added: Budget (Number)" "SUCCESS"

# 5. Date column
Add-PnPField -List $LibraryName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView | Out-Null
Write-Log "Added: Due Date (DateTime)" "SUCCESS"

# 6. Yes/No column
Add-PnPField -List $LibraryName -DisplayName "Is Approved" -InternalName "IsApproved" -Type Boolean -AddToDefaultView | Out-Null
Write-Log "Added: Is Approved (Boolean)" "SUCCESS"

# 7. Person column
Add-PnPField -List $LibraryName -DisplayName "Project Owner" -InternalName "ProjectOwner" -Type User -AddToDefaultView | Out-Null
Write-Log "Added: Project Owner (User)" "SUCCESS"

# 8. Hyperlink column
Add-PnPField -List $LibraryName -DisplayName "Reference Link" -InternalName "ReferenceLink" -Type URL -AddToDefaultView | Out-Null
Write-Log "Added: Reference Link (URL)" "SUCCESS"

Write-Log "All custom columns created." "SUCCESS"

# ============================================================
# Create folder structure
# ============================================================
Write-Log "Creating folder structure..."

$folders = @(
    "Finance",
    "Finance/Reports",
    "Finance/Invoices",
    "HR",
    "HR/Policies",
    "IT",
    "IT/Architecture",
    "IT/Architecture/Diagrams"
)

foreach ($folder in $folders) {
    Resolve-PnPFolder -SiteRelativePath "$LibraryName/$folder" | Out-Null
    Write-Log "Created folder: $folder" "SUCCESS"
}

# ============================================================
# Create and upload dummy files with metadata
# ============================================================
Write-Log "Creating and uploading dummy files..."

$tempDir = Join-Path $env:TEMP "TestLibraryFiles_$(Get-Date -Format 'yyyyMMddHHmmss')"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Define test files with their folder location and metadata
$testFiles = @(
    @{
        Name     = "ProjectPlan.txt"
        Folder   = ""
        Content  = "This is the main project plan document.`nVersion 1.0`nCreated for testing purposes."
        Metadata = @{
            ProjectName        = "Digital Transformation"
            ProjectDescription = "Company-wide digital transformation initiative covering all departments."
            Department         = "IT"
            Budget             = 150000
            DueDate            = "2026-06-30"
            IsApproved         = $true
            ReferenceLink      = "https://example.com/project-plan, Project Plan Portal"
        }
    },
    @{
        Name     = "Q1_Budget_Report.txt"
        Folder   = "Finance/Reports"
        Content  = "Q1 Budget Report`n==================`nTotal Budget: 500,000`nSpent: 320,000`nRemaining: 180,000"
        Metadata = @{
            ProjectName        = "Q1 Financial Review"
            ProjectDescription = "Quarterly budget analysis and reporting."
            Department         = "Finance"
            Budget             = 500000
            DueDate            = "2026-04-15"
            IsApproved         = $true
        }
    },
    @{
        Name     = "Invoice_001.txt"
        Folder   = "Finance/Invoices"
        Content  = "Invoice #001`nVendor: Contoso Ltd`nAmount: 12,500.00`nDue: 2026-04-01"
        Metadata = @{
            ProjectName = "Vendor Payments"
            Department  = "Finance"
            Budget      = 12500
            DueDate     = "2026-04-01"
            IsApproved  = $false
        }
    },
    @{
        Name     = "Invoice_002.txt"
        Folder   = "Finance/Invoices"
        Content  = "Invoice #002`nVendor: Fabrikam Inc`nAmount: 8,750.00`nDue: 2026-04-15"
        Metadata = @{
            ProjectName = "Vendor Payments"
            Department  = "Finance"
            Budget      = 8750
            DueDate     = "2026-04-15"
            IsApproved  = $true
        }
    },
    @{
        Name     = "Employee_Handbook.txt"
        Folder   = "HR/Policies"
        Content  = "Employee Handbook v3.2`n=======================`nChapter 1: Introduction`nChapter 2: Code of Conduct`nChapter 3: Benefits"
        Metadata = @{
            ProjectName        = "HR Policy Update"
            ProjectDescription = "Annual review and update of all HR policies and handbooks."
            Department         = "HR"
            DueDate            = "2026-05-01"
            IsApproved         = $true
        }
    },
    @{
        Name     = "Leave_Policy.txt"
        Folder   = "HR/Policies"
        Content  = "Leave Policy 2026`n==================`nAnnual Leave: 25 days`nSick Leave: 10 days`nParental Leave: 16 weeks"
        Metadata = @{
            ProjectName = "HR Policy Update"
            Department  = "HR"
            DueDate     = "2026-05-01"
            IsApproved  = $false
        }
    },
    @{
        Name     = "Onboarding_Checklist.txt"
        Folder   = "HR"
        Content  = "New Employee Onboarding Checklist`n1. IT Setup`n2. Badge Access`n3. Orientation Session`n4. Team Introduction"
        Metadata = @{
            ProjectName = "Onboarding Process"
            Department  = "HR"
            IsApproved  = $true
        }
    },
    @{
        Name     = "Network_Architecture.txt"
        Folder   = "IT/Architecture"
        Content  = "Corporate Network Architecture`n==============================`nVLAN 10: Corporate`nVLAN 20: Guest`nVLAN 30: IoT"
        Metadata = @{
            ProjectName        = "Network Redesign"
            ProjectDescription = "Redesigning corporate network for zero-trust architecture."
            Department         = "IT"
            Budget             = 75000
            DueDate            = "2026-08-01"
            IsApproved         = $false
            ReferenceLink      = "https://example.com/network-docs, Network Documentation"
        }
    },
    @{
        Name     = "Cloud_Migration_Plan.txt"
        Folder   = "IT/Architecture/Diagrams"
        Content  = "Cloud Migration Plan`n====================`nPhase 1: Assessment`nPhase 2: Pilot Migration`nPhase 3: Full Migration"
        Metadata = @{
            ProjectName        = "Cloud Migration"
            ProjectDescription = "Migrating on-premises workloads to Azure."
            Department         = "IT"
            Budget             = 250000
            DueDate            = "2026-12-31"
            IsApproved         = $true
            ReferenceLink      = "https://example.com/azure-migration, Azure Migration Guide"
        }
    },
    @{
        Name     = "Marketing_Campaign_Brief.txt"
        Folder   = ""
        Content  = "Spring 2026 Campaign Brief`n===========================`nTarget: Enterprise customers`nChannels: LinkedIn, Email, Webinars"
        Metadata = @{
            ProjectName        = "Spring Campaign 2026"
            ProjectDescription = "Multi-channel marketing campaign targeting enterprise segment."
            Department         = "Marketing"
            Budget             = 45000
            DueDate            = "2026-04-30"
            IsApproved         = $true
        }
    }
)

foreach ($file in $testFiles) {
    # Create the temp file
    $tempFile = Join-Path $tempDir $file.Name
    Set-Content -Path $tempFile -Value $file.Content -Encoding UTF8

    # Determine upload folder
    $targetFolder = $LibraryName
    if ($file.Folder) {
        $targetFolder = "$LibraryName/$($file.Folder)"
    }

    # Upload and set metadata in one call
    if ($file.Metadata.Count -gt 0) {
        Add-PnPFile -Path $tempFile -Folder $targetFolder -Values $file.Metadata | Out-Null
    } else {
        Add-PnPFile -Path $tempFile -Folder $targetFolder | Out-Null
    }

    $displayPath = if ($file.Folder) { "$($file.Folder)/$($file.Name)" } else { $file.Name }
    Write-Log "Uploaded: $displayPath" "SUCCESS"
}

# Cleanup temp files
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
# Summary
# ============================================================
Disconnect-PnPOnline

Write-Log "============================================"
Write-Log "Test library setup complete!" "SUCCESS"
Write-Log "  Library: $LibraryName"
Write-Log "  Custom columns: 8"
Write-Log "  Folders: $($folders.Count)"
Write-Log "  Files: $($testFiles.Count)"
Write-Log "============================================"
Write-Log ""
Write-Log "You can now test Copy-SPOLibrary.ps1 with:"
Write-Log "  .\Copy-SPOLibrary.ps1 -SourceSiteUrl '$SiteUrl' -TargetSiteUrl '<TARGET_URL>' -SourceLibraryName '$LibraryName'"
