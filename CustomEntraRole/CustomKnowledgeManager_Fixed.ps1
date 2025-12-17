# Sign in with the needed scopes
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory","Directory.Read.All"

# 1) Get the built-in Knowledge Manager role definition (directory role, not Azure RBAC)
Write-Host "Step 1: Getting Knowledge Manager role definition..." -ForegroundColor Green
$km = Get-MgRoleManagementDirectoryRoleDefinition -All `
  | Where-Object { $_.DisplayName -eq "Knowledge Manager" }

if (-not $km) { throw "Built-in role 'Knowledge Manager' not found." }
Write-Host "Found Knowledge Manager role: $($km.DisplayName)" -ForegroundColor Green
Write-Host "Role ID: $($km.Id)" -ForegroundColor Green

# 2) Extract allowed resource actions from the built-in role
Write-Host "`nStep 2: Extracting permissions..." -ForegroundColor Green
$allowed = @()

# Extract permissions from RolePermissions property
foreach ($perm in $km.RolePermissions) {
    if ($perm.AllowedResourceActions) {
        $allowed += $perm.AllowedResourceActions
    }
}

$allowed = $allowed | Sort-Object -Unique   # de-duplicate cleanly

Write-Host "Total permissions found: $($allowed.Count)" -ForegroundColor Green
Write-Host "All permissions in Knowledge Manager role:" -ForegroundColor Green
$allowed | ForEach-Object { Write-Host "  - $_" }

# 3) Remove the six group write/update actions you specified
Write-Host "`nStep 3: Removing excluded permissions..." -ForegroundColor Green
$excluded = @(
  "microsoft.directory/groups.security/basic/update",
  "microsoft.directory/groups.security/create",
  "microsoft.directory/groups.security/createAsOwner",
  "microsoft.directory/groups.security/delete",
  "microsoft.directory/groups.security/members/update",
  "microsoft.directory/groups.security/owners/update"
)

Write-Host "Permissions to exclude:" -ForegroundColor Yellow
$excluded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }

$clean = $allowed | Where-Object { $excluded -notcontains $_ }

Write-Host "`nAfter filtering - permissions remaining: $($clean.Count)" -ForegroundColor Green

# Check if any permissions were actually removed
$removedCount = $allowed.Count - $clean.Count
Write-Host "Permissions removed: $removedCount" -ForegroundColor Green

# 4) Remove permissions that are not supported in custom roles
Write-Host "`nStep 4: Removing unsupported permissions for custom roles..." -ForegroundColor Green

# These are known to be unsupported in custom roles based on the error message
$unsupportedInCustomRoles = @(
    "microsoft.office365.knowledge/contentUnderstanding/analytics/allProperties/read",
    "microsoft.office365.knowledge/knowledgeNetwork/topicVisibility/allProperties/allTasks"
)

Write-Host "Unsupported permissions in custom roles:" -ForegroundColor Yellow
$unsupportedInCustomRoles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }

$cleanForCustomRole = $clean | Where-Object { $unsupportedInCustomRoles -notcontains $_ }

Write-Host "`nAfter removing unsupported permissions - remaining: $($cleanForCustomRole.Count)" -ForegroundColor Green
Write-Host "Final permissions for custom role:" -ForegroundColor Green
$cleanForCustomRole | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }

if ($cleanForCustomRole.Count -eq 0) {
    Write-Host "Warning: No permissions remaining after filtering!" -ForegroundColor Red
    Write-Host "The custom role will be created but will have no permissions." -ForegroundColor Red
}

# 5) Build the custom role definition body
Write-Host "`nStep 5: Building role definition..." -ForegroundColor Green
$body = @{
  displayName   = "Knowledge Manager – Limited (no SecGroup write)"
  description   = "Clone of built-in Knowledge Manager without Security Group create/delete/update permissions."
  isEnabled     = $true
  rolePermissions = @(@{ allowedResourceActions = $cleanForCustomRole })
}

Write-Host "Role definition body created:" -ForegroundColor Green
Write-Host "  Display Name: $($body.displayName)" -ForegroundColor Green
Write-Host "  Description: $($body.description)" -ForegroundColor Green
Write-Host "  Is Enabled: $($body.isEnabled)" -ForegroundColor Green
Write-Host "  Role Permissions Count: $($body.rolePermissions[0].allowedResourceActions.Count)" -ForegroundColor Green

# 6) Create the custom directory role
Write-Host "`nStep 6: Creating custom directory role..." -ForegroundColor Green
try {
    $result = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $body
    Write-Host "Success! Custom role created:" -ForegroundColor Green
    Write-Host "  Role ID: $($result.Id)" -ForegroundColor Green
    Write-Host "  Display Name: $($result.DisplayName)" -ForegroundColor Green
    Write-Host "  Description: $($result.Description)" -ForegroundColor Green
    Write-Host "  Is Enabled: $($result.IsEnabled)" -ForegroundColor Green
    
    Write-Host "`nRole created successfully! You can now assign this role to users." -ForegroundColor Green
    Write-Host "The role excludes the following group security permissions:" -ForegroundColor Green
    $excluded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    
} catch {
    Write-Host "Error creating role: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nFull error details:" -ForegroundColor Red
    $_.Exception | Format-List * -Force
    
    # Additional debugging
    Write-Host "`nDetailed body inspection:" -ForegroundColor Yellow
    $body | ConvertTo-Json -Depth 10
}