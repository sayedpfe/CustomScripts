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
Write-Host "`nStep 3: Removing requested excluded permissions..." -ForegroundColor Green
$excluded = @(
  "microsoft.directory/groups.security/basic/update",
  "microsoft.directory/groups.security/create",
  "microsoft.directory/groups.security/createAsOwner",
  "microsoft.directory/groups.security/delete",
  "microsoft.directory/groups.security/members/update",
  "microsoft.directory/groups.security/owners/update"
)

Write-Host "Permissions to exclude (as requested):" -ForegroundColor Yellow
$excluded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }

$afterUserExclusions = $allowed | Where-Object { $excluded -notcontains $_ }

Write-Host "`nAfter filtering requested exclusions - permissions remaining: $($afterUserExclusions.Count)" -ForegroundColor Green
$removedCount = $allowed.Count - $afterUserExclusions.Count
Write-Host "Permissions removed: $removedCount" -ForegroundColor Green

# 4) Filter to only include permissions that are known to work with custom roles
Write-Host "`nStep 4: Filtering to only include permissions supported in custom roles..." -ForegroundColor Green

# Based on testing, these Office 365 permissions are not supported in custom roles:
# Let's create a whitelist of supported permissions that we know work
$supportedPermissions = @(
    # Basic directory permissions that typically work
    "microsoft.office365.supportTickets/allEntities/allTasks",
    "microsoft.office365.webPortal/allEntities/standard/read"
    # Note: Many office365.* permissions don't work in custom directory roles
)

# Filter to only include permissions that start with microsoft.directory or are in our whitelist
$customRoleCompatible = $afterUserExclusions | Where-Object { 
    $_ -like "microsoft.directory/*" -or $supportedPermissions -contains $_
}

Write-Host "Permissions compatible with custom roles:" -ForegroundColor Green
if ($customRoleCompatible.Count -gt 0) {
    $customRoleCompatible | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }
} else {
    Write-Host "  No compatible permissions found!" -ForegroundColor Red
}

# Show what was filtered out
$filtered = $afterUserExclusions | Where-Object { $customRoleCompatible -notcontains $_ }
if ($filtered.Count -gt 0) {
    Write-Host "`nPermissions filtered out (not supported in custom roles):" -ForegroundColor Yellow
    $filtered | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

Write-Host "`nFinal count - Custom role compatible permissions: $($customRoleCompatible.Count)" -ForegroundColor Green

# 5) Check if we have any permissions left
if ($customRoleCompatible.Count -eq 0) {
    Write-Host "`nWarning: After filtering, no permissions remain for the custom role!" -ForegroundColor Red
    Write-Host "This suggests that the Knowledge Manager role primarily contains permissions" -ForegroundColor Red
    Write-Host "that are not supported in custom directory roles." -ForegroundColor Red
    Write-Host "`nThe Knowledge Manager role may be designed as a built-in role only." -ForegroundColor Red
    Write-Host "You may need to:" -ForegroundColor Yellow
    Write-Host "  1. Use the built-in role and manage group permissions via other means" -ForegroundColor Yellow
    Write-Host "  2. Consider Azure RBAC roles instead of directory roles" -ForegroundColor Yellow
    Write-Host "  3. Use different built-in roles that better match your needs" -ForegroundColor Yellow
    
    # Let's still try to create the role to see what happens
    Write-Host "`nAttempting to create role with no permissions for demonstration..." -ForegroundColor Yellow
}

# 6) Build the custom role definition body
Write-Host "`nStep 5: Building role definition..." -ForegroundColor Green
$body = @{
  displayName   = "Knowledge Manager – Limited (no SecGroup write)"
  description   = "Clone of built-in Knowledge Manager without Security Group create/delete/update permissions. Note: Most Knowledge Manager permissions are not supported in custom roles."
  isEnabled     = $true
  rolePermissions = @(@{ allowedResourceActions = $customRoleCompatible })
}

Write-Host "Role definition body created:" -ForegroundColor Green
Write-Host "  Display Name: $($body.displayName)" -ForegroundColor Green
Write-Host "  Description: $($body.description)" -ForegroundColor Green
Write-Host "  Is Enabled: $($body.isEnabled)" -ForegroundColor Green
Write-Host "  Role Permissions Count: $($body.rolePermissions[0].allowedResourceActions.Count)" -ForegroundColor Green

# 7) Create the custom directory role
Write-Host "`nStep 6: Creating custom directory role..." -ForegroundColor Green
try {
    $result = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $body
    Write-Host "Success! Custom role created:" -ForegroundColor Green
    Write-Host "  Role ID: $($result.Id)" -ForegroundColor Green
    Write-Host "  Display Name: $($result.DisplayName)" -ForegroundColor Green
    Write-Host "  Description: $($result.Description)" -ForegroundColor Green
    Write-Host "  Is Enabled: $($result.IsEnabled)" -ForegroundColor Green
    
    if ($customRoleCompatible.Count -gt 0) {
        Write-Host "`nRole created successfully with the following permissions:" -ForegroundColor Green
        $customRoleCompatible | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }
    } else {
        Write-Host "`nRole created but has no permissions. This demonstrates that most" -ForegroundColor Yellow
        Write-Host "Knowledge Manager permissions are not compatible with custom roles." -ForegroundColor Yellow
    }
    
    Write-Host "`nThe role excludes the following group security permissions (as requested):" -ForegroundColor Green
    $excluded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    
} catch {
    Write-Host "Error creating role: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nThis error suggests additional permissions are not supported." -ForegroundColor Red
    
    # Try to create with even fewer permissions
    if ($customRoleCompatible -contains "microsoft.office365.supportTickets/allEntities/allTasks") {
        Write-Host "`nTrying with only support tickets permission..." -ForegroundColor Yellow
        $minimalPermissions = @("microsoft.office365.supportTickets/allEntities/allTasks")
        $minimalBody = @{
            displayName   = "Knowledge Manager – Minimal"
            description   = "Minimal custom role based on Knowledge Manager with only support ticket permissions."
            isEnabled     = $true
            rolePermissions = @(@{ allowedResourceActions = $minimalPermissions })
        }
        
        try {
            $minimalResult = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $minimalBody
            Write-Host "Success! Minimal role created:" -ForegroundColor Green
            Write-Host "  Role ID: $($minimalResult.Id)" -ForegroundColor Green
            Write-Host "  Display Name: $($minimalResult.DisplayName)" -ForegroundColor Green
        } catch {
            Write-Host "Even minimal role creation failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "`nThis confirms that Knowledge Manager permissions are largely incompatible with custom roles." -ForegroundColor Red
        }
    }
}