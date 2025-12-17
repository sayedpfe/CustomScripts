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
foreach ($perm in $km.Permissions) {
  $allowed += $perm.AllowedResourceActions
}
$allowed = $allowed | Sort-Object -Unique   # de-duplicate cleanly

Write-Host "Total permissions found: $($allowed.Count)" -ForegroundColor Green
Write-Host "First few permissions:" -ForegroundColor Green
$allowed | Select-Object -First 5 | ForEach-Object { Write-Host "  - $_" }

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

# Check for null or empty values in $clean
$nullCount = ($clean | Where-Object { [string]::IsNullOrEmpty($_) }).Count
Write-Host "Null or empty permissions in clean array: $nullCount" -ForegroundColor Yellow

if ($nullCount -gt 0) {
    Write-Host "Warning: Found null or empty permissions. Cleaning up..." -ForegroundColor Red
    $clean = $clean | Where-Object { -not [string]::IsNullOrEmpty($_) }
    Write-Host "After cleanup - permissions remaining: $($clean.Count)" -ForegroundColor Green
}

# 4) Build the custom role definition body
Write-Host "`nStep 4: Building role definition..." -ForegroundColor Green
$body = @{
  displayName   = "Knowledge Manager – Limited (no SecGroup write)"
  description   = "Clone of built-in Knowledge Manager without Security Group create/delete/update."
  isEnabled     = $true
  rolePermissions = @(@{ allowedResourceActions = $clean })
  isBuiltIn     = $false
  # Assignable at the directory root ("/") so it behaves like a directory role
  # If your org prefers scoping to specific administrative units, change this
  # to their resource IDs.
  scopeIds      = @("/")  
}

Write-Host "Role definition body created:" -ForegroundColor Green
Write-Host "  Display Name: $($body.displayName)" -ForegroundColor Green
Write-Host "  Description: $($body.description)" -ForegroundColor Green
Write-Host "  Is Enabled: $($body.isEnabled)" -ForegroundColor Green
Write-Host "  Is Built In: $($body.isBuiltIn)" -ForegroundColor Green
Write-Host "  Scope IDs: $($body.scopeIds -join ', ')" -ForegroundColor Green
Write-Host "  Role Permissions Count: $($body.rolePermissions[0].allowedResourceActions.Count)" -ForegroundColor Green

# Show a sample of the permissions that will be included
Write-Host "`nSample permissions that will be included:" -ForegroundColor Green
$body.rolePermissions[0].allowedResourceActions | Select-Object -First 10 | ForEach-Object { 
    Write-Host "  - $_" -ForegroundColor Green 
}

# 5) Create the custom directory role
Write-Host "`nStep 5: Creating custom directory role..." -ForegroundColor Green
try {
    $result = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $body
    Write-Host "Success! Custom role created:" -ForegroundColor Green
    Write-Host "  Role ID: $($result.Id)" -ForegroundColor Green
    Write-Host "  Display Name: $($result.DisplayName)" -ForegroundColor Green
} catch {
    Write-Host "Error creating role: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details:" -ForegroundColor Red
    $_ | Format-List * -Force
    
    # Additional debugging
    Write-Host "`nDetailed body inspection:" -ForegroundColor Yellow
    $body | ConvertTo-Json -Depth 10
}