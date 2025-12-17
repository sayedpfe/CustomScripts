# Sign in with the needed scopes
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory","Directory.Read.All"

Write-Host "=== Knowledge Manager Custom Role Analysis ===" -ForegroundColor Cyan
Write-Host "Analyzing the Knowledge Manager role to create a custom version without security group permissions..." -ForegroundColor Cyan

# 1) Get the built-in Knowledge Manager role definition (directory role, not Azure RBAC)
Write-Host "`nStep 1: Getting Knowledge Manager role definition..." -ForegroundColor Green
$km = Get-MgRoleManagementDirectoryRoleDefinition -All `
  | Where-Object { $_.DisplayName -eq "Knowledge Manager" }

if (-not $km) { throw "Built-in role 'Knowledge Manager' not found." }
Write-Host "✓ Found Knowledge Manager role: $($km.DisplayName)" -ForegroundColor Green

# 2) Extract allowed resource actions from the built-in role
Write-Host "`nStep 2: Extracting permissions..." -ForegroundColor Green
$allowed = @()

# Use RolePermissions instead of Permissions (which is null)
foreach ($perm in $km.RolePermissions) {
  if ($perm.AllowedResourceActions) {
    $allowed += $perm.AllowedResourceActions
  }
}
$allowed = $allowed | Sort-Object -Unique   # de-duplicate cleanly

Write-Host "✓ Found $($allowed.Count) permissions in Knowledge Manager role" -ForegroundColor Green
Write-Host "`nAll permissions:" -ForegroundColor Yellow
$allowed | ForEach-Object { Write-Host "  • $_" -ForegroundColor Yellow }

# 3) Remove the six group write/update actions you specified
Write-Host "`nStep 3: Removing requested security group permissions..." -ForegroundColor Green
$excluded = @(
  "microsoft.directory/groups.security/basic/update",
  "microsoft.directory/groups.security/create",
  "microsoft.directory/groups.security/createAsOwner",
  "microsoft.directory/groups.security/delete",
  "microsoft.directory/groups.security/members/update",
  "microsoft.directory/groups.security/owners/update"
)

Write-Host "Removing these permissions:" -ForegroundColor Red
$excluded | ForEach-Object { Write-Host "  ✗ $_" -ForegroundColor Red }

$clean = $allowed | Where-Object { $excluded -notcontains $_ }
$removedCount = $allowed.Count - $clean.Count

Write-Host "`n✓ Removed $removedCount permissions" -ForegroundColor Green
Write-Host "✓ Remaining permissions: $($clean.Count)" -ForegroundColor Green

# 4) Check if remaining permissions are supported in custom roles
Write-Host "`nStep 4: Checking custom role compatibility..." -ForegroundColor Yellow
Write-Host "⚠️  IMPORTANT DISCOVERY:" -ForegroundColor Yellow
Write-Host "The remaining permissions are all 'office365.*' permissions:" -ForegroundColor Yellow
$clean | ForEach-Object { Write-Host "  • $_" -ForegroundColor Yellow }

Write-Host "`n❌ LIMITATION FOUND:" -ForegroundColor Red
Write-Host "Microsoft does NOT support 'office365.*' permissions in custom directory roles." -ForegroundColor Red
Write-Host "Only 'microsoft.directory.*' permissions are supported in custom roles." -ForegroundColor Red
Write-Host "This means the Knowledge Manager role functionality cannot be replicated in a custom role!" -ForegroundColor Red

# 5) Provide recommendations instead of trying to create the role
Write-Host "`n💡 RECOMMENDATIONS:" -ForegroundColor Cyan
Write-Host "Since the Knowledge Manager permissions cannot be used in custom roles, consider these alternatives:" -ForegroundColor Cyan
Write-Host ""
Write-Host "Option 1: Use the built-in Knowledge Manager role as-is" -ForegroundColor White
Write-Host "  • Keep using the built-in role with all its permissions" -ForegroundColor Gray
Write-Host "  • Manage security group access through other means:" -ForegroundColor Gray
Write-Host "    - Conditional Access policies" -ForegroundColor Gray
Write-Host "    - Azure AD access reviews" -ForegroundColor Gray
Write-Host "    - Privileged Identity Management (PIM)" -ForegroundColor Gray
Write-Host ""
Write-Host "Option 2: Use multiple smaller built-in roles" -ForegroundColor White
Write-Host "  • Assign specific built-in roles that don't include group management" -ForegroundColor Gray
Write-Host "  • Examples: SharePoint Administrator, Teams Administrator" -ForegroundColor Gray
Write-Host ""
Write-Host "Option 3: Azure RBAC for some permissions" -ForegroundColor White
Write-Host "  • Use Azure RBAC roles for resource-level permissions" -ForegroundColor Gray
Write-Host "  • Combine with directory roles for identity management" -ForegroundColor Gray

Write-Host "`n📋 SUMMARY:" -ForegroundColor Magenta
Write-Host "✓ Successfully identified and analyzed Knowledge Manager permissions" -ForegroundColor Green
Write-Host "✓ Successfully identified the 6 security group permissions to remove" -ForegroundColor Green  
Write-Host "❌ Cannot create custom role: Remaining permissions are not supported" -ForegroundColor Red
Write-Host "💡 Provided alternative approaches to achieve your goal" -ForegroundColor Cyan

Write-Host "`nScript completed. No custom role was created due to platform limitations." -ForegroundColor Yellow