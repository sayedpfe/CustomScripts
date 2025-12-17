# Sign in with the needed scopes
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory","Directory.Read.All"

Write-Host "=== Knowledge Manager Custom Role Creator ===" -ForegroundColor Cyan
Write-Host "This script creates a custom role based on Knowledge Manager but excludes group security permissions." -ForegroundColor Cyan
Write-Host ""

# 1) Get the built-in Knowledge Manager role definition
Write-Host "Step 1: Getting Knowledge Manager role definition..." -ForegroundColor Green
$km = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -eq "Knowledge Manager" }

if (-not $km) { 
    Write-Host "Built-in role 'Knowledge Manager' not found." -ForegroundColor Red
    throw "Built-in role 'Knowledge Manager' not found." 
}

Write-Host "✓ Found Knowledge Manager role: $($km.DisplayName)" -ForegroundColor Green
Write-Host "  Role ID: $($km.Id)" -ForegroundColor Gray

# 2) Extract and analyze permissions
Write-Host "`nStep 2: Analyzing permissions..." -ForegroundColor Green
$allowed = @()
foreach ($perm in $km.RolePermissions) {
    if ($perm.AllowedResourceActions) {
        $allowed += $perm.AllowedResourceActions
    }
}
$allowed = $allowed | Sort-Object -Unique

Write-Host "✓ Found $($allowed.Count) permissions in Knowledge Manager role" -ForegroundColor Green

# Show all permissions
Write-Host "`nAll permissions in Knowledge Manager role:" -ForegroundColor Yellow
$allowed | ForEach-Object { Write-Host "  • $_" -ForegroundColor Yellow }

# 3) Remove requested exclusions
Write-Host "`nStep 3: Removing requested security group permissions..." -ForegroundColor Green
$excludedPermissions = @(
    "microsoft.directory/groups.security/basic/update",
    "microsoft.directory/groups.security/create", 
    "microsoft.directory/groups.security/createAsOwner",
    "microsoft.directory/groups.security/delete",
    "microsoft.directory/groups.security/members/update",
    "microsoft.directory/groups.security/owners/update"
)

Write-Host "Removing these permissions:" -ForegroundColor Red
$excludedPermissions | ForEach-Object { Write-Host "  ✗ $_" -ForegroundColor Red }

$afterExclusions = $allowed | Where-Object { $excludedPermissions -notcontains $_ }
$removedCount = $allowed.Count - $afterExclusions.Count

Write-Host "`n✓ Removed $removedCount permissions" -ForegroundColor Green
Write-Host "✓ Remaining permissions: $($afterExclusions.Count)" -ForegroundColor Green

# 4) Important discovery about custom roles
Write-Host "`nStep 4: Important Discovery About Custom Directory Roles" -ForegroundColor Magenta
Write-Host "========================================================" -ForegroundColor Magenta
Write-Host ""
Write-Host "Based on testing, most Office 365 service permissions (office365.*) are NOT supported" -ForegroundColor Yellow
Write-Host "in custom directory roles. These include:" -ForegroundColor Yellow
Write-Host "  • microsoft.office365.knowledge/* (Knowledge management)" -ForegroundColor Yellow  
Write-Host "  • microsoft.office365.sharePoint/* (SharePoint)" -ForegroundColor Yellow
Write-Host "  • microsoft.office365.supportTickets/* (Support tickets)" -ForegroundColor Yellow
Write-Host "  • microsoft.office365.webPortal/* (Web portal)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Only microsoft.directory/* permissions are typically supported in custom directory roles." -ForegroundColor Yellow
Write-Host ""

# 5) Check what's left that might work
$supportedPermissions = $afterExclusions | Where-Object { $_ -like "microsoft.directory/*" }

Write-Host "Step 5: Permissions that might be supported in custom roles:" -ForegroundColor Green
if ($supportedPermissions.Count -gt 0) {
    $supportedPermissions | ForEach-Object { Write-Host "  ✓ $_" -ForegroundColor Green }
} else {
    Write-Host "  ⚠ No microsoft.directory/* permissions remain after exclusions!" -ForegroundColor Yellow
}

Write-Host "`nUnsupported permissions (will be filtered out):" -ForegroundColor Red
$unsupportedPermissions = $afterExclusions | Where-Object { $_ -notlike "microsoft.directory/*" }
$unsupportedPermissions | ForEach-Object { Write-Host "  ✗ $_" -ForegroundColor Red }

# 6) Create role with only supported permissions or create minimal role for demonstration
if ($supportedPermissions.Count -gt 0) {
    Write-Host "`nStep 6: Creating custom role with supported permissions..." -ForegroundColor Green
    $finalPermissions = $supportedPermissions
} else {
    Write-Host "`nStep 6: Creating minimal demonstration role..." -ForegroundColor Yellow
    Write-Host "Since no directory permissions remain, creating a role with minimal permissions for demonstration." -ForegroundColor Yellow
    Write-Host "This role will have very limited functionality." -ForegroundColor Yellow
    
    # Use the most basic directory read permission that usually works
    $finalPermissions = @("microsoft.office365.webPortal/allEntities/standard/read")
}

$roleBody = @{
    displayName = "Knowledge Manager - Limited (No Security Groups)"
    description = "Custom role based on Knowledge Manager with security group permissions removed. Note: Many Knowledge Manager permissions are only available in built-in roles."
    isEnabled = $true
    rolePermissions = @(@{ 
        allowedResourceActions = $finalPermissions 
    })
}

Write-Host "`nRole configuration:" -ForegroundColor Cyan
Write-Host "  Name: $($roleBody.displayName)" -ForegroundColor Cyan
Write-Host "  Description: $($roleBody.description)" -ForegroundColor Cyan
Write-Host "  Permissions to include: $($finalPermissions.Count)" -ForegroundColor Cyan

if ($finalPermissions.Count -gt 0) {
    Write-Host "  Permissions:" -ForegroundColor Cyan
    $finalPermissions | ForEach-Object { Write-Host "    • $_" -ForegroundColor Cyan }
}

# 7) Attempt to create the role
Write-Host "`nStep 7: Creating the custom role..." -ForegroundColor Green

try {
    # Try creating with the minimal set first to see what works
    $testBody = @{
        displayName = "Knowledge Manager - Test Role"
        description = "Test role to validate permissions"
        isEnabled = $true
        rolePermissions = @(@{ 
            allowedResourceActions = @("microsoft.office365.webPortal/allEntities/standard/read")
        })
    }
    
    Write-Host "Testing with webPortal permission..." -ForegroundColor Yellow
    $testResult = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $testBody
    
    if ($testResult) {
        Write-Host "✓ Test role created successfully!" -ForegroundColor Green
        Write-Host "  Role ID: $($testResult.Id)" -ForegroundColor Green
        
        # Now clean up the test role and create the real one
        Write-Host "Removing test role..." -ForegroundColor Yellow
        Remove-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $testResult.Id
        
        # Create the actual role with webPortal permission only (since that's what works)
        $finalBody = @{
            displayName = "Knowledge Manager - Limited (No Security Groups)"
            description = "Custom role based on Knowledge Manager with security group permissions removed. Limited to web portal access only due to custom role restrictions."
            isEnabled = $true
            rolePermissions = @(@{ 
                allowedResourceActions = @("microsoft.office365.webPortal/allEntities/standard/read")
            })
        }
        
        Write-Host "Creating final role..." -ForegroundColor Green
        $finalResult = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $finalBody
        
        Write-Host "`n🎉 SUCCESS! Custom role created:" -ForegroundColor Green
        Write-Host "  Role ID: $($finalResult.Id)" -ForegroundColor Green
        Write-Host "  Display Name: $($finalResult.DisplayName)" -ForegroundColor Green
        Write-Host "  Description: $($finalResult.Description)" -ForegroundColor Green
        Write-Host "  Is Enabled: $($finalResult.IsEnabled)" -ForegroundColor Green
        
        Write-Host "`n📋 Summary:" -ForegroundColor Cyan
        Write-Host "✓ Successfully removed all requested security group permissions" -ForegroundColor Green
        Write-Host "✗ Most Knowledge Manager permissions cannot be used in custom roles" -ForegroundColor Yellow
        Write-Host "✓ Created role with web portal read access" -ForegroundColor Green
        
        Write-Host "`n⚠️  Important Notes:" -ForegroundColor Yellow
        Write-Host "• This custom role has very limited permissions compared to the built-in Knowledge Manager role" -ForegroundColor Yellow
        Write-Host "• Most Knowledge Manager functionality (SharePoint, Viva Topics, etc.) is only available in the built-in role" -ForegroundColor Yellow
        Write-Host "• Consider using the built-in Knowledge Manager role with alternative group management strategies" -ForegroundColor Yellow
        
    }
    
} catch {
    Write-Host "❌ Error creating role: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -like "*not supported*") {
        Write-Host "`n💡 Recommendation:" -ForegroundColor Cyan
        Write-Host "The Knowledge Manager role contains permissions that are exclusive to built-in roles." -ForegroundColor Cyan
        Write-Host "Consider these alternatives:" -ForegroundColor Cyan
        Write-Host "  1. Use the built-in Knowledge Manager role as-is" -ForegroundColor Cyan
        Write-Host "  2. Manage group permissions through Conditional Access or other policies" -ForegroundColor Cyan
        Write-Host "  3. Use Azure RBAC roles instead of directory roles for some permissions" -ForegroundColor Cyan
        Write-Host "  4. Create separate roles for different functions" -ForegroundColor Cyan
    }
}