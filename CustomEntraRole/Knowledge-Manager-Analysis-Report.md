# Knowledge Manager Custom Role - Analysis Report

## Issue Summary
You wanted to create a custom role based on the Knowledge Manager role but without these security group permissions:
- `microsoft.directory/groups.security/basic/update`
- `microsoft.directory/groups.security/create`
- `microsoft.directory/groups.security/createAsOwner`
- `microsoft.directory/groups.security/delete`
- `microsoft.directory/groups.security/members/update`
- `microsoft.directory/groups.security/owners/update`

## Root Cause Analysis

### What We Discovered
1. **Knowledge Manager role contains 11 permissions total**
2. **6 permissions are the security group permissions you want to remove**
3. **5 remaining permissions are all `office365.*` permissions**
4. **Microsoft does NOT support `office365.*` permissions in custom directory roles**

### The Technical Limitation
```
❌ Error: "Action 'microsoft.office365.knowledge/...' is not supported for Custom Role creation"
```

Microsoft restricts custom directory roles to only `microsoft.directory.*` permissions. This is intentional - they want to maintain control over certain service-level permissions.

### Knowledge Manager Permissions Breakdown
```
✓ Requested to Remove (6 permissions):
  • microsoft.directory/groups.security/basic/update
  • microsoft.directory/groups.security/create
  • microsoft.directory/groups.security/createAsOwner
  • microsoft.directory/groups.security/delete
  • microsoft.directory/groups.security/members/update
  • microsoft.directory/groups.security/owners/update

❌ Cannot Use in Custom Roles (5 permissions):
  • microsoft.office365.knowledge/contentUnderstanding/analytics/allProperties/read
  • microsoft.office365.knowledge/knowledgeNetwork/topicVisibility/allProperties/allTasks
  • microsoft.office365.sharePoint/allEntities/allTasks
  • microsoft.office365.supportTickets/allEntities/allTasks
  • microsoft.office365.webPortal/allEntities/standard/read
```

## Recommended Solutions

### Option 1: Use Built-in Role + Access Controls (Recommended)
Keep the Knowledge Manager role but control group access through other means:

**Conditional Access Policies:**
- Create policies that restrict group management operations
- Target specific users/groups
- Require additional authentication for group operations

**Privileged Identity Management (PIM):**
- Make the Knowledge Manager role assignment time-limited
- Require approval for activation
- Add additional verification steps

**Azure AD Access Reviews:**
- Regular reviews of who has the Knowledge Manager role
- Automated removal of unused assignments

### Option 2: Alternative Built-in Roles
Instead of Knowledge Manager, use combinations of these roles:

```powershell
# SharePoint focused
"SharePoint Administrator"

# Viva/Knowledge focused  
"Viva Engage Administrator"
"Knowledge Administrator" (if available)

# Support focused
"Service Support Administrator"
"Helpdesk Administrator"
```

### Option 3: Hybrid Approach
- Use Azure RBAC for resource-level permissions
- Use minimal directory roles for identity operations
- Combine multiple smaller roles

## Implementation Scripts

### Script 1: List Knowledge Manager Role Assignments
```powershell
# See who currently has the Knowledge Manager role
$kmRole = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -eq "Knowledge Manager" }
$assignments = Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.RoleDefinitionId -eq $kmRole.Id }

$assignments | ForEach-Object {
    $principal = Get-MgDirectoryObject -DirectoryObjectId $_.PrincipalId
    Write-Host "User: $($principal.DisplayName) - Assignment: $($_.Id)"
}
```

### Script 2: Create Conditional Access Policy (Template)
```powershell
# Template for restricting group operations
# This would need to be customized for your environment
$policy = @{
    displayName = "Restrict Group Management for Knowledge Managers"
    state = "enabled"
    conditions = @{
        applications = @{
            includeApplications = @("797f4846-ba00-4fd7-ba43-dac1f8f63013") # Azure AD Graph
        }
        users = @{
            includeRoles = @($kmRole.Id)
        }
    }
    grantControls = @{
        operator = "OR"
        builtInControls = @("mfa")
    }
}
# Note: This is a template - actual implementation requires careful testing
```

## Alternative Analysis Tools

If you want to analyze other roles for custom role creation:

```powershell
# Check which permissions in a role are custom-role compatible
function Test-CustomRoleCompatibility {
    param([string]$RoleName)
    
    $role = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -eq $RoleName }
    $permissions = $role.RolePermissions[0].AllowedResourceActions
    
    $compatible = $permissions | Where-Object { $_ -like "microsoft.directory/*" }
    $incompatible = $permissions | Where-Object { $_ -like "microsoft.office365.*" }
    
    Write-Host "Role: $RoleName"
    Write-Host "Compatible: $($compatible.Count)"
    Write-Host "Incompatible: $($incompatible.Count)"
}

# Test various roles
Test-CustomRoleCompatibility "User Administrator"
Test-CustomRoleCompatibility "Groups Administrator" 
Test-CustomRoleCompatibility "SharePoint Administrator"
```

## Conclusion

The Knowledge Manager role is specifically designed as a built-in role with permissions that cannot be replicated in custom roles. This is a platform limitation, not a script error.

**Bottom Line:** You cannot create a custom version of the Knowledge Manager role without security group permissions because the remaining permissions are not supported in custom roles.

**Recommendation:** Use Option 1 (Built-in role + access controls) as it maintains the full Knowledge Manager functionality while controlling group access through policy.