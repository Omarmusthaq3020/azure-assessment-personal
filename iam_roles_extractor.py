"""
IAM Roles Extractor for Azure Assessment Tool
==============================================
Extracts role assignments with inheritance from subscription, resource group, and resource levels.
"""

import subprocess
import json
from collections import defaultdict
import threading

_ROLE_ASSIGNMENTS_CACHE = {}
_ROLE_ASSIGNMENTS_CACHE_LOCK = threading.Lock()

def _fetch_role_assignments(subscription_ids=None, env=None):
    """Fetch role assignments globally or scoped per subscription."""
    cache_key = tuple(sorted(subscription_ids)) if subscription_ids else ("__all__",)
    with _ROLE_ASSIGNMENTS_CACHE_LOCK:
        cached = _ROLE_ASSIGNMENTS_CACHE.get(cache_key)
    if cached is not None:
        return list(cached)

    role_assignments = []

    if subscription_ids:
        for sub_id in subscription_ids:
            cmd = [
                "az", "role", "assignment", "list",
                "--all",
                "--subscription", sub_id,
                "--output", "json"
            ]
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True,
                env=env
            )
            role_assignments.extend(json.loads(result.stdout))
    else:
        cmd = [
            "az", "role", "assignment", "list",
            "--all",
            "--output", "json"
        ]
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
            env=env
        )
        role_assignments = json.loads(result.stdout)

    with _ROLE_ASSIGNMENTS_CACHE_LOCK:
        _ROLE_ASSIGNMENTS_CACHE[cache_key] = list(role_assignments)
    return role_assignments

def extract_iam_roles(subscription_ids=None, env=None):
    """
    Extract all IAM role assignments across subscriptions with inheritance.
    
    Returns:
        dict: Mapping of resource_id -> list of role assignments (including inherited)
    """
    print("🔐 Extracting IAM role assignments...")
    
    try:
        role_assignments = _fetch_role_assignments(subscription_ids=subscription_ids, env=env)
        print(f"   ✅ Found {len(role_assignments)} role assignments")
        
        # Organize by scope (resource ID)
        roles_by_resource = defaultdict(list)
        subscription_roles = defaultdict(list)
        resource_group_roles = defaultdict(list)
        
        for assignment in role_assignments:
            scope = assignment.get("scope", "").lower()
            role_name = assignment.get("roleDefinitionName", "Unknown")
            principal_name = assignment.get("principalName", "Unknown")
            principal_type = assignment.get("principalType", "Unknown")
            
            assignment_info = {
                "role": role_name,
                "principal": principal_name,
                "principal_type": principal_type,
                "scope": scope,
                "scope_type": get_scope_type(scope)
            }
            
            # Categorize by scope type
            if "/subscriptions/" in scope and "/resourcegroups/" not in scope:
                # Subscription-level role
                subscription_id = extract_subscription_id(scope)
                subscription_roles[subscription_id].append(assignment_info)
            elif "/resourcegroups/" in scope and "/providers/" not in scope:
                # Resource group-level role
                rg_id = scope
                resource_group_roles[rg_id].append(assignment_info)
            else:
                # Resource-level role
                roles_by_resource[scope].append(assignment_info)
        
        # Build inheritance map
        roles_with_inheritance = build_inheritance_map(
            roles_by_resource,
            subscription_roles,
            resource_group_roles
        )
        
        print(f"   ✅ Processed role assignments with inheritance")
        return roles_with_inheritance
        
    except subprocess.CalledProcessError as e:
        print(f"   ⚠️ Failed to extract IAM roles: {e.stderr}")
        return {}
    except Exception as e:
        print(f"   ⚠️ Error extracting IAM roles: {e}")
        return {}


def get_scope_type(scope):
    """Determine the type of scope (Subscription, Resource Group, or Resource)"""
    if not scope:
        return "Unknown"
    
    scope_lower = scope.lower()
    
    if "/providers/" in scope_lower:
        return "Resource"
    elif "/resourcegroups/" in scope_lower:
        return "Resource Group"
    elif "/subscriptions/" in scope_lower:
        return "Subscription"
    else:
        return "Unknown"


def extract_subscription_id(scope):
    """Extract subscription ID from scope"""
    parts = scope.split("/")
    try:
        sub_index = parts.index("subscriptions")
        return parts[sub_index + 1].lower()
    except (ValueError, IndexError):
        return None


def extract_resource_group_id(scope):
    """Extract resource group ID from scope"""
    parts = scope.split("/")
    try:
        sub_index = parts.index("subscriptions")
        rg_index = parts.index("resourceGroups")
        subscription_id = parts[sub_index + 1]
        rg_name = parts[rg_index + 1]
        return f"/subscriptions/{subscription_id}/resourceGroups/{rg_name}".lower()
    except (ValueError, IndexError):
        return None


def build_inheritance_map(resource_roles, subscription_roles, resource_group_roles):
    """
    Build a map of resources to their roles including inherited roles.
    
    Inheritance hierarchy:
    1. Subscription-level roles (inherited by all resources in subscription)
    2. Resource Group-level roles (inherited by all resources in RG)
    3. Resource-level roles (direct assignments)
    """
    roles_with_inheritance = {}
    
    for resource_id, direct_roles in resource_roles.items():
        all_roles = []
        
        # Get subscription ID and RG ID for this resource
        subscription_id = extract_subscription_id(resource_id)
        rg_id = extract_resource_group_id(resource_id)
        
        # 1. Add inherited subscription-level roles
        if subscription_id and subscription_id in subscription_roles:
            for role in subscription_roles[subscription_id]:
                inherited_role = role.copy()
                inherited_role["inherited_from"] = "Subscription"
                all_roles.append(inherited_role)
        
        # 2. Add inherited resource group-level roles
        if rg_id and rg_id in resource_group_roles:
            for role in resource_group_roles[rg_id]:
                inherited_role = role.copy()
                inherited_role["inherited_from"] = "Resource Group"
                all_roles.append(inherited_role)
        
        # 3. Add direct resource-level roles
        for role in direct_roles:
            direct_role = role.copy()
            direct_role["inherited_from"] = "Direct"
            all_roles.append(direct_role)
        
        roles_with_inheritance[resource_id] = all_roles
    
    return roles_with_inheritance


def format_roles_for_excel(roles_list):
    """
    Format role assignments for Excel display.
    
    Args:
        roles_list: List of role assignment dictionaries
        
    Returns:
        str: Formatted string for Excel cell
    """
    if not roles_list:
        return "No roles assigned"
    
    formatted_roles = []
    for role in roles_list:
        inheritance = role.get("inherited_from", "Direct")
        role_name = role.get("role", "Unknown")
        principal = role.get("principal", "Unknown")
        principal_type = role.get("principal_type", "Unknown")
        
        if inheritance == "Direct":
            formatted_roles.append(f"🔹 {role_name} → {principal} ({principal_type})")
        elif inheritance == "Subscription":
            formatted_roles.append(f"📁 {role_name} → {principal} ({principal_type}) [Inherited from Subscription]")
        elif inheritance == "Resource Group":
            formatted_roles.append(f"📂 {role_name} → {principal} ({principal_type}) [Inherited from RG]")
    
    return "\n".join(formatted_roles)


def enrich_resources_with_iam(resources, roles_map):
    """
    Enrich resources with IAM role information.
    
    Args:
        resources: List of resource dictionaries
        roles_map: Dictionary mapping resource_id -> list of roles
    """
    print("🔐 Enriching resources with IAM role assignments...")
    
    enriched_count = 0
    for resource in resources:
        resource_id = resource.get("id", "").lower()
        
        if resource_id in roles_map:
            roles = roles_map[resource_id]
            resource["IAM Roles"] = format_roles_for_excel(roles)
            resource["IAM Role Count"] = len(roles)
            
            # Count direct vs inherited
            direct_count = sum(1 for r in roles if r.get("inherited_from") == "Direct")
            inherited_count = len(roles) - direct_count
            
            resource["IAM Direct Roles"] = direct_count
            resource["IAM Inherited Roles"] = inherited_count
            enriched_count += 1
        else:
            resource["IAM Roles"] = "No roles assigned"
            resource["IAM Role Count"] = 0
            resource["IAM Direct Roles"] = 0
            resource["IAM Inherited Roles"] = 0
    
    print(f"   ✅ Enriched {enriched_count} resources with IAM data")


def extract_all_role_assignments(subscription_ids=None, env=None):
    """
    Extract all IAM role assignments for creating a dedicated IAM sheet.
    Returns a flat list suitable for Excel export.
    """
    print("🔐 Extracting all IAM role assignments for dedicated sheet...")
    
    try:
        role_assignments = _fetch_role_assignments(subscription_ids=subscription_ids, env=env)
        processed_assignments = []
        
        for assignment in role_assignments:
            scope = assignment.get("scope", "")
            scope_type = get_scope_type(scope)
            subscription_id = extract_subscription_id(scope)
            
            # Extract resource details from scope
            resource_name = "N/A"
            resource_type = "N/A"
            resource_group = "N/A"
            
            if "/providers/" in scope:
                parts = scope.split("/")
                try:
                    resource_name = parts[-1]
                    provider_index = parts.index("providers")
                    resource_type = f"{parts[provider_index + 1]}/{parts[provider_index + 2]}"
                    rg_index = parts.index("resourceGroups")
                    resource_group = parts[rg_index + 1]
                except (ValueError, IndexError):
                    pass
            elif "/resourceGroups/" in scope:
                parts = scope.split("/")
                try:
                    rg_index = parts.index("resourceGroups")
                    resource_group = parts[rg_index + 1]
                    resource_name = resource_group
                    resource_type = "Resource Group"
                except (ValueError, IndexError):
                    pass
            elif "/subscriptions/" in scope:
                resource_name = subscription_id or "Unknown"
                resource_type = "Subscription"
            
            processed_assignments.append({
                "Scope Type": scope_type,
                "Subscription ID": subscription_id or "Unknown",
                "Resource Group": resource_group,
                "Resource Type": resource_type,
                "Resource Name": resource_name,
                "Role": assignment.get("roleDefinitionName", "Unknown"),
                "Principal Name": assignment.get("principalName", "Unknown"),
                "Principal Type": assignment.get("principalType", "Unknown"),
                "Principal ID": assignment.get("principalId", "Unknown"),
                "Scope": scope
            })
        
        print(f"   ✅ Processed {len(processed_assignments)} role assignments")
        return processed_assignments
        
    except subprocess.CalledProcessError as e:
        print(f"   ⚠️ Failed to extract IAM roles: {e.stderr}")
        return []
    except Exception as e:
        print(f"   ⚠️ Error: {e}")
        return []


if __name__ == "__main__":
    # Test the extractor
    roles = extract_iam_roles()
    print(f"\n📊 Summary:")
    print(f"   Total resources with roles: {len(roles)}")
    
    for resource_id, role_list in list(roles.items())[:5]:  # Show first 5
        print(f"\n   Resource: {resource_id}")
        print(f"   Roles: {format_roles_for_excel(role_list)}")

