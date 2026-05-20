# Hybrid Dependency + AI Analysis Functions
# Add these functions to script.py after the build_dependency_graph function

import os
import json
import time
import random
import concurrent.futures
import threading
from openai import AzureOpenAI
from backend.move_support_fetcher import fetch_move_support_data

# Import centralized configuration
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import AzureOpenAIConfig, PerformanceConfig

_NON_MOVABLE_MEMORY_CACHE = None
_NON_MOVABLE_CACHE_LOCK = threading.Lock()
_AI_RESULT_CACHE = {}
_AI_RESULT_CACHE_LOCK = threading.Lock()
_KNOWN_NON_MOVABLE_FALLBACK_TYPES = {
    "microsoft.network/bastionhosts",
}

def build_dependency_graph(resources):
    """Build a dependency graph showing relationships between resources."""
    graph = {}
    resource_map = {res.get("id", "").lower(): res for res in resources}
    
    for res in resources:
        res_id = res.get("id", "").lower()
        res_type = res.get("type", "").lower()
        dependencies = []
        
        # VM dependencies
        if res_type == "microsoft.compute/virtualmachines":
            # NICs
            for nic in res.get("properties", {}).get("networkProfile", {}).get("networkInterfaces", []):
                if nic.get("id"):
                    nic_id = nic["id"].lower()
                    dependencies.append(nic_id)
                    
                    # Resolve VNet via NIC
                    if nic_id in resource_map:
                        nic_res = resource_map[nic_id]
                        for ip_cfg in nic_res.get("properties", {}).get("ipConfigurations", []):
                            subnet_id = ip_cfg.get("properties", {}).get("subnet", {}).get("id")
                            if subnet_id and "/virtualnetworks/" in subnet_id.lower():
                                vnet_id = subnet_id.lower().split("/subnets/")[0]
                                if vnet_id not in dependencies:
                                    dependencies.append(vnet_id)
            # Disks
            os_disk_id = res.get("properties", {}).get("storageProfile", {}).get("osDisk", {}).get("managedDisk", {}).get("id")
            if os_disk_id:
                dependencies.append(os_disk_id.lower())
            for data_disk in res.get("properties", {}).get("storageProfile", {}).get("dataDisks", []):
                disk_id = data_disk.get("managedDisk", {}).get("id")
                if disk_id:
                    dependencies.append(disk_id.lower())
        
        # NIC dependencies
        elif res_type == "microsoft.network/networkinterfaces":
            # Public IPs, NSGs, Subnets
            for ip_cfg in res.get("properties", {}).get("ipConfigurations", []):
                pub_ip = ip_cfg.get("properties", {}).get("publicIPAddress", {}).get("id")
                if pub_ip:
                    dependencies.append(pub_ip.lower())
                subnet = ip_cfg.get("properties", {}).get("subnet", {}).get("id")
                if subnet:
                    dependencies.append(subnet.lower())
                    # Resolve VNet from Subnet
                    if "/virtualnetworks/" in subnet.lower():
                        vnet_id = subnet.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)

            nsg = res.get("properties", {}).get("networkSecurityGroup", {}).get("id")
            if nsg:
                dependencies.append(nsg.lower())
        
        # Public IP dependencies
        elif res_type == "microsoft.network/publicipaddresses":
            ip_cfg = res.get("properties", {}).get("ipConfiguration", {}).get("id")
            if ip_cfg:
                dependencies.append(ip_cfg.lower())
        
        # Load Balancer dependencies
        elif res_type == "microsoft.network/loadbalancers":
            for frontend in res.get("properties", {}).get("frontendIPConfigurations", []):
                pub_ip = frontend.get("properties", {}).get("publicIPAddress", {}).get("id")
                if pub_ip:
                    dependencies.append(pub_ip.lower())
                subnet = frontend.get("properties", {}).get("subnet", {}).get("id")
                if subnet:
                    dependencies.append(subnet.lower())
                    # Resolve VNet from Subnet
                    if "/virtualnetworks/" in subnet.lower():
                        vnet_id = subnet.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)
                subnet = frontend.get("properties", {}).get("subnet", {}).get("id")
                if subnet:
                    dependencies.append(subnet.lower())
        
        # VNet dependencies (peerings only - reverse dependencies tracked separately)
        elif res_type == "microsoft.network/virtualnetworks":
            for peering in res.get("properties", {}).get("virtualNetworkPeerings", []):
                remote_vnet = peering.get("properties", {}).get("remoteVirtualNetwork", {}).get("id")
                if remote_vnet:
                    dependencies.append(remote_vnet.lower())

        # AKS dependencies
        elif res_type == "microsoft.containerservice/managedclusters":
            # Check agent pool profiles for VNet subnet
            for profile in res.get("properties", {}).get("agentPoolProfiles", []):
                subnet_id = profile.get("vnetSubnetID")
                if subnet_id:
                    dependencies.append(subnet_id.lower())
                    # Resolve VNet from Subnet
                    if "/virtualnetworks/" in subnet_id.lower():
                        vnet_id = subnet_id.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)

        # App Service dependencies
        elif res_type == "microsoft.web/sites":
            # VNet Integration
            subnet_id = res.get("properties", {}).get("virtualNetworkSubnetId")
            if subnet_id:
                dependencies.append(subnet_id.lower())
                # Resolve VNet from Subnet
                if "/virtualnetworks/" in subnet_id.lower():
                    vnet_id = subnet_id.lower().split("/subnets/")[0]
                    if vnet_id not in dependencies:
                        dependencies.append(vnet_id)

        # SQL Managed Instance dependencies
        elif res_type == "microsoft.sql/managedinstances":
            subnet_id = res.get("properties", {}).get("subnetId")
            if subnet_id:
                dependencies.append(subnet_id.lower())
                # Resolve VNet from Subnet
                if "/virtualnetworks/" in subnet_id.lower():
                    vnet_id = subnet_id.lower().split("/subnets/")[0]
                    if vnet_id not in dependencies:
                        dependencies.append(vnet_id)
        
        # VMSS dependencies
        elif res_type == "microsoft.compute/virtualmachinescalesets":
            # Load Balancer
            for lb_ref in res.get("properties", {}).get("virtualMachineProfile", {}).get("networkProfile", {}).get("networkInterfaceConfigurations", []):
                for ip_cfg in lb_ref.get("properties", {}).get("ipConfigurations", []):
                    # Load Balancer Backend Pools
                    for pool in ip_cfg.get("properties", {}).get("loadBalancerBackendAddressPools", []):
                        lb_id = pool.get("id", "")
                        if lb_id and "/loadbalancers/" in lb_id.lower():
                            lb_id = lb_id.lower().split("/backendaddresspools/")[0]
                            if lb_id not in dependencies:
                                dependencies.append(lb_id)
                    # Subnet/VNet
                    subnet_id = ip_cfg.get("properties", {}).get("subnet", {}).get("id")
                    if subnet_id:
                        dependencies.append(subnet_id.lower())
                        if "/virtualnetworks/" in subnet_id.lower():
                            vnet_id = subnet_id.lower().split("/subnets/")[0]
                            if vnet_id not in dependencies:
                                dependencies.append(vnet_id)
                    # Public IP (if any)
                    pub_ip_cfg = ip_cfg.get("properties", {}).get("publicIPAddressConfiguration")
                    if pub_ip_cfg:
                        # VMSS creates public IPs dynamically, track the configuration
                        pass
        
        # Application Gateway dependencies
        elif res_type == "microsoft.network/applicationgateways":
            # Public IPs
            for frontend in res.get("properties", {}).get("frontendIPConfigurations", []):
                pub_ip = frontend.get("properties", {}).get("publicIPAddress", {}).get("id")
                if pub_ip:
                    dependencies.append(pub_ip.lower())
                subnet = frontend.get("properties", {}).get("subnet", {}).get("id")
                if subnet:
                    dependencies.append(subnet.lower())
                    if "/virtualnetworks/" in subnet.lower():
                        vnet_id = subnet.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)
            # Backend pools (VMs/VMSS)
            for pool in res.get("properties", {}).get("backendAddressPools", []):
                for addr in pool.get("properties", {}).get("backendAddresses", []):
                    ip_addr = addr.get("ipAddress")
                    # Note: Can't directly resolve IP to resource ID without additional lookup
        
        # VPN Gateway dependencies
        elif res_type == "microsoft.network/virtualnetworkgateways":
            # Public IPs
            for ip_cfg in res.get("properties", {}).get("ipConfigurations", []):
                pub_ip = ip_cfg.get("properties", {}).get("publicIPAddress", {}).get("id")
                if pub_ip:
                    dependencies.append(pub_ip.lower())
                subnet = ip_cfg.get("properties", {}).get("subnet", {}).get("id")
                if subnet:
                    dependencies.append(subnet.lower())
                    if "/virtualnetworks/" in subnet.lower():
                        vnet_id = subnet.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)
        
        # Private Endpoints
        elif res_type == "microsoft.network/privateendpoints":
            # Subnet
            subnet_id = res.get("properties", {}).get("subnet", {}).get("id")
            if subnet_id:
                dependencies.append(subnet_id.lower())
                if "/virtualnetworks/" in subnet_id.lower():
                    vnet_id = subnet_id.lower().split("/subnets/")[0]
                    if vnet_id not in dependencies:
                        dependencies.append(vnet_id)
            # Target resource (privateLinkServiceConnections)
            for conn in res.get("properties", {}).get("privateLinkServiceConnections", []):
                target_id = conn.get("properties", {}).get("privateLinkServiceId")
                if target_id:
                    dependencies.append(target_id.lower())
        
        # Network Security Groups (NSGs)
        elif res_type == "microsoft.network/networksecuritygroups":
            # NSGs don't have forward dependencies, but we track reverse (subnets/NICs using them)
            pass
        
        # Storage Accounts
        elif res_type == "microsoft.storage/storageaccounts":
            # Private Endpoints (tracked via reverse dependencies)
            # VNet Rules are in firewall config, not direct dependencies
            pass
        
        # Function Apps
        elif res_type == "microsoft.web/sites" and res.get("kind", "").startswith("functionapp"):
            # Storage Account (required for Function Apps)
            storage_conn = res.get("properties", {}).get("siteConfig", {}).get("appSettings", [])
            # Note: Storage account connection string is in app settings, hard to parse
            # VNet Integration
            subnet_id = res.get("properties", {}).get("virtualNetworkSubnetId")
            if subnet_id:
                dependencies.append(subnet_id.lower())
                if "/virtualnetworks/" in subnet_id.lower():
                    vnet_id = subnet_id.lower().split("/subnets/")[0]
                    if vnet_id not in dependencies:
                        dependencies.append(vnet_id)
            # App Service Plan
            server_farm_id = res.get("properties", {}).get("serverFarmId")
            if server_farm_id:
                dependencies.append(server_farm_id.lower())
        
        # Key Vault
        elif res_type == "microsoft.keyvault/vaults":
            # Private Endpoints (tracked via reverse dependencies)
            # VNet Rules
            for rule in res.get("properties", {}).get("networkAcls", {}).get("virtualNetworkRules", []):
                subnet_id = rule.get("id")
                if subnet_id:
                    dependencies.append(subnet_id.lower())
                    if "/virtualnetworks/" in subnet_id.lower():
                        vnet_id = subnet_id.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)
        
        # Azure SQL Database/Server
        elif res_type == "microsoft.sql/servers":
            # Private Endpoints (tracked via reverse dependencies)
            pass
        
        # Container Registry (ACR)
        elif res_type == "microsoft.containerregistry/registries":
            # Private Endpoints (tracked via reverse dependencies)
            # Geo-replications
            for replication in res.get("properties", {}).get("replications", []):
                # Replications are separate resources, tracked independently
                pass
        
        # Azure Data Factory
        elif res_type == "microsoft.datafactory/factories":
            # Managed VNet
            managed_vnet = res.get("properties", {}).get("managedVirtualNetwork")
            if managed_vnet:
                # Managed VNet is internal to ADF
                pass
        
        # API Management
        elif res_type == "microsoft.apimanagement/service":
            # VNet Integration
            vnet_config = res.get("properties", {}).get("virtualNetworkConfiguration")
            if vnet_config:
                subnet_id = vnet_config.get("subnetResourceId")
                if subnet_id:
                    dependencies.append(subnet_id.lower())
                    if "/virtualnetworks/" in subnet_id.lower():
                        vnet_id = subnet_id.lower().split("/subnets/")[0]
                        if vnet_id not in dependencies:
                            dependencies.append(vnet_id)
            # Public IP (for external VNet mode)
            pub_ip_ids = res.get("properties", {}).get("publicIPAddresses", [])
            for pub_ip_ref in pub_ip_ids:
                # Handle case where it's a string (IP address or ID) or dict
                if isinstance(pub_ip_ref, dict):
                    if pub_ip_ref.get("id"):
                        dependencies.append(pub_ip_ref["id"].lower())
                elif isinstance(pub_ip_ref, str):
                    # Only add if it looks like a resource ID
                    if "/publicipaddresses/" in pub_ip_ref.lower():
                        dependencies.append(pub_ip_ref.lower())
        
        # ---------------------------------------------------------
        # GENERIC FALLBACK: Scan properties for ANY Resource IDs
        # ---------------------------------------------------------
        # This ensures we catch dependencies for complex or unknown resource types 
        # that aren't explicitly handled above (e.g., Stream Analytics, Logic Apps).
        def extract_ids_recursive(obj):
            ids = set()
            if isinstance(obj, dict):
                for v in obj.values():
                    ids.update(extract_ids_recursive(v))
            elif isinstance(obj, list):
                for i in obj:
                    ids.update(extract_ids_recursive(i))
            elif isinstance(obj, str):
                # Check for Azure Resource ID format
                # Must start with /subscriptions/ and contain /providers/ (or be a known core type like resourceGroups)
                if obj.lower().startswith("/subscriptions/") and (
                    "/providers/" in obj.lower() or "/resourcegroups/" in obj.lower()
                ):
                    ids.add(obj.lower())
            return ids

        # Scan the entire resource properties structure
        found_generic_ids = extract_ids_recursive(res.get("properties", {}))
        
        # Add found IDs if not already in dependencies and not self
        for fid in found_generic_ids:
            # Avoid self-reference
            if fid != res_id and fid not in dependencies:
                # Optional: Filter out IDs that are just child resources of the current resource
                if not fid.startswith(res_id + "/"):
                    dependencies.append(fid)

        graph[res_id] = {
            "resource": res,
            "dependencies": dependencies
        }
    
    # Build reverse dependencies for ALL resources (not just specific types)
    # This ensures we catch all dependency relationships regardless of resource type
    for res_id in graph.keys():
        reverse_deps = []
        # Find all resources that reference this resource
        for other_id, other_data in graph.items():
            if res_id in other_data.get("dependencies", []):
                reverse_deps.append(other_id)
        
        if reverse_deps:
            graph[res_id]["reverse_dependencies"] = reverse_deps
    
    return graph

def extract_dependency_facts(resource, dependency_graph, non_movable_list=None):
    """
    Extract concrete facts about a resource's dependencies and configuration.
    Returns a structured summary for AI analysis.
    """
    res_id = resource.get("id", "").lower()
    res_type = resource.get("type", "").lower()
    res_name = resource.get("name", "")
    
    facts = {
        "resource_name": res_name,
        "resource_type": res_type,
        "location": resource.get("location", ""),
        "sku": resource.get("sku") or {},
        "dependencies": [],
        "blocking_factors": []
    }
    
    # Get dependency info
    dep_info = dependency_graph.get(res_id, {})
    dependencies = dep_info.get("dependencies", [])
    
    # Analyze based on resource type
    if res_type == "microsoft.compute/virtualmachines":
        # Check for Public IPs through NICs
        for dep_id in dependencies:
            dep_res = dependency_graph.get(dep_id, {}).get("resource", {})
            dep_type = dep_res.get("type", "").lower()
            
            if dep_type == "microsoft.network/networkinterfaces":
                nic_name = dep_res.get("name", "")
                # Check if NIC has public IP
                for ip_cfg in dep_res.get("properties", {}).get("ipConfigurations", []):
                    pub_ip_id = ip_cfg.get("properties", {}).get("publicIPAddress", {}).get("id", "").lower()
                    if pub_ip_id:
                        pub_ip_res = dependency_graph.get(pub_ip_id, {}).get("resource", {})
                        if pub_ip_res:
                            pub_ip_sku = pub_ip_res.get("sku", {}).get("name", "Unknown")
                            pub_ip_name = pub_ip_res.get("name", "")
                            facts["blocking_factors"].append({
                                "type": "Public IP Attached",
                                "details": f"Public IP '{pub_ip_name}' (SKU: {pub_ip_sku}) attached via NIC '{nic_name}'"
                            })
        
        # Check for Availability Set
        avset = resource.get("properties", {}).get("availabilitySet", {}).get("id")
        if avset:
            facts["blocking_factors"].append({
                "type": "Availability Set",
                "details": f"VM is part of an Availability Set"
            })
        
        # NEW: Check if VNet dependency is blocked by non-movable resources
        if non_movable_list:
            for dep_id in dependencies:
                dep_res = dependency_graph.get(dep_id, {}).get("resource", {})
                dep_type = dep_res.get("type", "").lower()
                
                if dep_type == "microsoft.network/virtualnetworks":
                    vnet_name = dep_res.get("name", "")
                    # Check if this VNet has non-movable reverse dependencies
                    vnet_reverse_deps = dependency_graph.get(dep_id, {}).get("reverse_dependencies", [])
                    
                    non_movable_vnet_deps = []
                    for vnet_dep_id in vnet_reverse_deps:
                        vnet_dep_res = dependency_graph.get(vnet_dep_id, {}).get("resource", {})
                        vnet_dep_type = vnet_dep_res.get("type", "").lower()
                        vnet_dep_name = vnet_dep_res.get("name", "")
                        
                        if is_resource_non_movable(vnet_dep_type, non_movable_list):
                            non_movable_vnet_deps.append(f"{vnet_dep_name} ({vnet_dep_type.split('/')[-1]})")
                    
                    if non_movable_vnet_deps:
                        facts["blocking_factors"].append({
                            "type": "VNet Dependency Blocked (HARD BLOCKER)",
                            "details": f"VM depends on VNet '{vnet_name}' which has non-movable dependencies: {', '.join(non_movable_vnet_deps)}. Moving VNet would break these resources with no rollback capability. VM cannot move with VNet."
                        })
                        break
    
    elif res_type == "microsoft.network/virtualnetworks":
        # Check for VNet peerings
        peerings = resource.get("properties", {}).get("virtualNetworkPeerings", [])
        if peerings:
            # Extract peer VNet names
            peer_vnet_names = []
            for peering in peerings:
                # Extract peer VNet name from the remote VNet ID
                remote_vnet_id = peering.get("properties", {}).get("remoteVirtualNetwork", {}).get("id", "")
                if remote_vnet_id:
                    # ID format: /subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Network/virtualNetworks/{name}
                    peer_name = remote_vnet_id.split("/")[-1] if "/" in remote_vnet_id else "Unknown"
                    peer_vnet_names.append(peer_name)
            
            if peer_vnet_names:
                peer_list = " and ".join(peer_vnet_names) if len(peer_vnet_names) <= 2 else ", ".join(peer_vnet_names[:-1]) + ", and " + peer_vnet_names[-1]
                facts["blocking_factors"].append({
                    "type": "VNet Peering",
                    "details": f"VNet peering with {peer_list} - peering must be removed before move"
                })
            else:
                facts["blocking_factors"].append({
                    "type": "VNet Peering",
                    "details": f"Has {len(peerings)} VNet peering(s) - must be deleted before move"
                })
        
        # Add reverse dependencies (resources that depend on this VNet)
        reverse_deps = dep_info.get("reverse_dependencies", [])
        if reverse_deps:
            dep_count = len(reverse_deps)
            dep_types = {}
            non_movable_deps = []
            
            # Check if any reverse dependency is non-movable
            for dep_id in reverse_deps:
                dep_res = dependency_graph.get(dep_id, {}).get("resource", {})
                dep_type = dep_res.get("type", "")
                dep_type_short = dep_type.split("/")[-1] if dep_type else "Unknown"
                dep_name = dep_res.get("name", "")
                dep_types[dep_type_short] = dep_types.get(dep_type_short, 0) + 1
                
                # Check if this dependency is non-movable
                if non_movable_list and dep_type and is_resource_non_movable(dep_type.lower(), non_movable_list):
                    non_movable_deps.append(f"{dep_name} ({dep_type_short})")
            
            type_summary = ", ".join([f"{count} {dtype}(s)" for dtype, count in dep_types.items()])
            facts["dependencies"] = reverse_deps
            
            # If there are non-movable dependencies, this is a HARD BLOCKER
            if non_movable_deps:
                facts["blocking_factors"].append({
                    "type": "Non-Movable Dependency (HARD BLOCKER)",
                    "details": f"VNet has {len(non_movable_deps)} non-movable resource(s): {', '.join(non_movable_deps)}. Moving VNet would immediately break these resources with no rollback capability. Production-safe migration requires parallel infrastructure approach."
                })
            else:
                # All dependencies are movable
                facts["blocking_factors"].append({
                    "type": "Dependent Resources",
                    "details": f"{dep_count} resources depend on this VNet: {type_summary}. All must be moved together or dependencies updated."
                })
    
    elif res_type == "microsoft.network/publicipaddresses":
        # Check SKU and what it's attached to
        sku = resource.get("sku", {}).get("name", "Basic")
        allocation = resource.get("properties", {}).get("publicIPAllocationMethod", "Dynamic")
        facts["sku"] = {"name": sku}
        
        # Check if attached to VMSS (cannot move) or VM/NIC (can move after detach)
        ip_config_id = resource.get("properties", {}).get("ipConfiguration", {}).get("id", "")
        attached_to_vmss = False
        attached_to_vm = False
        attached_resource_name = ""
        
        if ip_config_id:
            # Check if it's part of a VMSS
            if "/virtualmachinescalesets/" in ip_config_id.lower():
                attached_to_vmss = True
                # Extract VMSS name
                parts = ip_config_id.split("/")
                for i, part in enumerate(parts):
                    if part.lower() == "virtualmachinescalesets" and i + 1 < len(parts):
                        attached_resource_name = parts[i + 1]
                        break
            else:
                # It's attached to a regular NIC/VM
                attached_to_vm = True
                # Try to find the NIC or VM name
                for dep_id in dependencies:
                    dep_res = dependency_graph.get(dep_id, {}).get("resource", {})
                    if dep_res.get("id", "").lower() == ip_config_id.lower():
                        attached_resource_name = dep_res.get("name", "")
                        break
        
        if sku == "Standard":
            if attached_to_vmss:
                facts["blocking_factors"].append({
                    "type": "Standard Public IP on VMSS",
                    "details": f"Standard SKU Public IP attached to VMSS '{attached_resource_name}'. VMSS with Standard Public IP cannot be moved."
                })
            elif attached_to_vm:
                facts["blocking_factors"].append({
                    "type": "Standard Public IP Attached",
                    "details": f"Standard SKU Public IP attached to '{attached_resource_name}'. Can be moved after detaching from NIC/VM."
                })
            else:
                facts["blocking_factors"].append({
                    "type": "Standard Public IP",
                    "details": f"Standard SKU Public IP (not currently attached). Can be moved."
                })
        else:
            facts["blocking_factors"].append({
                "type": "Public IP Configuration",
                "details": f"SKU: {sku}, Allocation: {allocation}"
            })
    
    # NEW: Check for NIC blocked by VNet with non-movable dependencies
    elif res_type == "microsoft.network/networkinterfaces":
        if non_movable_list:
            # Check if NIC is connected to a VNet that has non-movable dependencies
            for dep_id in dependencies:
                dep_res = dependency_graph.get(dep_id, {}).get("resource", {})
                dep_type = dep_res.get("type", "").lower()
                
                if dep_type == "microsoft.network/virtualnetworks":
                    vnet_name = dep_res.get("name", "")
                    vnet_reverse_deps = dependency_graph.get(dep_id, {}).get("reverse_dependencies", [])
                    
                    non_movable_vnet_deps = []
                    for vnet_dep_id in vnet_reverse_deps:
                        vnet_dep_res = dependency_graph.get(vnet_dep_id, {}).get("resource", {})
                        vnet_dep_type = vnet_dep_res.get("type", "").lower()
                        vnet_dep_name = vnet_dep_res.get("name", "")
                        
                        if is_resource_non_movable(vnet_dep_type, non_movable_list):
                            non_movable_vnet_deps.append(f"{vnet_dep_name} ({vnet_dep_type.split('/')[-1]})")
                    
                    if non_movable_vnet_deps:
                        # Check if NIC is attached to a VM
                        reverse_deps = dep_info.get("reverse_dependencies", [])
                        vm_name = "Unknown VM"
                        for rev_dep_id in reverse_deps:
                            rev_dep_res = dependency_graph.get(rev_dep_id, {}).get("resource", {})
                            rev_dep_type = rev_dep_res.get("type", "").lower()
                            if rev_dep_type == "microsoft.compute/virtualmachines":
                                vm_name = rev_dep_res.get("name", "Unknown VM")
                                break
                        
                        facts["blocking_factors"].append({
                            "type": "VNet Dependency Blocked (HARD BLOCKER)",
                            "details": f"NIC is attached to VM '{vm_name}' and connected to VNet '{vnet_name}' which has non-movable dependencies: {', '.join(non_movable_vnet_deps)}. NIC cannot move with VM/VNet."
                        })
                        break
    
    return facts

def analyze_with_ai(resource_facts, knowledge_base):
    """
    Use AI to interpret dependency facts against Microsoft documentation.
    """
    res_type = resource_facts["resource_type"]
    res_name = resource_facts["resource_name"]
    
    # Get provider namespace
    provider_parts = res_type.split('/')
    if len(provider_parts) < 2:
        return None
    
    provider_namespace = provider_parts[0]
    resource_type_name = provider_parts[1] if len(provider_parts) > 1 else ""
    
    # Special namespace mapping: Front Door Standard/Premium resources should use Microsoft.Cdn documentation
    # See: https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources
    frontdoor_types = ["frontdoors", "frontdoorwebapplicationfirewallpolicies", "rulesengines"]
    if provider_namespace.lower() == "microsoft.network" and resource_type_name.lower() in frontdoor_types:
        # Also look up Microsoft.Cdn for Front Door Standard/Premium documentation
        lookup_namespaces = [provider_namespace, "microsoft.cdn"]
    else:
        lookup_namespaces = [provider_namespace]
    
    # Look up knowledge base (try multiple namespaces for Front Door)
    kb_entry = None
    for namespace in lookup_namespaces:
        for k, v in knowledge_base.items():
            if k.lower() == namespace.lower():
                kb_entry = v
                break
        if kb_entry:
            break
    
    if not kb_entry:
        print(f"⚠️ No knowledge base entry found for provider: {provider_namespace}")
        return None
        
    if not kb_entry.get('rules_text'):
        print(f"⚠️ No rules text found for provider: {provider_namespace}")
        return None
    
    # Reuse identical AI decisions across resources with same move characteristics.
    blocking_signature = tuple(
        sorted(
            (bf.get("type", ""), bf.get("details", ""))
            for bf in resource_facts.get("blocking_factors", [])
        )
    )
    ai_cache_key = (
        resource_facts.get("resource_type", ""),
        resource_facts.get("location", ""),
        resource_facts.get("sku", {}).get("name", ""),
        blocking_signature,
        kb_entry.get("detail_url", kb_entry.get("summary_url", "")),
    )
    with _AI_RESULT_CACHE_LOCK:
        cached_result = _AI_RESULT_CACHE.get(ai_cache_key)
    if cached_result is not None:
        return dict(cached_result)

    # Use Azure OpenAI configuration from config
    endpoint = AzureOpenAIConfig.ENDPOINT
    api_key = AzureOpenAIConfig.API_KEY
    deployment_name = AzureOpenAIConfig.DEPLOYMENT_NAME
    
    if not all([endpoint, api_key, deployment_name]):
        return None
    
    try:
        client = AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version=AzureOpenAIConfig.API_VERSION
        )
        
        # Build focused prompt with dependency facts
        blocking_summary = "\n".join([
            f"- {bf['type']}: {bf['details']}" 
            for bf in resource_facts.get("blocking_factors", [])
        ]) or "None detected"
        
        all_deps_summary = ", ".join(resource_facts.get("all_dependencies", [])) or "None"
        
        prompt = f"""You are an Azure migration expert. Analyze if this resource can be moved **to a new subscription or tenant** based on CONCRETE FACTS from dependency analysis and official Microsoft documentation.

RESOURCE INFORMATION:
- Name: {res_name}
- Type: {res_type}
- Location: {resource_facts.get('location', 'Unknown')}
- SKU: {resource_facts.get('sku', {}).get('name', 'N/A')}

DEPENDENCY ANALYSIS FINDINGS:
**Blocking Factors (Must be resolved):**
{blocking_summary}

**All Connected Resources (For Context Only - NOT necessarily blocking):**
{all_deps_summary}

MICROSOFT MOVE LIMITATIONS (focus on specific restrictions mentioned):
{kb_entry['rules_text'][:5000]}

CRITICAL INSTRUCTIONS:
1. **EXACT RESOURCE TYPE MATCH**: Find the EXACT resource type '{res_type}' in the documentation. Read any clarification notes (e.g., "This row is for X tier only"). Do NOT apply rules from similar resource types.
2. **ASSESSMENT GOAL**: Determine if this resource can be moved to a **different Subscription** or **Tenant**.
2. **IGNORE Region Move** limitations. This is NOT a region move.
3. **RESOLVABLE BLOCKERS**: If a resource is blocked by a configuration that can be changed (e.g., "Private DNS Zone Link", "Standard SKU Public IP attached"), mark it as **"can_move": true** but specify the **"action_required"**.
4. **HARD BLOCKERS**: Only mark "can_move": false if the resource type ITSELF cannot move or if the blocker is impossible to resolve (e.g., "Attached to VMSS").
5. **Private DNS Zone Links**: If a VNet has Private DNS Zone links, it CAN move, but the links must be removed first. Mark as **"can_move": true** with **"action_required": "Remove Private DNS Zone links before moving"**.
6. **Public IPs - IMPORTANT**: 
   - If the dependency findings show the Public IP is **ATTACHED** (has reverse dependencies), mark **"action_required": "Detach Public IP before moving"**.
   - If the dependency findings show the Public IP is **NOT ATTACHED** (no reverse dependencies), mark **"action_required": null**.
7. **Application Gateway Public IPs**: If attached to Application Gateway, mark as **"can_move": true** with **"action_required": "Detach from Application Gateway before moving"**.
8. **Front Door WAF Policies**: The resource type 'frontdoorwebapplicationfirewallpolicies' CAN be moved. Do NOT confuse with 'frontdoors' (Classic tier). WAF policies are separate resources.
9. **DEFAULT ASSUMPTION**: If the documentation does NOT explicitly state that a resource type CANNOT be moved across subscriptions, assume it CAN be moved. Only mark "can_move": false if there is explicit documentation stating the resource type does not support cross-subscription moves. The absence of restrictions means the resource is movable.
10. **Traffic Manager Profiles**: The resource type 'microsoft.network/trafficmanagerprofiles' CAN be moved across subscriptions. If endpoints are also being moved, they may need to be deleted and re-added after the move due to resource ID changes.
11. **VNet Peering (IMPORTANT - Wording Clarity)**:
    - If blocking_factors mentions "VNet peering" or "virtualnetworkpeerings":
      * Mark "can_move": true
      * Extract the ACTUAL peer VNet name(s) from the blocking_factors text
      * Example: If blocking_factors says "VNet peering to vnet-hub-prd-01", use "vnet-hub-prd-01"
      * Example: If blocking_factors says "Peered with vnet-spoke-stg and vnet-spoke-prd", use both names
      * Mark "action_required": "Remove VNet peering with {{actual-vnet-name}} before moving (both VNets can move together)"
      * Where {{actual-vnet-name}} is the REAL VNet name you extracted from blocking_factors
      * DO NOT use placeholder text like [peer-vnet-name] or <peer-vnet-name>
      * DO NOT say "VNet must be deleted/recreated" - this is confusing
      * The PEERING must be removed, not the VNet itself
      * Both peered VNets can move to the target subscription, then peering can be recreated
12. **HARD BLOCKERS - Non-Movable Dependencies (CRITICAL)**:
    - If blocking_factors contains "HARD BLOCKER" (e.g., "Non-Movable Dependency (HARD BLOCKER)", "VNet Dependency Blocked (HARD BLOCKER)"):
      * Mark "can_move": false
      * Mark "action_required": null
      * Reason: Must reference the specific non-movable dependency mentioned (e.g., "VNet has AKS cluster dependency")
      * Applicable Rule: "Resource has non-movable dependencies - moving would break production with no rollback capability. Use parallel infrastructure approach."
    - These are NOT resolvable by detaching or configuration changes
    - Moving would immediately break dependent resources with no ability to rollback
    - Production-safe migration requires parallel infrastructure approach (build new in target, migrate workloads, cutover)

ANALYSIS REQUIRED:
Based on the dependency findings and the documentation, can this resource be moved to another subscription/tenant?

OUTPUT (JSON only, no markdown):
{{
    "can_move": true or false,
    "action_required": "Step-by-step action to resolve the blocker (e.g., 'Detach Public IP', 'Remove Private DNS Link'). Leave null if no action needed.",
    "reason": "Specific reason referencing the actual dependency finding",
    "applicable_rule": "Exact quote from documentation if blocking, or 'No limitations found for cross-subscription move'"
}}"""
        
        # Retry logic for Rate Limit (429) errors - use config values
        max_retries = AzureOpenAIConfig.MAX_RETRIES
        retry_delay = AzureOpenAIConfig.INITIAL_RETRY_DELAY
        
        for attempt in range(max_retries):
            try:
                response = client.chat.completions.create(
                    model=deployment_name,
                    messages=[
                        {"role": "system", "content": "You are an Azure migration expert. Analyze based on concrete dependency facts and documentation."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=AzureOpenAIConfig.TEMPERATURE,
                    response_format={"type": "json_object"}
                )
                break # Success, exit loop
                
            except Exception as e:
                # Check for Rate Limit Error (429)
                error_msg = str(e)
                if "429" in error_msg or "RateLimit" in error_msg:
                    if attempt < max_retries - 1:
                        wait_time = min(retry_delay + random.uniform(0, 1), AzureOpenAIConfig.MAX_RETRY_DELAY)
                        print(f"⏳ Rate limit hit for {res_name}. Retrying in {wait_time:.1f}s... (Attempt {attempt+1}/{max_retries})")
                        time.sleep(wait_time)
                        retry_delay *= AzureOpenAIConfig.BACKOFF_MULTIPLIER
                        continue
                
                # If not 429 or retries exhausted, re-raise
                raise e
        
        result = json.loads(response.choices[0].message.content)
        result['source_url'] = kb_entry.get('detail_url', kb_entry.get('summary_url', ''))
        with _AI_RESULT_CACHE_LOCK:
            _AI_RESULT_CACHE[ai_cache_key] = dict(result)
        
        return result
        
    except Exception as e:
        import traceback
        print(f"⚠️ AI analysis failed for {res_name}: {e}")
        traceback.print_exc()
        return None

def load_non_movable_resources():
    """
    Load the list of non-movable resources with automatic cache management.
    
    This function now uses the dynamic cache manager which:
    - Checks if cache is stale (older than configured days)
    - Auto-updates from Microsoft docs if stale
    - Falls back to cached data if update fails
    """
    global _NON_MOVABLE_MEMORY_CACHE

    if _NON_MOVABLE_MEMORY_CACHE is not None:
        return _NON_MOVABLE_MEMORY_CACHE

    try:
        # Import the dynamic cache manager
        from backend.non_movable_resources_manager import load_non_movable_resources as load_dynamic
        
        print("   🔄 Using dynamic cache manager...")
        sys.stdout.flush()
        
        # Use the dynamic loader with auto-update
        result = load_dynamic()
        
        print(f"   ✅ Loaded {len(result)} namespaces from cache manager")
        sys.stdout.flush()
        
        with _NON_MOVABLE_CACHE_LOCK:
            _NON_MOVABLE_MEMORY_CACHE = result
        return _NON_MOVABLE_MEMORY_CACHE
    
    except ImportError as ie:
        # Fallback to old static method if manager is not available
        print(f"⚠️ Dynamic cache manager not available: {ie}, using static load")
        sys.stdout.flush()
        try:
            json_path = os.path.join(os.path.dirname(__file__), 'non_movable_resources.json')
            with open(json_path, 'r') as f:
                data = json.load(f)
            
            # Handle both old format (array) and new format (object with metadata)
            if isinstance(data, list):
                with _NON_MOVABLE_CACHE_LOCK:
                    _NON_MOVABLE_MEMORY_CACHE = data
                return _NON_MOVABLE_MEMORY_CACHE
            elif isinstance(data, dict) and 'non_movable_resources' in data:
                with _NON_MOVABLE_CACHE_LOCK:
                    _NON_MOVABLE_MEMORY_CACHE = data['non_movable_resources']
                return _NON_MOVABLE_MEMORY_CACHE
            else:
                print(f"⚠️ Unexpected JSON format in non_movable_resources.json")
                return []
        except Exception as e:
            print(f"⚠️ Failed to load non-movable resources list: {e}")
            return []
    
    except Exception as e:
        print(f"⚠️ Error in dynamic cache manager: {e}")
        sys.stdout.flush()
        import traceback
        traceback.print_exc()
        return []

def is_resource_non_movable(res_type, non_movable_list):
    """Check if a resource type is in the non-movable list."""
    res_type_lower = res_type.lower()
    
    for item in non_movable_list:
        namespace = item.get("namespace", "").lower()
        if namespace in res_type_lower:
            for type_name in item.get("types", []):
                # Construct full type string to check
                # Handle cases where type might be just the sub-type or full path
                full_check = f"{namespace}/{type_name}".lower()
                if full_check == res_type_lower or f"/{type_name}".lower() in res_type_lower:
                    return True
    return False

def analyze_move_feasibility(resource, dependency_graph, move_support_data, non_movable_list=None):
    """
    Hybrid approach: Extract dependency facts, then use AI to interpret against MS docs.
    Includes FAST PATH for known non-movable resources.
    """
    res_id = resource.get("id", "").lower()
    res_type = resource.get("type", "").lower()
    res_name = resource.get("name", "")
    
    # Extract dependencies from the dependency graph
    dep_info = dependency_graph.get(res_id, {})
    dependency_ids = dep_info.get("dependencies", [])
    reverse_dependency_ids = dep_info.get("reverse_dependencies", [])
    
    # Get dependency resource names (forward dependencies)
    dependency_names = []
    for dep_id in dependency_ids:
        dep_resource = dependency_graph.get(dep_id, {}).get("resource", {})
        dep_name = dep_resource.get("name", "")
        if dep_name:
            dependency_names.append(dep_name)
    
    # Get reverse dependency resource names (resources that depend on this one)
    reverse_dependency_names = []
    for dep_id in reverse_dependency_ids:
        dep_resource = dependency_graph.get(dep_id, {}).get("resource", {})
        dep_name = dep_resource.get("name", "")
        dep_type = dep_resource.get("type", "").split("/")[-1] if dep_resource.get("type") else ""
        if dep_name:
            reverse_dependency_names.append(f"{dep_name} ({dep_type})")
    
    # Combine both for display
    all_dependencies = dependency_names.copy()
    if reverse_dependency_names:
        all_dependencies.extend([f"← {name}" for name in reverse_dependency_names])
    
    # 🚀 FAST PATH: Check if resource is universally non-movable
    if non_movable_list is None:
        non_movable_list = load_non_movable_resources()
    if non_movable_list and is_resource_non_movable(res_type, non_movable_list):
        print(f"   ⚡ Fast Path: {res_name} ({res_type}) is known non-movable.")
        return {
            "status": "Cannot Move",
            "action": "Recreate in target subscription",
            "dependencies": all_dependencies,
            "notes": "According to Microsoft documentation, this resource type does not support cross-subscription move operations.",
            "source_url": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
        }
    if not non_movable_list and res_type in _KNOWN_NON_MOVABLE_FALLBACK_TYPES:
        return {
            "status": "Cannot Move",
            "action": "Recreate in target subscription",
            "dependencies": all_dependencies,
            "notes": "Resource type is treated as non-movable by fallback rule when dynamic cache is unavailable.",
            "source_url": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
        }

    # Step 1: Extract concrete dependency facts
    resource_facts = extract_dependency_facts(resource, dependency_graph, non_movable_list)
    resource_facts["all_dependencies"] = all_dependencies

    # Fast-path #3: for low-risk resources with no detected blockers, skip AI.
    # This significantly improves runtime for small/clean environments.
    if PerformanceConfig.SKIP_AI_FOR_LOW_RISK and not resource_facts.get("blocking_factors"):
        return {
            "status": "Can Move",
            "action": "No action required",
            "dependencies": all_dependencies,
            "notes": "No blocking factors detected from dependency analysis.",
            "source_url": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
        }
    
    # 🚀 FAST PATH #2: Check if resource has HARD BLOCKER (non-movable dependencies)
    # This skips AI analysis for resources that are clearly blocked, significantly improving performance
    has_hard_blocker = any(
        "HARD BLOCKER" in bf.get("type", "") 
        for bf in resource_facts.get("blocking_factors", [])
    )
    
    if has_hard_blocker:
        # Extract the blocking details for the user
        hard_blocker_details = next(
            (bf.get("details", "") for bf in resource_facts.get("blocking_factors", []) 
             if "HARD BLOCKER" in bf.get("type", "")),
            "Resource has non-movable dependencies"
        )
        
        print(f"   ⚡ Fast Path #2: {res_name} blocked by non-movable dependency (skipping AI)")
        return {
            "status": "Cannot Move",
            "action": "Use parallel infrastructure approach - recreate in target subscription",
            "dependencies": all_dependencies,
            "notes": f"{hard_blocker_details}\n\nProduction-safe migration requires parallel infrastructure: build new resources in target subscription, migrate workloads/data, validate, then cutover. Keep source running for rollback capability.",
            "source_url": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
        }
    
    # Step 2: Use AI to interpret facts against documentation
    ai_result = analyze_with_ai(resource_facts, move_support_data)
    
    if ai_result:
        action_req = ai_result.get('action_required')
        # Check if action_required is meaningful (not null, empty, or starts with "None")
        has_action = action_req and action_req.strip() and not action_req.strip().lower().startswith('none')
        
        if has_action:
            status = "Action Required"
            action = action_req
            # Append reason to notes for context
            notes = f"{ai_result.get('reason', '')}\n\nRule: {ai_result.get('applicable_rule', '')}"
        else:
            status = "Cannot Move" if not ai_result.get('can_move') else "Can Move"
            action = ai_result.get('reason', 'No action required')
            notes = ai_result.get('applicable_rule', '')

        return {
            "status": status,
            "action": action,
            "dependencies": all_dependencies,  # Now includes both forward and reverse
            "notes": notes,
            "source_url": ai_result.get('source_url', '')
        }
    
    # Fallback if AI fails - mark for manual review
    # Note: Non-movable resources are already caught by the fast path earlier using non_movable_resources.json
    
    return {
        "status": "Review Required",
        "action": "Manual verification needed",
        "dependencies": all_dependencies,
        "notes": "Resource type not found in knowledge base. Please review manually.",
        "source_url": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
    }

def identify_move_groups(resources, dependency_graph):
    """
    Group resources that must be moved together based on dependencies.
    Returns a dict mapping resource_id -> group_id
    Only assigns groups to resources that have dependencies.
    """
    # Initialize Union-Find structure
    parent_map = {}
    for res in resources:
        if isinstance(res, dict) and 'id' in res:
            parent_map[res['id'].lower()] = res['id'].lower()
    
    def find(i):
        if i not in parent_map: return i
        if parent_map[i] == i:
            return i
        parent_map[i] = find(parent_map[i])
        return parent_map[i]
    
    def union(i, j):
        root_i = find(i)
        root_j = find(j)
        if root_i != root_j:
            # If one is missing from map (e.g. external dependency), don't break
            if root_i in parent_map and root_j in parent_map:
                parent_map[root_i] = root_j
    
    # Track which resources have dependencies
    resources_with_deps = set()
    
    # Iterate through dependencies and union connected resources
    for res_id, info in dependency_graph.items():
        res_id = res_id.lower()
        if res_id not in parent_map: continue
        
        # Union with Forward dependencies
        for dep_id in info.get('dependencies', []):
            dep_id = dep_id.lower()
            if dep_id in parent_map:
                union(res_id, dep_id)
                resources_with_deps.add(res_id)
                resources_with_deps.add(dep_id)
                
    # Assign Group IDs only to resources with dependencies
    groups = {}
    group_id_map = {}
    group_counter = 1
    
    # Sort keys for deterministic group numbering
    sorted_ids = sorted(parent_map.keys())
    
    for res_id in sorted_ids:
        # Only assign group if resource has dependencies
        if res_id in resources_with_deps:
            root = find(res_id)
            if root not in groups:
                groups[root] = f"MG-{group_counter:03d}"
                group_counter += 1
            group_id_map[res_id] = groups[root]
        else:
            # Standalone resources get no group
            group_id_map[res_id] = ""
        
    return group_id_map


def generate_resource_level_migration_plan(resources, sub_map):
    """Generate a detailed migration plan for each resource with dependency conflict detection."""
    print("🔍 Building dependency graph...")
    dependency_graph = build_dependency_graph(resources)
    
    # Fetch dynamic move support data
    print("🌐 Fetching dynamic move support notes...")
    move_support_data = fetch_move_support_data()
    
    print("📊 Analyzing move feasibility for each resource...")
    
    # Load non-movable resources list ONCE (not per-resource)
    print("   📋 Loading non-movable resources list...")
    non_movable_list = load_non_movable_resources()
    
    # PASS 1: Analyze each resource independently (Parallelized)
    resource_analysis = {}
    total_resources = len(resources)
    completed_count = 0
    
    def analyze_single_resource(res):
        res_id = res.get("id", "").lower()
        analysis = analyze_move_feasibility(res, dependency_graph, move_support_data, non_movable_list)
        return res_id, {
            "resource": res,
            "analysis": analysis
        }

    # Use ThreadPoolExecutor for parallel processing
    # Dynamic worker count based on environment size for optimal performance
    if total_resources < 20:
        max_workers = 5  # Small environment - conservative
    elif total_resources < 100:
        max_workers = 10  # Medium environment
    elif total_resources < 300:
        max_workers = 15  # Large environment
    else:
        max_workers = 20  # Very large environment - maximum parallelism
    
    print(f"   🚀 Using {max_workers} parallel workers for {total_resources} resources")
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_res = {executor.submit(analyze_single_resource, res): res for res in resources}
        
        for future in concurrent.futures.as_completed(future_to_res):
            try:
                res_id, result = future.result()
                resource_analysis[res_id] = result
                
                completed_count += 1
                
                # Progress reporting every 5% or every 10 resources (whichever is more frequent)
                progress_interval = max(1, min(10, total_resources // 20))
                if completed_count % progress_interval == 0 or completed_count == total_resources:
                    percentage = int((completed_count / total_resources) * 100)
                    print(f"   📊 Progress: {completed_count}/{total_resources} ({percentage}%) resources analyzed...")
                    
            except Exception as e:
                print(f"   ⚠️ Error analyzing resource: {e}")
    
    print(f"   ✅ Completed analysis of {total_resources} resources.")
    
    print("🔗 Detecting dependency conflicts...")
    
    # PASS 2: Iterative dependency conflict detection
    # Keep checking until no new conflicts are found (fixed-point iteration)
    max_iterations = 10  # Prevent infinite loops
    iteration = 0
    conflicts_detected_total = 0
    
    while iteration < max_iterations:
        iteration += 1
        conflicts_this_iteration = 0
        
        for res_id, data in resource_analysis.items():
            res = data["resource"]
            analysis = data["analysis"]
            
            # Skip if already marked as "Cannot Move"
            if analysis["status"] == "Cannot Move":
                continue
            
            dep_info = dependency_graph.get(res_id, {})
            
            # Check FORWARD dependencies (resources this one depends on)
            forward_deps = dep_info.get("dependencies", [])
            for dep_id in forward_deps:
                if dep_id in resource_analysis:
                    dep_analysis = resource_analysis[dep_id]["analysis"]
                    if dep_analysis["status"] == "Cannot Move":
                        dep_resource = resource_analysis[dep_id]["resource"]
                        dep_name = dep_resource.get("name", "Unknown")
                        dep_type = dep_resource.get("type", "").split("/")[-1]
                        
                        # Check if this dependency is already mentioned to avoid duplicate notes
                        if dep_name in analysis["notes"]:
                            break
                        
                        # Mark as Dependent - the resource itself CAN move if dependency is resolved
                        old_status = analysis["status"]
                        analysis["status"] = "Dependent - Action Required"
                        analysis["action"] = f"Resolve dependency: {dep_name} ({dep_type}) must be deleted/recreated"
                        analysis["notes"] = f"This resource CAN move, but depends on '{dep_name}' ({dep_type}) which cannot be moved. You must delete/recreate the dependency in the target. {analysis['notes']}"
                        
                        # Only increment if this is a NEW conflict (status changed or first time processing this dep)
                        if old_status != "Dependent - Action Required":
                            conflicts_this_iteration += 1
                            conflicts_detected_total += 1
                            print(f"   [Debug] Forward conflict: {res.get('name')} blocked by {dep_name}")
                        
                        break  # One blocker is enough
                    
                    elif dep_analysis["status"] == "Action Required":
                        dep_resource = resource_analysis[dep_id]["resource"]
                        dep_name = dep_resource.get("name", "Unknown")
                        
                        # Mark this resource as needing review (Yellow)
                        if analysis["status"] != "Dependent - Action Required":
                            analysis["status"] = "Dependent - Action Required"
                            analysis["action"] = f"Dependency requires action: {dep_name}"
                            analysis["notes"] = f"Depends on '{dep_name}' which requires action: {dep_analysis['action']}. {analysis['notes']}"
            
            # Skip reverse dependency check if already blocked by forward dependency
            if analysis["status"] == "Cannot Move":
                continue
            
            # Check REVERSE dependencies (resources that depend on this one)
            reverse_deps = dep_info.get("reverse_dependencies", [])
            blocking_reverse_deps = []
            action_required_reverse_deps = []
            
            for rev_dep_id in reverse_deps:
                if rev_dep_id in resource_analysis:
                    rev_dep_analysis = resource_analysis[rev_dep_id]["analysis"]
                    rev_dep_resource = resource_analysis[rev_dep_id]["resource"]
                    rev_dep_name = rev_dep_resource.get("name", "Unknown")
                    rev_dep_type = rev_dep_resource.get("type", "").split("/")[-1]
                    
                    if rev_dep_analysis["status"] == "Cannot Move":
                        blocking_reverse_deps.append(f"{rev_dep_name} ({rev_dep_type})")
                    elif rev_dep_analysis["status"] in ["Action Required", "Can be Moved/Action Required"]:
                        action_required_reverse_deps.append(f"{rev_dep_name} ({rev_dep_type})")
            
            if blocking_reverse_deps:
                # If already processed, check if we need to add more blocking deps
                if analysis["status"] in ["Action Required", "Dependent - Action Required"]:
                    # Extract resource names from existing action and new blockers to avoid duplicates
                    existing_action = analysis.get("action", "")
                    new_blockers = []
                    
                    for blocker in blocking_reverse_deps:
                        # Extract just the resource name (before the parenthesis)
                        blocker_name = blocker.split(" (")[0] if " (" in blocker else blocker
                        # Check if this specific resource name is already mentioned
                        if blocker_name not in existing_action:
                            new_blockers.append(blocker)
                    
                    if new_blockers:
                        # Append only truly new blockers to existing action
                        new_blockers_str = ", ".join(new_blockers)
                        analysis["action"] = f"{existing_action}; Also resolve: {new_blockers_str}"
                    continue
                    
                # Special case: Public IPs can be moved once detached (except VMSS)
                res_type_lower = res.get("type", "").lower()
                if res_type_lower == "microsoft.network/publicipaddresses":
                    # Check if any blocking reverse dep is a VMSS (truly cannot move)
                    has_vmss_dependency = any(
                        "virtualmachinescalesets" in resource_analysis[rev_dep_id]["resource"].get("type", "").lower()
                        for rev_dep_id in reverse_deps
                        if rev_dep_id in resource_analysis and resource_analysis[rev_dep_id]["analysis"]["status"] == "Cannot Move"
                    )
                    
                    if not has_vmss_dependency:
                        # Public IP can be moved, just needs to be detached first
                        analysis["status"] = "Action Required"
                        blockers_str = ", ".join(blocking_reverse_deps)
                        # Extract resource type from blocker name (e.g., "cityfurnish-prod-awg (applicationgateways)")
                        attached_resource_type = blocking_reverse_deps[0].split("(")[-1].rstrip(")") if blocking_reverse_deps else "attached resource"
                        analysis["action"] = f"Detach from {attached_resource_type} before moving"
                        analysis["notes"] = f"This Public IP is attached to {blockers_str}. Detach it before moving. {analysis['notes']}"
                        continue  # Skip marking as Cannot Move
                
                # Special case: Managed Disks can be moved once detached from VM
                elif res_type_lower == "microsoft.compute/disks":
                    # Disks can always be detached and moved independently
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Detach from VM before moving"
                    analysis["notes"] = f"This disk is attached to {blockers_str}. Detach it from the VM before moving. {analysis['notes']}"
                    continue  # Skip marking as Cannot Move
                
                # Special case: NSGs can always be detached from subnets/NICs and moved
                elif res_type_lower == "microsoft.network/networksecuritygroups":
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Detach from subnet/NIC before moving"
                    analysis["notes"] = f"This NSG is associated with {blockers_str}. Detach it before moving. {analysis['notes']}"
                    continue
                
                # Special case: Route Tables can be disassociated from subnets
                elif res_type_lower == "microsoft.network/routetables":
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Disassociate from subnet before moving"
                    analysis["notes"] = f"This Route Table is associated with {blockers_str}. Disassociate it before moving. {analysis['notes']}"
                    continue
                
                # Special case: NAT Gateways can be disassociated from subnets
                elif res_type_lower == "microsoft.network/natgateways":
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Disassociate from subnet before moving"
                    analysis["notes"] = f"This NAT Gateway is associated with {blockers_str}. Disassociate it before moving. {analysis['notes']}"
                    continue
                
                # Special case: Application Security Groups can be removed from NIC configs
                elif res_type_lower == "microsoft.network/applicationsecuritygroups":
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Remove from NIC configurations before moving"
                    analysis["notes"] = f"This ASG is used by {blockers_str}. Remove it from NIC configs before moving. {analysis['notes']}"
                    continue
                
                # Special case: Load Balancers can have backend pools emptied
                elif res_type_lower == "microsoft.network/loadbalancers":
                    analysis["status"] = "Action Required"
                    blockers_str = ", ".join(blocking_reverse_deps)
                    analysis["action"] = f"Remove backend pool associations before moving"
                    analysis["notes"] = f"This Load Balancer has backends: {blockers_str}. Remove associations before moving. {analysis['notes']}"
                    continue
                
                # Mark this resource as blocked by reverse dependencies, but allow move if dependencies are resolved
                old_status = analysis["status"]
                analysis["status"] = "Dependent - Action Required"
                blockers_str = ", ".join(blocking_reverse_deps)
                analysis["action"] = f"Resolve dependent resource(s): {blockers_str}"
                analysis["notes"] = f"This resource is blocked by {blockers_str} which cannot be moved. You must delete or detach these dependent resources before moving this resource. {analysis['notes']}"
                
                # Only increment if this is a NEW conflict
                if old_status != "Dependent - Action Required":
                    conflicts_this_iteration += 1
                    conflicts_detected_total += 1
                    print(f"   [Debug] Reverse conflict: {res.get('name')} blocked by {blockers_str}")
            
            elif action_required_reverse_deps and analysis["status"] == "Can Move":
                # If dependent resources need action, mark this as needing review
                analysis["status"] = "Dependent - Action Required"
                deps_str = ", ".join(action_required_reverse_deps)
                analysis["action"] = f"Dependent resource(s) require action: {deps_str}"
                analysis["notes"] = f"This resource is used by {deps_str} which require action before moving. {analysis['notes']}"
                conflicts_this_iteration += 1
                conflicts_detected_total += 1
        
        # If no new conflicts found, we've reached a fixed point
        if conflicts_this_iteration == 0:
            print(f"   ✅ Conflict detection converged after {iteration} iteration(s)")
            break
        else:
            print(f"   🔄 Iteration {iteration}: Found {conflicts_this_iteration} new conflict(s)")
    
    if conflicts_detected_total > 0:
        print(f"   ⚠️ Total conflicts detected: {conflicts_detected_total}")
    else:
        print(f"   ✅ No dependency conflicts detected")
    
    # 🔗 Identify Move Groups (Connected Components)
    print("   🔗 Identifying move groups...")
    move_groups = identify_move_groups(resources, dependency_graph)
    
    # Build final migration plan
    migration_plan = []
    full_analysis_list = []  # Ungrouped list for dashboard
    
    # Group fast-path non-movable resources
    non_movable_groups = {}
    
    for res_id, data in resource_analysis.items():
        res = data["resource"]
        analysis = data["analysis"]
        
        # Create the standard entry
        entry = {
            "Resource Name": res.get("name", ""),
            "Resource Type": res.get("type", ""),
            "Subscription": sub_map.get(res.get("subscriptionId", ""), "Unknown"),
            "subscriptionId": res.get("subscriptionId", ""),  # Required for grouping
            "Resource Group": res.get("resourceGroup", ""),
            "Location": res.get("location", ""),
            "Move Status": analysis["status"],
            "Move Group": move_groups.get(res_id, "MG-Unknown"),  # New Field
            "Action Required": analysis["action"],
            "Dependencies": ", ".join(analysis["dependencies"]) if analysis["dependencies"] else "None",
            "Notes": analysis["notes"],
            "Source URL": analysis.get("source_url", "")
        }
        
        # Always add to full list (for dashboard)
        full_analysis_list.append(entry)
        
        # Check if this was a fast-path non-movable resource
        # We identify this by the specific note we added earlier
        is_fast_path = "According to Microsoft documentation, this resource type does not support cross-subscription move operations" in analysis.get("notes", "")
        
        if is_fast_path:
            res_type = res.get("type", "Unknown Type")
            sub_id = res.get("subscriptionId", "Unknown")
            sub_name = sub_map.get(sub_id, "Unknown")
            
            key = (res_type, sub_name)
            if key not in non_movable_groups:
                non_movable_groups[key] = {
                    "count": 0,
                    "example_name": res.get("name", ""),
                    "notes": analysis["notes"],
                    "source_url": analysis.get("source_url", "")
                }
            non_movable_groups[key]["count"] += 1
        else:
            # Add normal resources individually to the plan
            migration_plan.append(entry)

    # Add grouped non-movable resources to the plan
    for (res_type, sub_name), info in non_movable_groups.items():
        friendly_type = res_type.split("/")[-1]
        display_name = f"{friendly_type} ({info['count']} resources)"
        
        migration_plan.append({
            "Resource Name": display_name,
            "Resource Type": res_type,
            "Subscription": sub_name,
            "Resource Group": "Multiple",
            "Location": "Various",
            "Move Status": "Cannot Move",
            "Move Group": "",  # Grouped resources don't have move groups
            "Action Required": "Recreate in target subscription",
            "Dependencies": "N/A",
            "Notes": f"{info['notes']} (Grouped {info['count']} items, e.g., '{info['example_name']}')",
            "Source URL": info['source_url']
        })

    return migration_plan, full_analysis_list
