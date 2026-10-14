import json
import argparse
import subprocess
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import re
import os
import os
from collections import defaultdict, Counter
import sys
import tempfile
import concurrent.futures
import uuid
import time

# Performance tracking
START_TIME = time.time()

# Add backend directory to path to import fetcher
sys.path.append(os.path.join(os.path.dirname(__file__), 'backend'))
from move_support_fetcher import fetch_move_support_data
from openai import AzureOpenAI
from backend.improved_dashboard import create_improved_dashboard

# Import centralized configuration
from config import (
    AppConfig,
    AzureResourceGraphConfig,
    AzureQuotaConfig,
    ExcelExportConfig,
    ResourceTypeFriendlyNames,
    AKSDetectionConfig,
    PerformanceConfig
)

# Import hybrid analysis functions  
from backend.hybrid_analysis_functions import (
    extract_dependency_facts,
    analyze_with_ai,
    analyze_move_feasibility,
    generate_resource_level_migration_plan
)

# Import IAM and Spot extractors
from iam_roles_extractor import extract_iam_roles, enrich_resources_with_iam, extract_all_role_assignments
from vm_spot_extractor import extract_spot_vm_details

def load_friendly_type_map():
    if not os.path.exists('providers_raw.json'):
        print("âŒ 'providers_raw.json' not found.")
        return {}

    with open('providers_raw.json') as f:
        raw = json.load(f)

    providers = raw.get("value", [])
    friendly_map = {}

    for provider in providers:
        ns = provider.get("namespace")
        for rt in provider.get("resourceTypes", []):
            resource_type = rt.get("resourceType")
            display_name = rt.get("properties", {}).get("displayName")
            full_type = f"{ns}/{resource_type}".lower()

            if not display_name:
                raw_name = resource_type.split("/")[-1]
                display_name = camel_to_title(raw_name.replace("-", " ").replace("_", " "))
            friendly_map[full_type] = display_name
    
    # Apply friendly name overrides from config
    friendly_map.update(ResourceTypeFriendlyNames.OVERRIDES)
    return friendly_map

def camel_to_title(s):
    return re.sub(r'(?<!^)(?=[A-Z])', ' ', s).title()

def convert_nsg_json_to_text(nsg_json):
    """
    Convert Azure NSG security rules from JSON format into a readable text format.
    Priority, Access, Name, Direction, Protocol, SourceAddr, SourcePort, DstAddr, DstPort.
    """
    if not nsg_json or 'properties' not in nsg_json:
        return "No properties found"
    
    props = nsg_json.get('properties', {})
    rules = props.get('securityRules', [])
    
    if not rules:
        return "No custom security rules defined."
        
    rule_texts = []
    
    # Sort rules by priority (lower number = higher priority)
    sorted_rules = sorted(rules, key=lambda x: x.get('properties', {}).get('priority', 65535))
    
    for rule in sorted_rules:
        name = rule.get('name', 'Unnamed')
        p = rule.get('properties', {})
        
        priority = p.get('priority', 'N/A')
        access = p.get('access', 'Allow')
        direction = p.get('direction', 'Inbound')
        proto = p.get('protocol', '*')
        
        # Source addresses
        src_addr = p.get('sourceAddressPrefix') or ", ".join(p.get('sourceAddressPrefixes', [])) or "*"
        src_port = p.get('sourcePortRange') or ", ".join(p.get('sourcePortRanges', [])) or "*"
        
        # Destination addresses
        dst_addr = p.get('destinationAddressPrefix') or ", ".join(p.get('destinationAddressPrefixes', [])) or "*"
        dst_port = p.get('destinationPortRange') or ", ".join(p.get('destinationPortRanges', [])) or "*"
        
        rule_text = f"[{priority}] {access} {direction} {proto}: {src_addr}:{src_port} -> {dst_addr}:{dst_port} (Rule: {name})"
        rule_texts.append(rule_text)
        
    return "\n".join(rule_texts)

def sanitize_sheet_name(name):
    return re.sub(r'[\\/*?:\[\]]', '_', name)[:31]


def clean_columns(df):
    """Trim noisy columns produced by pd.json_normalize.

    Three passes:
      1. Drop columns whose dotted path contains any NOISE_COLUMN_PATH_SUBSTRINGS.
      2. Drop columns whose leaf segment matches NOISE_COLUMN_LEAVES.
      3. Collapse leaf-name duplicates (e.g. multiple ``identity.*.principalId``):
         keep the first non-empty value across the duplicates under the leaf name.
    Also drops columns that are entirely empty.
    """
    if df is None or df.empty:
        return df

    cols_to_drop = []
    for col in df.columns:
        col_lower = col.lower()
        if any(s in col_lower for s in AzureResourceGraphConfig.NOISE_COLUMN_PATH_SUBSTRINGS):
            cols_to_drop.append(col)
            continue
        leaf = col_lower.rsplit(".", 1)[-1]
        if leaf in AzureResourceGraphConfig.NOISE_COLUMN_LEAVES:
            cols_to_drop.append(col)
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    leaf_groups = defaultdict(list)
    for col in df.columns:
        leaf = col.rsplit(".", 1)[-1]
        leaf_groups[leaf].append(col)

    for leaf, group in leaf_groups.items():
        if len(group) <= 1:
            continue
        coalesced = df[group[0]]
        for other in group[1:]:
            coalesced = coalesced.where(coalesced.notna() & (coalesced != ""), df[other])
        df = df.drop(columns=group)
        df[leaf] = coalesced

    empty_cols = [c for c in df.columns if df[c].isna().all() or (df[c].astype(str).str.strip() == "").all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    return df

def _arg_query_single_scope(env=None, subscription_ids=None):
    """Run the paginated ARG query for one scope (a subscription subset, or all).

    Raises subprocess.CalledProcessError if the `az graph query` call fails so
    the caller can decide whether to skip just this scope or fail outright.
    """
    scope_resources = []
    skip_token = None

    while True:
        cmd = [
            'az', 'graph', 'query',
            '-q', AzureResourceGraphConfig.QUERY,
            '--first', str(AzureResourceGraphConfig.BATCH_SIZE),
            '--output', 'json'
        ]

        # Scope query to selected subscriptions for faster execution and smaller payloads.
        if subscription_ids:
            cmd.extend(['--subscriptions'] + list(subscription_ids))

        # If we have a token from the previous run, use it to get the next batch
        if skip_token:
            cmd.extend(['--skip-token', skip_token])

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=True,
            env=env
        )

        response = json.loads(result.stdout)
        scope_resources.extend(response.get('data', []))

        # Check if Azure gave us a token for the next page
        skip_token = response.get('skip_token')
        if not skip_token:
            break  # No more data, stop looping

        print(f"   ...fetched {len(scope_resources)} resources so far...")

    return scope_resources


def query_azure_resource_graph(env=None, subscription_ids=None):
    all_resources = []

    if subscription_ids and len(subscription_ids) > 1:
        # Combined mode: query each subscription INDEPENDENTLY so that one
        # empty or inaccessible subscription cannot wipe out the results for
        # all the others. A single `az graph query --subscriptions A B C`
        # fails as a whole if any one of A/B/C is invalid or unreadable, which
        # is why the combined workbook previously came back empty.
        print(f"🔍 Querying Azure Resource Graph for {len(subscription_ids)} selected subscription(s) (per-subscription)...")
        for sid in subscription_ids:
            try:
                subset = _arg_query_single_scope(env=env, subscription_ids=[sid])
                print(f"   ✅ {sid}: {len(subset)} resources")
                all_resources.extend(subset)
            except subprocess.CalledProcessError as e:
                print(f"   ⚠️ Skipping subscription {sid} — ARG query failed (no access, invalid, or wrong tenant): {e.stderr}")
                continue
    else:
        if subscription_ids:
            print("🔍 Querying Azure Resource Graph for 1 selected subscription...")
        else:
            print("🔍 Querying Azure Resource Graph (fetching all resources)...")
        try:
            all_resources = _arg_query_single_scope(env=env, subscription_ids=subscription_ids)
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed ARG query: {e.stderr}")
            all_resources = []

    print(f"✅ Total resources fetched: {len(all_resources)}")
    return all_resources

def get_subscription_map(env=None):
    try:
        result = subprocess.run([
            'az', 'account', 'list', '--output', 'json'
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True, env=env)
        subs = json.loads(result.stdout)
        return {s['id']: s['name'] for s in subs}
    except subprocess.CalledProcessError as e:
        print("âŒ Failed to get subscriptions:", e.stderr)
        return {}

def build_nic_to_public_ip_map(env=None, subscription_ids=None):
    query = """
    Resources
    | where type == "microsoft.network/publicipaddresses"
    | where isnotempty(properties.ipConfiguration.id)
    | project
        ipAddress = tostring(properties.ipAddress),
        nicRef = tostring(properties.ipConfiguration.id),
        pipName = name
    """

    all_pips = []
    skip_token = None
    try:
        while True:
            cmd = [
                "az", "graph", "query",
                "-q", query,
                "--first", str(AzureResourceGraphConfig.BATCH_SIZE),
                "--output", "json"
            ]
            if subscription_ids:
                cmd.extend(["--subscriptions"] + list(subscription_ids))
            if skip_token:
                cmd.extend(["--skip-token", skip_token])

            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True,
                env=env
            )
            response = json.loads(result.stdout)
            all_pips.extend(response.get("data", []))
            skip_token = response.get("skip_token")
            if not skip_token:
                break
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to query public IPs via ARG: {e.stderr}")
        return {}

    nic_ip_map = {}
    for pip in all_pips:
        ip = pip.get("ipAddress")
        nic_ref = pip.get("nicRef")
        if nic_ref:
            match = re.search(r'/networkInterfaces/([^/]+)/ipConfigurations', nic_ref)
            if match:
                nic_name = match.group(1)
                id_parts = nic_ref.split("/")
                sub = id_parts[2] if len(id_parts) > 2 else ""
                rg = id_parts[4] if len(id_parts) > 4 else ""
                nic_id = f"/subscriptions/{sub}/resourceGroups/{rg}/providers/Microsoft.Network/networkInterfaces/{nic_name}"
                if ip and nic_id:
                    nic_ip_map[nic_id.lower()] = ip
                else:
                    print(f"⚠️ Skipping public IP - missing IP or NIC ID: {pip.get('pipName')}")

    return nic_ip_map

def fetch_network_details_arg(env=None):
    """
    Fetch all NIC and Subnet details using Azure Resource Graph in a single batch.
    Returns a dictionary mapping NIC ID to network details.
    """
    print("🔍 Querying Azure Resource Graph for network details (Bulk Optimization)...")
    
    query = """
    Resources
    | where type == "microsoft.network/networkinterfaces"
    | mv-expand ipconfig = properties.ipConfigurations
    | project 
        id,
        privateIp = ipconfig.properties.privateIPAddress,
        subnetId = ipconfig.properties.subnet.id,
        publicIpId = ipconfig.properties.publicIPAddress.id
    """
    
    cmd = [
        'az', 'graph', 'query',
        '-q', query,
        '--first', '1000', # Fetch up to 1000 NICs
        '--output', 'json'
    ]
    
    try:
        result = subprocess.run(
            cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            text=True, 
            check=True, 
            env=env
        )
        data = json.loads(result.stdout).get('data', [])
        
        # Build lookup map: nic_id -> details
        nic_details = {}
        for item in data:
            nic_id = item.get('id', '').lower()
            subnet_id = item.get('subnetId', '')
            
            # Parse VNet and Subnet names from ID
            vnet_name = "None"
            subnet_name = "None"
            subnet_prefix = "Unknown" 
            
            if subnet_id:
                parts = subnet_id.split('/')
                try:
                    vnet_idx = parts.index('virtualNetworks')
                    vnet_name = parts[vnet_idx + 1]
                    subnet_idx = parts.index('subnets')
                    subnet_name = parts[subnet_idx + 1]
                except (ValueError, IndexError):
                    pass
            
            nic_details[nic_id] = {
                'private_ip': item.get('privateIp'),
                'vnet_name': vnet_name,
                'subnet_name': subnet_name,
                'subnet_prefix': subnet_prefix,
                'public_ip_id': item.get('publicIpId')
            }
            
        print(f"✅ Fetched details for {len(nic_details)} NICs via ARG")
        return nic_details
        
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Failed ARG network query: {e.stderr}")
        return {}

def enrich_vm_network_details(resources, nic_ip_map, env=None):
    # Pre-fetch all network details in one go
    nic_details_map = fetch_network_details_arg(env)
    
    nic_cache = {}
    subnet_cache = {}

    for res in resources:
        if res.get("type", "").lower() != "microsoft.compute/virtualmachines":
            continue
        shared_disk_count = 0
        os_disk = res.get("properties", {}).get("storageProfile", {}).get("osDisk", {})
        data_disks = res.get("properties", {}).get("storageProfile", {}).get("dataDisks", [])

        # Check OS disk (rare, but possible)
        if os_disk.get("managedDisk", {}).get("shareable"):
            shared_disk_count += 1

        # Check all data disks
        for disk in data_disks:
            if disk.get("managedDisk", {}).get("shareable"):
                shared_disk_count += 1

        res["Shared Disks"] = shared_disk_count
        res["Data Disks Count"] = len(data_disks) if isinstance(data_disks, list) else 0

        vm_name = res.get("name", "Unknown")
        nic_ids = [nic.get("id") for nic in res.get("properties", {}).get("networkProfile", {}).get("networkInterfaces", []) if nic.get("id")]

        public_ip = None
        private_ip = None
        vnet_name = None
        subnet_prefix = None
        subnet_name = None

        for nic_id in nic_ids:
            nic_id_lower = nic_id.lower()
            
            # FAST PATH: Check if we have details from ARG
            if nic_id_lower in nic_details_map:
                details = nic_details_map[nic_id_lower]
                if not private_ip:
                    private_ip = details['private_ip']
                if not vnet_name or vnet_name == "None":
                    vnet_name = details['vnet_name']
                if not subnet_name or subnet_name == "None":
                    subnet_name = details['subnet_name']
                if not subnet_prefix or subnet_prefix == "Unknown":
                    subnet_prefix = details['subnet_prefix']
                
                # Check for public IP in our map
                if nic_id_lower in nic_ip_map:
                    public_ip = nic_ip_map[nic_id_lower]
                continue

            # SLOW PATH: Fallback to individual CLI calls (only if ARG missed it)
            if nic_id_lower in nic_cache:
                nic_data = nic_cache[nic_id_lower]
            else:
                try:
                    nic_data = json.loads(subprocess.check_output([
                        "az", "network", "nic", "show", "--ids", nic_id, "--output", "json"
                    ], env=env))
                    nic_cache[nic_id_lower] = nic_data
                except subprocess.CalledProcessError:
                    continue

            public_ip = nic_ip_map.get(nic_id_lower)

            for ip_cfg in nic_data.get("ipConfigurations", []):
                if not private_ip:
                    private_ip = ip_cfg.get("privateIPAddress")

                subnet = ip_cfg.get("subnet", {})
                subnet_id = subnet.get("id")
                if subnet_id and (not subnet_prefix or not subnet_name):
                    if subnet_id in subnet_cache:
                        subnet_data = subnet_cache[subnet_id]
                    else:
                        try:
                            subnet_data = json.loads(subprocess.check_output([
                                "az", "network", "vnet", "subnet", "show", "--ids", subnet_id, "--output", "json"
                            ], env=env))
                            subnet_cache[subnet_id] = subnet_data
                        except subprocess.CalledProcessError:
                            continue

                    subnet_prefix = (
                        subnet_data.get("addressPrefix")
                        or (subnet_data.get("addressPrefixes", [None])[0])
)

                    subnet_name = subnet_data.get("name")
                    # Extract VNet name from subnet ID
            subnet_parts = subnet_id.split("/")
            try:
                vnet_name = subnet_parts[subnet_parts.index("virtualNetworks") + 1]
            except (ValueError, IndexError):
                vnet_name = None

        # Save clean, friendly keys for Excel
        res["Private IP"] = private_ip or "None"
        res["VNet"] = vnet_name or "None"
        res["Subnet"] = subnet_name or "None"           # Subnet Name
        res["Subnet Prefix"] = subnet_prefix or "None"  # CIDR
        res["Public IP"] = public_ip or "None"

def _export_empty_workbook(sub_map):
    """Write a valid, minimal workbook when a subscription has 0 resources.

    An empty subscription is NOT a failure. We still emit an AllResources sheet
    so the job completes cleanly and the combined report is generated even when
    one (or all) of the selected subscriptions has no resources.
    """
    unique_id = str(uuid.uuid4())[:8]
    file_name = f"{ExcelExportConfig.FILE_NAME_PREFIX}_{datetime.now().strftime(ExcelExportConfig.FILE_NAME_DATE_FORMAT)}_{unique_id}.xlsx"
    file_path = os.path.join(tempfile.gettempdir(), file_name)
    empty_df = pd.DataFrame(
        columns=["Name", "Resource Type", "Resource Group", "Location", "Subscription", "Type"]
    )
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        empty_df.to_excel(writer, sheet_name="AllResources", index=False)
        ws = writer.sheets["AllResources"]
        note_row = len(empty_df) + 3
        ws.cell(row=note_row, column=1,
                value="No resources found for the selected subscription(s).")
        for offset, (sid, name) in enumerate((sub_map or {}).items(), start=1):
            ws.cell(row=note_row + offset, column=1,
                    value=f"Subscription: {name} ({sid}) - 0 resources")
    print(f"✅ Export complete: {file_path}")
    return file_path


def export_to_excel(resources, sub_map, friendly_map, schema_map):
    if not resources:
        # Empty subscription -> emit a minimal workbook instead of failing.
        return _export_empty_workbook(sub_map)
    if not resources:  # unreachable legacy guard (kept to preserve original body below)
        print("âš ï¸ No data to export.")
        return None

    # âœ… Enrich each resource
    for res in resources:
        res["Subscription Name"] = sub_map.get(res["subscriptionId"], "Unknown")
        full_type = res.get("type", "").lower()
        res["Type"] = friendly_map.get(full_type, camel_to_title(full_type.split("/")[-1]))

        if isinstance(res.get("tags"), dict):
            res["tags"] = json.dumps(res["tags"])

    df_all = pd.DataFrame(resources)

    all_res_columns = ["name", "type", "resourceGroup", "location", "subscriptionId", "Type"]
    all_res_renamed = {
        "name": "Name",
        "type": "Resource Type",
        "resourceGroup": "Resource Group",
        "location": "Location",
        "subscriptionId": "Subscription",
        "Type": "Type"
    }

    df_all_res = df_all[all_res_columns].rename(columns=all_res_renamed)


    unique_id = str(uuid.uuid4())[:8]
    file_name = f"{ExcelExportConfig.FILE_NAME_PREFIX}_{datetime.now().strftime(ExcelExportConfig.FILE_NAME_DATE_FORMAT)}_{unique_id}.xlsx"
    file_path = os.path.join(tempfile.gettempdir(), file_name)
    print(f"💾 Saving report to: {file_path}")
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df_all_res.to_excel(writer, sheet_name="AllResources", index=False)
        
        # Add metadata at the bottom
        ws_all = writer.sheets["AllResources"]
        last_row = len(df_all_res) + 3
        duration_sec = int(time.time() - START_TIME)
        duration_str = f"{duration_sec // 60}m {duration_sec % 60}s" if duration_sec > 60 else f"{duration_sec}s"
        
        ws_all.cell(row=last_row, column=1, value="Assessment Date:")
        ws_all.cell(row=last_row, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ws_all.cell(row=last_row+1, column=1, value="Assessment Duration:")
        ws_all.cell(row=last_row+1, column=2, value=duration_str)
        
        ws_all.cell(row=last_row, column=1).font = Font(bold=True)
        ws_all.cell(row=last_row+1, column=1).font = Font(bold=True)

        grouped = defaultdict(list)
        for res in resources:
            grouped[res["type"].lower()].append(res)

        for res_type, items in sorted(grouped.items(), key=lambda x: friendly_map.get(x[0], x[0])):
            # Flatten subnet details for vnets
            if res_type == "microsoft.network/virtualnetworks":
                for item in items:
                    subnets = item.get("properties", {}).get("subnets", [])
                    names, prefixes, states, delegs, endpoints = [], [], [], [], []
                    pe_policies, nsg_ids, rt_ids = [], [], []

                    for subnet in subnets:
                        props = subnet.get("properties", {})

                        names.append(subnet.get("name", ""))
                        prefixes.append(
                            props.get("addressPrefix") or
                            (props.get("addressPrefixes", [None])[0]) or ""
                        )
                        states.append(props.get("provisioningState", ""))
                        pe_policies.append(props.get("privateEndpointNetworkPolicies", ""))
                        nsg_ids.append(props.get("networkSecurityGroup", {}).get("id", ""))
                        rt_ids.append(props.get("routeTable", {}).get("id", ""))

                        deleg_names = [d.get("properties", {}).get("serviceName", "")
                                       for d in props.get("delegations", [])]
                        delegs.append(", ".join(filter(None, deleg_names)))

                        endpoint_services = [e.get("service", "") for e in props.get("serviceEndpoints", [])]
                        endpoints.append(", ".join(filter(None, endpoint_services)))

                    item["Subnet Names"] = "; ".join(names)
                    item["Subnet Prefixes"] = "; ".join(prefixes)
                    item["Subnet Provisioning States"] = "; ".join(states)
                    item["Subnet Delegations"] = "; ".join(delegs)
                    item["Subnet Service Endpoints"] = "; ".join(endpoints)
                    item["Subnet PE Network Policies"] = "; ".join(pe_policies)
                    item["Subnet NSG IDs"] = "; ".join(nsg_ids)
                    item["Subnet Route Table IDs"] = "; ".join(rt_ids)
                    
                    # Flatten peering details for vnets
                    peerings = item.get("properties", {}).get("virtualNetworkPeerings", [])
                    peering_details = []
                    for peering in peerings:
                        p_name = peering.get("name", "Unknown")
                        p_state = peering.get("properties", {}).get("peeringState", "Unknown")
                        remote_vnet_id = peering.get("properties", {}).get("remoteVirtualNetwork", {}).get("id", "")
                        remote_vnet_name = remote_vnet_id.split("/")[-1] if "/" in remote_vnet_id else "Unknown"
                        peering_details.append(f"{p_name} ({p_state}) -> {remote_vnet_name}")
                    item["Virtual Network Peerings"] = "; ".join(peering_details)

            elif res_type in ["microsoft.web/sites", "microsoft.web/sites/slots"]:
                for item in items:
                    props = item.get("properties", {})
                    site_config = props.get("siteConfig", {})
                    
                    # Add requested App Service columns
                    item["STATUS"] = props.get("state", "Unknown")
                    
                    sku_obj = item.get("sku", {})
                    if isinstance(sku_obj, dict):
                        item["PRICING TIER"] = sku_obj.get("name", props.get("sku", "Unknown"))
                    else:
                        item["PRICING TIER"] = sku_obj or props.get("sku", "Unknown")
                    
                    sf_id = props.get("serverFarmId", "")
                    item["APP SERVICE PLAN"] = sf_id.split("/")[-1] if sf_id else "None"
                    
                    item["APP TYPE"] = item.get("kind", "app")
                    
                    # Publishing model
                    fxv = site_config.get("linuxFxVersion", "") or site_config.get("windowsFxVersion", "")
                    item["Publishing model"] = fxv.split("|")[0] if fxv and "|" in fxv else (fxv or "Code")
                    
                    # Application insights
                    has_ins = "No"
                    tags = item.get("tags", {})
                    if isinstance(tags, dict):
                        for k in tags.keys():
                            if "App-Insights-Resource-Id" in k:
                                has_ins = "Yes"
                                break
                    elif isinstance(tags, str) and "App-Insights-Resource-Id" in tags:
                        has_ins = "Yes"
                    item["Application insights"] = has_ins

                    item["Deployment Slot"] = "Yes" if "/slots/" in item.get("id", "").lower() else "No"
                    
                    hostnames = props.get("hostNames", [])
                    default_hostname = props.get("defaultHostName", "")
                    custom_domains = [h for h in hostnames if h != default_hostname]
                    item["Custom domains"] = ", ".join(custom_domains) if custom_domains else "None"
                    
                    ssl_states = props.get("hostNameSslStates", [])
                    item["Certificates"] = sum(1 for s in ssl_states if s.get("thumbprint"))
                    
                    vnet_id = props.get("virtualNetworkSubnetId")
                    pe_conns = props.get("privateEndpointConnections", [])
                    net_info = []
                    if vnet_id: net_info.append("VNet Integrated")
                    if pe_conns: net_info.append("Private Endpoints")
                    item["networking"] = ", ".join(net_info) if net_info else "Public"
                    
                    if vnet_id:
                        parts = vnet_id.split("/")
                        item["Vnet Integration"] = parts[parts.index("virtualNetworks")+1] if "virtualNetworks" in parts else "Integrated"
                    else:
                        item["Vnet Integration"] = "None"
                    
                    identity = item.get("identity", {})
                    item["Managed identity"] = identity.get("type", "None") if isinstance(identity, dict) else "None"
                    
                    item["Private endpoint"] = "Yes" if pe_conns else "No"

            elif res_type == "microsoft.network/networksecuritygroups":
                for item in items:
                    item["NSG Rules"] = convert_nsg_json_to_text(item)

            elif res_type == "microsoft.network/publicipaddresses":
                # Build NIC id -> VM name map so a PIP -> NIC -> VM chain
                # can be resolved to the owning VM.
                nic_to_vm = {}
                for vm in grouped.get("microsoft.compute/virtualmachines", []):
                    vm_name = vm.get("name", "")
                    for nic_ref in vm.get("properties", {}).get("networkProfile", {}).get("networkInterfaces", []):
                        nic_id = (nic_ref or {}).get("id")
                        if nic_id:
                            nic_to_vm[nic_id.lower()] = vm_name

                for item in items:
                    props = item.get("properties", {})
                    ip_cfg = props.get("ipConfiguration") or {}
                    nat_gw = props.get("natGateway") or {}
                    attached_id = ip_cfg.get("id") or nat_gw.get("id") or ""

                    if attached_id:
                        parts = attached_id.split("/")
                        try:
                            p_idx = next(i for i, p in enumerate(parts) if p.lower() == "providers")
                            parent_type_raw = f"{parts[p_idx + 1]}/{parts[p_idx + 2]}".lower()
                            parent_name = parts[p_idx + 3]
                            friendly = friendly_map.get(parent_type_raw, camel_to_title(parent_type_raw.split("/")[-1]))
                            item["Attached"] = "Yes"
                            item["Attached To"] = f"{friendly}: {parent_name}"

                            # If attached via a NIC, also surface the owning VM.
                            if parent_type_raw == "microsoft.network/networkinterfaces":
                                nic_resource_id = "/".join(parts[:p_idx + 4])
                                vm_name = nic_to_vm.get(nic_resource_id.lower(), "")
                                item["Attached VM"] = vm_name or "None"
                            else:
                                item["Attached VM"] = "N/A"
                        except (StopIteration, IndexError):
                            item["Attached"] = "Yes"
                            item["Attached To"] = attached_id
                            item["Attached VM"] = "None"
                    else:
                        item["Attached"] = "No"
                        item["Attached To"] = "None"
                        item["Attached VM"] = "None"

            df = pd.json_normalize(items, max_level=AzureResourceGraphConfig.MAX_FLATTEN_DEPTH)

            # Drop common redundant columns
            cols_to_drop = ["id", "tenantId", "type", "subscriptionId"]
            df.drop(columns=[col for col in cols_to_drop if col in df.columns], inplace=True)

            df = clean_columns(df)

            # Reorder to put 'name' column first
            cols = df.columns.tolist()
            name_col = next((c for c in cols if c.lower() == "name"), None)
            if name_col:
                cols.remove(name_col)
                cols = [name_col] + cols

            # Data Sanitization Loop: Clean strings of '!' and emojis, and normalize numeric types
            # This handles Excel warnings like 'number stored as text' and removes unwanted symbols
            def safe_sanitize(x):
                if x is None or (isinstance(x, float) and pd.isna(x)):
                    return x
                return str(x).replace('!', '').replace('⚠️', '').replace('✅', '').replace('❌', '')

            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].apply(safe_sanitize)
                # Convert common numeric-looking columns to numeric types.
                # Skip string SKU identifiers — vmSize holds values like
                # "Standard_D2s_v5", and sku/sku.name are also string IDs;
                # coercing them to numeric drops the column from the sheet.
                col_lower = col.lower()
                leaf = col_lower.rsplit('.', 1)[-1]
                is_sku_string_col = leaf in ('vmsize', 'sku') or col_lower.endswith('.sku.name')
                if any(k in col_lower for k in ['size', 'gb', 'count', 'memory']) and not is_sku_string_col:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(' GB', '', case=False).replace('Unknown', '0').replace('', '0'), errors='coerce')

            sheet_name = sanitize_sheet_name(friendly_map.get(res_type.lower(), res_type.split("/")[-1]).replace('!', '').replace('⚠️', '').replace('✅', '').replace('❌', ''))
            df[cols].to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Export complete: {file_path}")
    return file_path

def format_excel_with_openpyxl(file_path):
    wb = load_workbook(file_path)
    border = Border(
        left=Side(style=ExcelExportConfig.BORDER_STYLE), 
        right=Side(style=ExcelExportConfig.BORDER_STYLE),
        top=Side(style=ExcelExportConfig.BORDER_STYLE), 
        bottom=Side(style=ExcelExportConfig.BORDER_STYLE)
    )
    header_fill = PatternFill(
        start_color=ExcelExportConfig.HEADER_COLOR, 
        end_color=ExcelExportConfig.HEADER_COLOR, 
        fill_type='solid'
    )
    header_font = Font(color=ExcelExportConfig.HEADER_FONT_COLOR, bold=True)

    for ws in wb.worksheets:
        # Step 1: Process and clean headers
        # We process headers in a list first to avoid index shifting issues during deletion
        header_row = list(ws[1])
        cols_to_delete = []
        
        for cell in header_row:
            if not cell.value:
                continue
                
            orig_val = str(cell.value).strip()
            
            # Context-aware header formatting
            if '.' in orig_val:
                parts = orig_val.split('.')
                # If the last part is too generic, keep the parent part for context (e.g., "Omsagent Enabled")
                generic_names = ["enabled", "type", "id", "name", "status", "kind", "mode", "tier", "sku", "state", "count", "version", "enabled"]
                if parts[-1].lower() in generic_names and len(parts) > 1:
                    # Join last 2 parts with space and convert to title
                    clean_val = f"{camel_to_title(parts[-2])} {camel_to_title(parts[-1])}"
                else:
                    clean_val = camel_to_title(parts[-1])
            else:
                clean_val = camel_to_title(orig_val)
                
            # Check for restricted columns
            if clean_val.lower() in ["id", "tenant id"]:
                cols_to_delete.append(cell.column)
                continue
                
            # Apply formatting
            cell.value = clean_val
            cell.fill = header_fill
            cell.font = header_font
            
        # Step 2: Delete restricted columns in reverse order to maintain indices
        for col_idx in sorted(cols_to_delete, reverse=True):
            ws.delete_cols(col_idx)

        # Delete columns that are fully empty (row 2 to 11)
        for col in reversed(range(1, ws.max_column + 1)):
            col_cells = [ws.cell(row=r, column=col).value for r in range(2, 12)]
            if all(c in (None, "") for c in col_cells):
                ws.delete_cols(col)

        # Get last data row
        last_data_row = 1
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            if any(cell.value not in (None, "") for cell in row):
                last_data_row = row[0].row

        # âœ… Apply border to the entire row if any cell in the row has a value
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            if any(cell.value not in (None, "") for cell in row):
                for cell in row:
                    cell.border = border

        # Auto-fit column widths
        for col in ws.columns:
            header_val = str(col[0].value).strip() if col[0].value else ""
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            
            # Special case for NSG Rules - wrap text and set fixed width
            if header_val.lower() == "nsg rules":
                ws.column_dimensions[col_letter].width = 80
                for cell in col[1:]:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
                continue

            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

    wb.save(file_path)
    print(f"âœ… Excel formatting complete: {file_path}")

def add_migration_plan_sheets(file_path, resources, sub_map):
    from collections import defaultdict

    wb = load_workbook(file_path)

    # Generate resource-level migration plan
    # This uses the imported generate_resource_level_migration_plan from backend.hybrid_analysis_functions
    migration_plan, full_analysis_list = generate_resource_level_migration_plan(resources, sub_map)
    
    # Group by subscription
    sub_plans = defaultdict(list)
    for item in migration_plan:
        sub_name = item["Subscription"]
        sub_plans[sub_name].append(item)

    # Style definitions
    header_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Status Fills
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')  # Light Red
    blue_fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')  # Light Blue
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')  # Light Yellow
    gray_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')  # Light Gray
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')  # Light Green

    base_index = 1  # Place sheets after Dashboard and AllResources

    # Sort by subscription name for deterministic ordering
    sorted_plans = sorted(sub_plans.items())
    for idx, (sub_name, plan_items) in enumerate(sorted_plans):
        # Create Low-Level Migration Plan (detailed resource-level view)
        sheet_name = f"Low-Level Migration Plan_{sub_name.replace(' ', '_')}"[:31]

        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name, index=base_index + idx)

        # Sort items by Move Group to keep related resources together
        plan_items.sort(key=lambda x: x.get("Move Group", ""))

        # Add Legend Row (Row 1)
        legend_labels = [
            "🔴 Red = Cannot Move (Resource type cannot be moved, must recreate)",
            "🔵 Blue = Can Move (Action needed on THIS resource - see Action Required column)",
            "🟡 Yellow = Can Move (Action needed on ANOTHER resource - see Action Required column)",
            "⚪ Gray = Review Required (AI analysis failed - manual review needed)",
            "🟢 Green = Can Move (Ready to move immediately, no action needed)"
        ]
        ws.append(legend_labels + [""] * 5)  # Fill remaining columns
        
        # Style legend row
        legend_row = 1
        ws.cell(row=legend_row, column=1).fill = red_fill
        ws.cell(row=legend_row, column=1).font = Font(bold=True)
        ws.cell(row=legend_row, column=2).fill = blue_fill
        ws.cell(row=legend_row, column=2).font = Font(bold=True)
        ws.cell(row=legend_row, column=3).fill = yellow_fill
        ws.cell(row=legend_row, column=3).font = Font(bold=True)
        ws.cell(row=legend_row, column=4).fill = gray_fill
        ws.cell(row=legend_row, column=4).font = Font(bold=True)
        ws.cell(row=legend_row, column=5).fill = green_fill
        ws.cell(row=legend_row, column=5).font = Font(bold=True)
        
        # Merge cells for better visibility
        ws.merge_cells(start_row=legend_row, start_column=6, end_row=legend_row, end_column=10)

        # Header (Row 2)
        headers = ["Move Group", "Resource Name", "Resource Type", "Resource Group", "Location", 
                   "Move Status", "Action Required", "Dependencies", "Notes", "Source URL"]
        ws.append(headers)
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border

        # Add data rows
        for item in plan_items:
            # Transform status for display (keep internal for coloring)
            display_status = item.get("Move Status", "")
            if display_status in ["Action Required", "Dependent - Action Required"]:
                display_status = "Can Move"
            
            row_data = [
                item.get("Move Group", ""),
                item.get("Resource Name", ""),
                item.get("Resource Type", ""),
                item.get("Resource Group", ""),
                item.get("Location", ""),
                display_status,  # Use transformed status
                item.get("Action Required", ""),
                item.get("Dependencies", ""),
                item.get("Notes", ""),
                item.get("Source URL", "")
            ]
            ws.append(row_data)
            
            # Apply conditional formatting based on status
            status = item.get("Move Status", "")
            fill = None
            if status == "Cannot Move":
                fill = red_fill
            elif status == "Action Required":
                fill = blue_fill
            elif status in ["Can be Moved/Action Required", "Dependent - Action Required"]:
                fill = yellow_fill
            elif status == "Review Required":
                fill = gray_fill
            elif status == "Can Move":
                fill = green_fill
            
            if fill:
                for cell in ws[ws.max_row]:
                    cell.fill = fill

        # Apply borders and auto-fit columns
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=len(headers)):
            for cell in row:
                cell.border = border

        # Auto-fit columns with special handling for Notes and Action Required columns
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(ws.cell(row=row, column=col_idx).value or ""))
                for row in range(1, ws.max_row + 1)
            )
            
            # Special handling for Notes and Action Required columns (allow wider, enable text wrap)
            if header == "Notes":
                ws.column_dimensions[col_letter].width = min(max_len + 4, 100)  # Cap at 100 for Notes
                # Enable text wrapping for Notes column
                for row in range(3, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            elif header in ["Action Required", "Source URL"]:
                ws.column_dimensions[col_letter].width = 50  # Fixed width for Action Required and Source URL
                # Enable text wrapping
                for row in range(3, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    cell.alignment = Alignment(wrap_text=True, vertical='top')
            else:
                ws.column_dimensions[col_letter].width = min(max_len + 4, 50)  # Cap at 50 for others

    wb.save(file_path)
    print("📄 Resource-level migration plan sheets created successfully.")
    return full_analysis_list

def add_high_level_migration_plan(file_path, resources, sub_map, friendly_map):
    """
    Create high-level migration plan sheets grouped by resource type.
    Shows aggregated move status per resource type for each subscription.
    """
    from collections import defaultdict
    
    wb = load_workbook(file_path)
    
    # Group resources by subscription
    sub_resources = defaultdict(list)
    for res in resources:
        sub_id = res.get("subscriptionId")
        if sub_id:
            sub_resources[sub_id].append(res)
    
    # Style definitions
    header_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    green_font = Font(color='00B050', bold=True)  # Green text for "Can Move"
    
    # Calculate position: After Dashboard and AllResources, before Low-Level Migration Plans
    base_index = 1
    
    # Sort subscriptions by name to match the Low-Level plan ordering
    # This ensures High-Level and Low-Level sheets for the same subscription stay together
    sorted_subs = sorted(sub_resources.items(), key=lambda x: sub_map.get(x[0], x[0]))
    
    for idx, (sub_id, sub_res_list) in enumerate(sorted_subs):
        sub_name = sub_map.get(sub_id, sub_id)
        
        # Create High-Level Migration Plan sheet
        sheet_name = f"High-Level Migration Plan_{sub_name.replace(' ', '_')}"[:31]
        
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        
        # Insert before the corresponding Low-Level sheet
        # Position: base_index + (idx * 2) to leave room for Low-Level sheets
        ws = wb.create_sheet(title=sheet_name, index=base_index + (idx * 2))
        
        # Group resources by resource type
        type_analysis = defaultdict(lambda: {
            'total': 0,
            'can_move': 0,
            'action_required': 0,
            'dependent_action': 0,
            'cannot_move': 0,
            'review_required': 0,
            'resources': []
        })
        
        for res in sub_res_list:
            res_type = res.get("Resource Type", "").lower()
            friendly_type = friendly_map.get(res_type, res_type.split("/")[-1] if "/" in res_type else res_type)
            status = res.get("Move Status", "Unknown")
            
            type_analysis[res_type]['total'] += 1
            type_analysis[res_type]['resources'].append(res)
            
            if status == "Can Move":
                type_analysis[res_type]['can_move'] += 1
            elif status == "Action Required":
                type_analysis[res_type]['action_required'] += 1
            elif status == "Dependent - Action Required":
                type_analysis[res_type]['dependent_action'] += 1
            elif status == "Cannot Move":
                type_analysis[res_type]['cannot_move'] += 1
            elif status == "Review Required":
                type_analysis[res_type]['review_required'] += 1
        
        # Add headers
        headers = ["Type", "Resource Type in Azure", "Move Status", "Can be Moved (%)", "Remarks/Comments"]
        ws.append(headers)
        
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        
        # Add data rows (sorted by total count descending)
        sorted_types = sorted(type_analysis.items(), key=lambda x: x[1]['total'], reverse=True)
        
        for res_type, stats in sorted_types:
            # Get friendly name
            friendly_type = friendly_map.get(res_type, res_type.split("/")[-1] if "/" in res_type else res_type)
            
            # Calculate movable count (Can Move + Action Required + Dependent Action)
            movable_count = stats['can_move'] + stats['action_required'] + stats['dependent_action']
            total = stats['total']
            percentage = (movable_count / total * 100) if total > 0 else 0
            
            # Determine Move Status based on 80% threshold
            if percentage >= 80:
                move_status = "Can Move"
                use_green = True
            else:
                move_status = "Reconfigure"
                use_green = False
            
            # Generate Remarks
            if percentage == 100:
                remarks = f"All {total} resource{'s' if total > 1 else ''} can move"
            else:
                reconfigure_count = total - movable_count
                reconfigure_pct = 100 - percentage
                remarks = f"{reconfigure_pct:.0f}% ({reconfigure_count} resource{'s' if reconfigure_count > 1 else ''}) must be reconfigured"
            
            # Add row
            ws.append([
                friendly_type,
                res_type,
                move_status,
                percentage / 100,  # Store as decimal for formatting
                remarks
            ])
            
            # Apply styling
            current_row = ws.max_row
            for col in range(1, 6):
                cell = ws.cell(row=current_row, column=col)
                cell.border = border
                
                # Apply green font to "Can Move" in Move Status column
                if col == 3 and use_green:
                    cell.font = green_font
                
                # Apply percentage formatting to "Can be Moved (%)" column
                if col == 4:
                    cell.number_format = '0%'

                # Center align "Move Status" and "Can be Moved (%)" columns
                if col in [3, 4]:
                    cell.alignment = Alignment(horizontal='center')
        
        # Auto-fit columns
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(ws.cell(row=row, column=col_idx).value or ""))
                for row in range(1, ws.max_row + 1)
            )
            
            # Special handling for Remarks column
            if header == "Remarks/Comments":
                ws.column_dimensions[col_letter].width = min(max_len + 4, 80)
            else:
                ws.column_dimensions[col_letter].width = min(max_len + 4, 50)
    
    wb.save(file_path)
    print("📊 High-level migration plan sheets created successfully.")


def add_iam_roles_sheet(file_path, subscription_ids=None, env=None):
    """Add a dedicated IAM Roles sheet to the Excel file."""
    print("🔐 Adding IAM Roles sheet...")
    
    role_assignments = extract_all_role_assignments(subscription_ids=subscription_ids, env=env)
    
    if not role_assignments:
        print("   ⚠️ No IAM role assignments found")
        return
    
    df = pd.DataFrame(role_assignments)
    wb = load_workbook(file_path)
    
    # Position: Append to the end (after all migration plans)
    sheet_name = "IAM Roles"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    
    header_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    
    subscription_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
    rg_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    resource_fill = PatternFill(start_color='D9EAD3', end_color='D9EAD3', fill_type='solid')
    
    headers = list(df.columns)
    ws.append(headers)
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    
    for _, row in df.iterrows():
        row_data = [row[col] for col in headers]
        ws.append(row_data)
        
        scope_type = row.get("Scope Type", "")
        fill = None
        if scope_type == "Subscription":
            fill = subscription_fill
        elif scope_type == "Resource Group":
            fill = rg_fill
        elif scope_type == "Resource":
            fill = resource_fill
        
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill
                cell.border = border
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            if not cell.fill or cell.fill.start_color.rgb == '00000000':
                cell.border = border
    
    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=row, column=col_idx).value or ""))
                     for row in range(1, ws.max_row + 1))
        
        if header == "Scope":
            ws.column_dimensions[col_letter].width = min(max_len + 4, 80)
        elif header in ["Principal ID"]:
            ws.column_dimensions[col_letter].width = 40
        else:
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)
    
    wb.save(file_path)
    print(f"   ✅ IAM Roles sheet created with {len(role_assignments)} assignments")

def fetch_all_resource_groups(env=None, subscription_ids=None):
    """Query ARG `resourcecontainers` for every RG with its original-case name.

    ARG's `resourceGroup` field on individual resources is normalized to
    lowercase, and RGs that contain only filtered-out resources never show
    up in the resource list at all. Querying `resourcecontainers` directly
    returns the canonical RG record (case-preserved `name`) for every RG
    the caller has read access to — regardless of what's inside.
    """
    query = "resourcecontainers | where type == 'microsoft.resources/subscriptions/resourcegroups' | project name, subscriptionId, location"
    rgs = []
    skip_token = None
    try:
        while True:
            cmd = [
                "az", "graph", "query",
                "-q", query,
                "--first", str(AzureResourceGraphConfig.BATCH_SIZE),
                "--output", "json",
            ]
            if subscription_ids:
                cmd.extend(["--subscriptions"] + list(subscription_ids))
            if skip_token:
                cmd.extend(["--skip-token", skip_token])
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, check=True, env=env)
            response = json.loads(result.stdout)
            rgs.extend(response.get("data", []))
            skip_token = response.get("skip_token")
            if not skip_token:
                break
        print(f"✅ Fetched {len(rgs)} resource groups from ARG")
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Could not fetch resource groups: {e.stderr}")
    except Exception as e:
        print(f"⚠️ Could not fetch resource groups: {e}")
    return rgs


def add_rg_summary_sheet(file_path, resources, all_rgs=None):
    """Add a sheet named 'RG' showing resource counts per resource group.

    ``all_rgs`` is the canonical list from ARG's resourcecontainers query.
    When provided, EVERY resource group appears — even those whose only
    resources were filtered by EXCLUDE_KEYWORDS (e.g., NetworkWatcherRG).
    RG names are also shown in their original case (Azure preserves it in
    the container record; the per-resource `resourceGroup` field is
    lower-cased, which is why the previous report showed 'production'
    instead of 'Production').
    """
    print("📂 Adding RG summary sheet...")

    # Case-insensitive count from surviving resources.
    rg_counts_lower = Counter(
        (res.get("resourceGroup") or "Unknown").lower()
        for res in resources
        if res.get("resourceGroup")
    )

    # Build canonical (case-preserved) name → count map. Seed from the
    # ARG resourcecontainers query so RGs with 0 surviving resources still
    # appear; fall back to whatever case the resources gave us.
    rg_counts = {}
    if all_rgs:
        for rg in all_rgs:
            canonical = rg.get("name")
            if canonical:
                rg_counts[canonical] = rg_counts_lower.get(canonical.lower(), 0)
    # Any resource-derived RG not in the canonical list (shouldn't happen
    # normally, but be defensive if the resourcecontainers query failed
    # or the RG lives in a subscription the query didn't cover).
    seen_lower = {n.lower() for n in rg_counts}
    for lower_name, count in rg_counts_lower.items():
        if lower_name not in seen_lower:
            rg_counts[lower_name] = count

    wb = load_workbook(file_path)
    sheet_name = "RG"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    
    # Position the summary directly after AllResources.
    ws = wb.create_sheet(title=sheet_name, index=1)
    
    # Style definitions
    header_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    total_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
    total_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Headers
    headers = ["Row Labels", "Count of NAME"]
    ws.append(headers)
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    
    # Data rows (sorted by RG name)
    total_count = 0
    for rg, count in sorted(rg_counts.items()):
        ws.append([rg, count])
        total_count += count
        for col in range(1, 3):
            ws.cell(row=ws.max_row, column=col).border = border
            
    # Grand Total row
    ws.append(["Grand Total", total_count])
    total_row = ws.max_row
    for col in range(1, 3):
        cell = ws.cell(row=total_row, column=col)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        
    # Auto-fit columns
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 20
    
    wb.save(file_path)
    print(f"   ✅ RG summary sheet created with {len(rg_counts)} resource groups")

    def fetch_reservations(env=None):
    
        try:
            cmd = ["az", "reservations", "reservation-order", "list", "--output", "json"]
            result = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True, env=env)
            if result.strip() == "[]":
                print("ðŸš« No reservations found or no access to view them.")
                return

            reservations = json.loads(result)
            if not reservations:
                print("ðŸš« No reservations found or no access to view them.")
            else:
                print("âœ… Reservations:")
                for r in reservations:
                    print(f"- Reservation Order ID: {r.get('id')}")
        except subprocess.CalledProcessError as e:
            output = e.output.lower()
            if any(error in output for error in ["authorizationfailed", "not allowed", "no registered resource provider"]):
                print("ðŸš« No access to view reservations.")
            else:
                print(f"âŒ Error fetching reservations:\n{e.output.strip()}")


def _fetch_provider_quota(provider, api_version, subscription_id, region, env):
    # Use unsupported regions from config
    unsupported_regions = AzureQuotaConfig.UNSUPPORTED_REGIONS
    
    # Skip if region is known to be unsupported for this provider
    if region.lower() in unsupported_regions.get(provider, []):
        return []
    
    cmd = [
        "az", "rest",
        "--method", "get",
        "--uri", f"https://management.azure.com/subscriptions/{subscription_id}/providers/{provider}/locations/{region}/usages?api-version={api_version}",
        "--query", "value[].{Name:name.localizedValue, Current:currentValue, Limit:limit}",
        "--output", "json"
    ]

    try:
        # Suppress stderr to avoid noisy Azure API errors in logs
        result = subprocess.check_output(cmd, text=True, env=env, stderr=subprocess.DEVNULL)
        parsed = json.loads(result)
        if isinstance(parsed, list):
            return parsed
    except subprocess.CalledProcessError:
        # Silently skip - error is expected for some region/provider combinations
        pass
    return []

def fetch_quota_for_region(subscription_id, region, env=None):
    # Use quota providers from config
    providers = AzureQuotaConfig.PROVIDER_API_VERSIONS

    all_quotas = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=AzureQuotaConfig.MAX_WORKERS) as executor:
        futures = [
            executor.submit(_fetch_provider_quota, provider, api_version, subscription_id, region, env)
            for provider, api_version in providers.items()
        ]
        
        for future in concurrent.futures.as_completed(futures):
            all_quotas.extend(future.result())

    return all_quotas

def append_quotas_to_excel(file_path, used_regions_map, sub_map, env=None):
    print("ðŸ“¥ Appending quota usage to Excel...")

    wb = load_workbook(file_path)

    # Thin border style for all cells
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for sub_id, regions in used_regions_map.items():
        sub_name = sub_map.get(sub_id, sub_id)
        sheet_name = f"{sub_name}_Quota"[:31]  # Excel limit

        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(title=sheet_name)
            headers = ["Subscription", "Region", "Category", "Name", "Current", "Limit"]
            ws.append(headers)

            # ðŸŽ¨ Apply header formatting
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
                cell.border = thin_border

        for region in regions:
            if region.lower() == "global":
                print(f"â­ï¸ Skipping unsupported region: {region}")
                continue

            try:
                quota_data = fetch_quota_for_region(sub_id, region, env=env)
                if not quota_data:
                    continue

                used_quotas = [
                    item for item in quota_data
                    if item.get("Current") not in (None, 0, "0", "0.0")
                ]
                if not used_quotas:
                    print(f"â„¹ï¸ No usage in region: {region} (sub: {sub_name})")
                    continue

                categories = {
    "Compute": [
        "vcpus", "cores", "virtual machines", "dedicated", "low-priority", "spot", 
        "availability set", "proximity placement", "host group", "scale set", "gpu"
    ],
    "Storage": [
        "storage", "disk", "unmanaged", "managed", "blob", "queue", "file", "table", 
        "standardssd", "premium storage", "snapshots"
    ],
    "Networking": [
        "ip address", "load balancer", "public ip", "private ip", "network interface", 
        "virtual network", "gateway", "vnet", "subnet", "dns", "expressroute", "vpn", 
        "application gateway", "firewall", "nat gateway"
    ],
    "Databases": [
        "sql", "mysql", "postgres", "cosmos", "mariadb", "elastic pool", "database"
    ],
    "Containers": [
        "kubernetes", "aks", "container instance", "container group", "container registry"
    ],
    "AI/ML": [
        "cognitive services", "openai", "ml workspace", "bot service"
    ],
    "Integration": [
        "logic app", "function", "event grid", "event hub", "service bus", "data factory", 
        "data integration"
    ],
    "Security": [
        "key vault", "security center", "defender", "sentinel", "policy", "identity", 
        "access"
    ],
    "Monitoring": [
        "log analytics", "application insights", "monitor", "metrics", "alerts", 
        "autoscalesetting", "action group"
    ],
    "Web & App Services": [
        "app service", "app plan", "web app", "api management", "static site", "function app"
    ]
}

                def get_category(name):
                    name = name.lower()
                    for category, keywords in categories.items():
                        if any(keyword in name for keyword in keywords):
                            return category
                    return "Other"


                for item in used_quotas:
                    row = [
                        sub_name,
                        region,
                        get_category(item.get("Name", "")),
                        item.get("Name", ""),
                        item.get("Current", ""),
                        item.get("Limit", "")
                    ]
                    ws.append(row)

                print(f"âœ… Added quota for {region} ({sub_name})")

            except Exception as e:
                print(f"âŒ Failed to get quota for {region} in {sub_id}: {e}")

        # ðŸ–Œï¸ Apply border to all cells and auto-fit columns
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=6):
            for cell in row:
                cell.border = thin_border

        for col in range(1, 7):  # Adjust columns A to F
            max_length = max(len(str(ws.cell(row=row, column=col).value or "")) for row in range(1, ws.max_row + 1))
            ws.column_dimensions[get_column_letter(col)].width = max_length + 2

    wb.save(file_path)
    print("âœ… Quota info added with formatting.")


# Functions build_dependency_graph, analyze_move_feasibility, and generate_resource_level_migration_plan are imported from backend.hybrid_analysis_functions


def run_assessment(subscription_ids=None, env=None):
    print("ðŸ” Fetching subscription names...")
    sub_map = get_subscription_map(env=env)

    # Filter sub_map if specific subscriptions are requested
    if subscription_ids:
        print(f"🎯 Filtering for requested subscriptions: {subscription_ids}")
        normalized_ids = [s.lower() for s in subscription_ids]
        
        # Switch context to the first requested subscription to ensure we can query its resources
        # (Handles cross-tenant scenarios where we might be authenticated to Tenant A but assessing Tenant B)
        try:
            target_sub = subscription_ids[0]
            print(f"🔄 Switching CLI context to subscription {target_sub}...")
            subprocess.run(["az", "account", "set", "--subscription", target_sub], check=True, env=env)
            print(f"✅ Context switched to {target_sub}")
        except Exception as e:
            print(f"⚠️ Failed to switch context (this might be fine if already in correct tenant): {e}")

        # Update sub_map to include requests even if not in initial list (if we trust the input)
        # But usually better to filter
        sub_map = {k: v for k, v in sub_map.items() if k.lower() in normalized_ids}
        
        if not sub_map:
             # If exact match failed, try to be lenient or blindly accept if context switch worked?
             # For now, let's assume if context switch worked, we can query.
             # But we need names. Recurse or re-fetch?
             # Let's keep existing behavior warning for now but proceed if we switched context.
             print(f"⚠️ Warning: Requested subscriptions not found in initial 'az account list'. Context switched, so proceed carefully.")
             # Add strictly provided IDs to map as best effort
             for sid in subscription_ids:
                 if sid not in sub_map:
                     sub_map[sid] = f"Subscription ({sid})"


    print("ðŸ“¦ Querying Azure Resource Graph...")
    # Pass subscription_ids to query_azure_resource_graph if needed, 
    # but currently query_azure_resource_graph doesn't accept args.
    # We can filter the results or modify query_azure_resource_graph.
    # For now, let's filter the results based on the sub_map we just filtered.
    target_sub_ids = list(sub_map.keys()) if subscription_ids else None
    raw_resources = query_azure_resource_graph(env=env, subscription_ids=target_sub_ids)
    print(f"🔍 Raw resources fetched: {len(raw_resources)}")

    # Diagnostic: per-type counts BEFORE filtering
    raw_type_counts = Counter((r.get("type") or "").lower() for r in raw_resources)
    raw_sub_counts = Counter((r.get("subscriptionId") or "").lower() for r in raw_resources)
    print("📊 [DIAG] Raw resource counts by type (top 30):")
    for t, c in raw_type_counts.most_common(30):
        print(f"     {c:5d}  {t}")
    print(f"📊 [DIAG] Raw resource subscriptions seen ({len(raw_sub_counts)}):")
    for sid, c in raw_sub_counts.most_common():
        print(f"     {c:5d}  {sid}")

    # Use exclude keywords from config. IMPORTANT: matched resources are NOT
    # dropped anymore — they're TAGGED with `is_hidden_type=True`. The full
    # set (`resources`) still goes into every report sheet (All Resources,
    # RG summary, per-type sheets) so the total matches what the Azure
    # portal shows. Only the migration analysis / dashboard's "top blocking
    # / critical action" logic filters them out — those are internal/system
    # resources (NetworkWatcher, backup fabrics, hidden monitor rules) that
    # shouldn't drive migration decisions.
    exclude_keywords = AzureResourceGraphConfig.EXCLUDE_KEYWORDS
    selected_subs_lower = [s.lower() for s in sub_map.keys()]

    hidden_count = Counter()              # type -> count (informational)
    dropped_by_subfilter = Counter()      # subscriptionId -> count (still dropped)
    hidden_examples = defaultdict(list)   # type -> [example resource names]
    resources = []
    for res in raw_resources:
        rtype = (res.get("type") or "").lower()
        rsub = (res.get("subscriptionId") or "").lower()
        # Subscription filter still drops — if the user picked a subset of
        # subscriptions, resources in other subs genuinely shouldn't appear.
        if subscription_ids and rsub not in selected_subs_lower:
            dropped_by_subfilter[rsub] += 1
            continue
        # Hidden-type filter now tags instead of drops.
        matched_kw = next((kw for kw in exclude_keywords if kw in rtype), None)
        if matched_kw:
            res["is_hidden_type"] = True
            res["_hidden_reason"] = matched_kw
            hidden_count[f"{rtype}  (kw='{matched_kw}')"] += 1
            if len(hidden_examples[rtype]) < 3:
                hidden_examples[rtype].append(res.get("name", "?"))
        else:
            res["is_hidden_type"] = False
        resources.append(res)

    hidden_total = sum(1 for r in resources if r.get("is_hidden_type"))
    print(f"✅ Resources in scope: {len(resources)} (of which {hidden_total} are tagged hidden for analysis)")
    print(f"📊 [DIAG] Hidden-type tagged (still shown in All Resources / RG): {sum(hidden_count.values())}")
    for label, c in hidden_count.most_common():
        rtype_only = label.split("  (kw=")[0]
        ex = hidden_examples.get(rtype_only, [])
        print(f"     {c:5d}  {label}   examples={ex}")
    if subscription_ids:
        print(f"📊 [DIAG] Dropped by SUBSCRIPTION filter: {sum(dropped_by_subfilter.values())}")
        for sid, c in dropped_by_subfilter.most_common():
            print(f"     {c:5d}  {sid}")
        print(f"📊 [DIAG] Selected subscriptions: {selected_subs_lower}")
    else:
        print("📊 [DIAG] No subscription filter applied (all subs in scope).")

    # Per-type counts AFTER filtering
    surviving_type_counts = Counter((r.get("type") or "").lower() for r in resources)
    print("📊 [DIAG] Surviving resource counts by type (top 30):")
    for t, c in surviving_type_counts.most_common(30):
        delta = raw_type_counts.get(t, 0) - c
        marker = f"  (-{delta} dropped)" if delta else ""
        print(f"     {c:5d}  {t}{marker}")

    # Spotlight on commonly-asked-about types
    spotlight_types = [
        "microsoft.network/publicipaddresses",
        "microsoft.network/networkinterfaces",
        "microsoft.compute/virtualmachines",
        "microsoft.network/loadbalancers",
        "microsoft.network/virtualnetworks",
        "microsoft.storage/storageaccounts",
    ]
    print("📊 [DIAG] Spotlight (raw -> surviving):")
    for t in spotlight_types:
        r = raw_type_counts.get(t, 0)
        s = surviving_type_counts.get(t, 0)
        print(f"     {t}: {r} -> {s}")

    print("🌐 Building NIC to Public IP map...")
    nic_ip_map = build_nic_to_public_ip_map(env=env, subscription_ids=target_sub_ids)

    print("🌐 Enriching VM resources with Public IP...")
    enrich_vm_network_details(resources, nic_ip_map, env=env)
    
    if PerformanceConfig.ENABLE_IAM_ANALYSIS:
        print("🔐 Extracting IAM role assignments...")
        iam_roles_map = extract_iam_roles(subscription_ids=target_sub_ids, env=env)
        enrich_resources_with_iam(resources, iam_roles_map)
    else:
        print("⏭️ Skipping IAM role extraction (ENABLE_IAM_ANALYSIS=false)")
    
    print("💰 Extracting VM Spot Instance details...")
    extract_spot_vm_details(resources)

    print("🧠 Loading friendly type mappings...")
    friendly_map = load_friendly_type_map()

    print("📚 Loading resource type schema map...")

    print("🔍 Detecting AKS-related resources...")

    # Detect AKS-like resources
    for res in resources:
        res_type = res.get("type", "").lower()
        name = res.get("name", "").lower()
        tags = res.get("tags", {}) or {}
        full_type = res.get("type", "").lower()


        res["Type"] = friendly_map.get(full_type, camel_to_title(full_type.split("/")[-1]))

        # Use AKS detection settings from config
        if res_type in [rt.lower() for rt in AKSDetectionConfig.AKS_RELATED_RESOURCE_TYPES]:
            if (
                any(keyword in name for keyword in AKSDetectionConfig.AKS_KEYWORDS) or
                any(keyword in str(k).lower() or keyword in str(v).lower() 
                    for k, v in tags.items() 
                    for keyword in AKSDetectionConfig.AKS_KEYWORDS)
            ):
                res["Type"] = AKSDetectionConfig.AKS_RESOURCE_LABEL

    print("ðŸ“„ Exporting to Excel...")
    file_path = export_to_excel(resources, sub_map, friendly_map, {})

    if file_path and not resources:
        # Empty subscription(s): export_to_excel already wrote a valid minimal
        # workbook. Skip migration/dashboard/quota/IAM analysis since there is
        # nothing to analyze and those steps assume resource data.
        print("No resources to analyze - skipping migration/dashboard/quota/IAM steps.")
        return file_path

    if file_path:
        print("ðŸŽ¨ Formatting Excel file...")
        format_excel_with_openpyxl(file_path)

        # Add RG summary sheet — fetch the canonical RG list from ARG's
        # resourcecontainers so ALL groups appear (even empty ones and ones
        # whose only resources were filtered) and names keep their case.
        all_rgs = fetch_all_resource_groups(env=env, subscription_ids=target_sub_ids)
        add_rg_summary_sheet(file_path, resources, all_rgs=all_rgs)

        # Migration analysis (and its downstream dashboard) sees ONLY the
        # non-hidden resources. Hidden types (NetworkWatcher, backup fabrics,
        # etc.) are still in the report sheets above, but they shouldn't
        # drive migration decisions or appear as top blockers in the
        # dashboard's critical-action panel.
        resources_for_analysis = [r for r in resources if not r.get("is_hidden_type")]
        print(f"📊 Adding migration plan with dependency analysis "
              f"({len(resources_for_analysis)} resources, "
              f"{len(resources) - len(resources_for_analysis)} hidden types skipped)...")
        migration_plan = add_migration_plan_sheets(file_path, resources_for_analysis, sub_map)
        
        print("📊 Adding high-level migration plan (aggregated by resource type)...")
        add_high_level_migration_plan(file_path, migration_plan, sub_map, friendly_map)

        if PerformanceConfig.ENABLE_IAM_ANALYSIS:
            if PerformanceConfig.ENABLE_IAM_SHEET:
                add_iam_roles_sheet(file_path, subscription_ids=target_sub_ids, env=env)
            else:
                print("⏭️ Skipping IAM Roles sheet (ENABLE_IAM_SHEET=false)")
        else:
            print("⏭️ Skipping IAM Roles sheet (ENABLE_IAM_ANALYSIS=false)")

        # âœ… Build used region map from the resources
        used_regions_map = {}
        for res in resources:
            sub_id = res.get("subscriptionId")
            region = res.get("location", "").lower()
            if sub_id and region:
                used_regions_map.setdefault(sub_id, set()).add(region)
        used_regions_map = {k: list(v) for k, v in used_regions_map.items()}

        if PerformanceConfig.ENABLE_QUOTA_ANALYSIS:
            print("ðŸ“ˆ Appending quota info based on usage & regions...")
            append_quotas_to_excel(file_path, used_regions_map, sub_map, env=env)
            print("ðŸš« No access to view reservations.")
        else:
            print("⏭️ Skipping quota analysis (ENABLE_QUOTA_ANALYSIS=false)")
        # Create Dashboard with comprehensive migration insights
        # Dashboard shows portal-matching total (analyzed + hidden) but
        # analysis breakdowns (Can Move / Blockers) only cover analyzed set.
        create_improved_dashboard(
            file_path, migration_plan, sub_map, friendly_map,
            total_all_resources=len(resources),
            hidden_count=len(resources) - len(resources_for_analysis),
        )
        
        return file_path
    return None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Azure Resource Assessment Script")
    parser.add_argument("--subscription-ids", nargs="*", help="List of subscription IDs to assess")
    parser.add_argument(
        "--workbook-mode",
        choices=["combined", "per_subscription"],
        default="combined",
        help="Generate one combined workbook or one workbook per selected subscription",
    )
    args = parser.parse_args()

    if args.workbook_mode == "per_subscription" and args.subscription_ids and len(args.subscription_ids) > 1:
        generated_files = []
        for index, subscription_id in enumerate(args.subscription_ids, start=1):
            print(
                f"Generating workbook {index}/{len(args.subscription_ids)} "
                f"for subscription {subscription_id}..."
            )
            file_path = run_assessment(subscription_ids=[subscription_id])
            if file_path:
                generated_files.append(file_path)
        if len(generated_files) != len(args.subscription_ids):
            raise SystemExit(
                f"Generated {len(generated_files)} of {len(args.subscription_ids)} requested workbooks"
            )
    else:
        run_assessment(subscription_ids=args.subscription_ids)

def run_assessment_with_credential(credential, subscription_ids=None):
    """
    Run assessment using a credential object (from SDK device code auth).
    Gets an access token and uses it to authenticate az CLI.
    """
    print("🔐 Running assessment with SDK credential...")
    
    try:
        # Get access token from the credential
        print("🔑 Getting access token from credential...")
        token = credential.get_token("https://management.azure.com/.default")
        access_token = token.token
        
        # Login to az CLI using the access token
        print("🔐 Logging into az CLI with access token...")
        result = subprocess.run(
            ["az", "login", "--access-token", access_token],
            capture_output=True,
            text=True,
            check=True
        )
        
        if result.returncode == 0:
            print("✅ Successfully authenticated with az CLI")
        
        # Now run the assessment with az CLI authenticated
        return run_assessment(subscription_ids=subscription_ids)
        
    except Exception as e:
        print(f"❌ Error authenticating with credential: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

