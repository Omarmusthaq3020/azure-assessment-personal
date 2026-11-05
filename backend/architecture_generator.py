"""
Architecture Diagram Generator
================================
Extracts environment data from assessment Excel reports and uses Azure OpenAI
to generate Mermaid architecture diagrams for visual understanding.

Supports: Azure, AWS, GCP assessments.
"""

import base64
import json
import os
import sys
import re
from datetime import datetime
from collections import defaultdict, Counter
from typing import Dict, List, Optional
from urllib.parse import quote
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
import ast

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import AzureOpenAIConfig

# ---------------------------------------------------------------------------
# Local Azure icon resolver
# ---------------------------------------------------------------------------
# Diagrams previously fetched icons from azureicons.com / twemoji, which was
# unreliable (occasional wrong icon, network dependency, broken on offline /
# air-gapped deploys). We now resolve icons from the local Azure-Icons folder
# the user downloaded and embed them as base64 data URIs so the resulting
# drawio XML is fully self-contained — works in the in-app playground AND
# when the XML is opened in app.diagrams.net or desktop drawio.

# Project root contains the Azure-Icons folder at the repo root level.
_ICONS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Azure-Icons",
)

# Override map for names where azureicons.com's path differs from the local
# Azure-Icons folder layout, or where we want to substitute a different icon.
# Keys are (category, name) used in code; values are local relative paths.
_LOCAL_ICON_OVERRIDES: Dict[tuple, str] = {
    # Microsoft's local distribution files Kubernetes under Compute, not
    # Containers (which is azureicons.com's category).
    ("Containers", "Kubernetes-Services"): "Compute/Kubernetes-Services.svg",
    # Twemoji glyphs were the previous fallback; map to Azure-native icons.
    ("Twemoji", "User"): "Identity/Users.svg",
    ("Twemoji", "Internet"): "Networking/Public-IP-Addresses.svg",
    ("Twemoji", "Peering"): "Other/Peerings.svg",
}

# Cache so each SVG is read + encoded only once per process.
_ICON_DATA_URI_CACHE: Dict[str, str] = {}


def _resolve_icon_path(category: str, name: str) -> Optional[str]:
    """Return absolute path to the local SVG for (category, name), or None."""
    key = (category, name)
    if key in _LOCAL_ICON_OVERRIDES:
        candidate = os.path.join(_ICONS_DIR, _LOCAL_ICON_OVERRIDES[key])
        if os.path.isfile(candidate):
            return candidate
    candidate = os.path.join(_ICONS_DIR, category, f"{name}.svg")
    if os.path.isfile(candidate):
        return candidate
    return None


def get_icon_uri(category: str, name: str, fallback_url: str = "") -> str:
    """Return a drawio-safe self-contained URI for the requested icon.

    Looks up Azure-Icons/<category>/<name>.svg (with override map for known
    layout differences) and returns a URL-encoded inline SVG data URI:
    ``data:image/svg+xml,<percent-encoded-svg>``.

    Why URL-encoded inline (NOT ``;base64,``):
    drawio's style parser splits ``key=value`` pairs on ``;``. A standard
    base64 data URI like ``data:image/svg+xml;base64,XXX`` therefore gets
    truncated at the ``;`` between the MIME type and the ``base64``
    indicator, so the image reference loses its actual content and the
    icon silently fails to render. URL-encoded inline SVG has no ``;`` in
    its structure (``,`` is the only separator), and ``urllib.parse.quote``
    percent-encodes every ``;`` and ``=`` found inside the SVG body so
    drawio's parser sees the URI as a single opaque value.

    If the file is missing, falls back to ``fallback_url`` (the remote
    URL) and prints a one-time warning so gaps in the local icon set are
    visible during diagram generation.
    """
    cache_key = f"{category}/{name}"
    if cache_key in _ICON_DATA_URI_CACHE:
        return _ICON_DATA_URI_CACHE[cache_key]

    path = _resolve_icon_path(category, name)
    if path:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                svg_text = fh.read()
            # Percent-encode the SVG body so drawio's style parser cannot
            # split it on ';' or interpret '=' as a key/value separator.
            # ``quote`` encodes ';' as %3B, '=' as %3D, '&' as %26, etc.
            encoded = quote(svg_text, safe="")
            uri = f"data:image/svg+xml,{encoded}"
            _ICON_DATA_URI_CACHE[cache_key] = uri
            return uri
        except OSError as e:
            print(f"⚠️ Local icon read failed for {cache_key}: {e}. Falling back to remote.")

    # Local file missing — fall back to remote URL if provided, else empty
    # so drawio renders the box without an image (visible gap, debuggable).
    if fallback_url:
        print(f"⚠️ Local Azure-Icons missing {category}/{name}.svg — using remote fallback.")
    _ICON_DATA_URI_CACHE[cache_key] = fallback_url
    return fallback_url

# ---------------------------------------------------------------------------
# 1. Data Extraction — Parse Excel to structured JSON
# ---------------------------------------------------------------------------

def extract_environment_summary(file_path: str, cloud_type: str = None) -> dict:
    """
    Parse an assessment Excel report and return a compact environment summary.

    Auto-detects cloud type from sheet names if not provided.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Report not found: {file_path}")

    wb = load_workbook(file_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames

    # Auto-detect cloud type
    if cloud_type is None:
        cloud_type = _detect_cloud_type(sheet_names, file_path)

    if cloud_type == "azure":
        summary = _extract_azure_summary(wb, sheet_names)
    elif cloud_type == "aws":
        summary = _extract_aws_summary(wb, sheet_names)
    elif cloud_type == "gcp":
        summary = _extract_gcp_summary(wb, sheet_names)
    else:
        summary = _extract_generic_summary(wb, sheet_names)

    summary["cloud"] = cloud_type
    wb.close()
    return summary


def _detect_cloud_type(sheet_names: list, file_path: str) -> str:
    """Detect cloud type from sheet names and filename."""
    names_lower = [s.lower() for s in sheet_names]
    basename = os.path.basename(file_path).lower()

    if "aws" in basename or any("ec2" in n for n in names_lower):
        return "aws"
    if "gcp" in basename or any("instance" in n and "gcp" in basename for n in names_lower):
        return "gcp"
    # Azure is the default (has migration plan sheets)
    if any("migration plan" in n for n in names_lower):
        return "azure"
    if any("virtualnetworks" in n or "virtualmachines" in n for n in names_lower):
        return "azure"
    return "azure"  # Default


def _sheet_to_records(wb, sheet_name: str, max_rows: int = 500) -> list:
    """Read a worksheet into a list of dicts (capped for token budget)."""
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return []
    headers = [str(h or "").strip() for h in rows[0]]
    records = []
    for row in rows[1:max_rows + 1]:
        record = {}
        for i, val in enumerate(row):
            if i < len(headers) and headers[i]:
                key = headers[i]
                # Some worksheets contain duplicate header names (e.g., multiple "Name" columns).
                # Keep the first non-empty occurrence to avoid later technical columns overwriting
                # primary identity fields used for topology extraction.
                if key in record:
                    if (record[key] is None or str(record[key]).strip() == "") and val is not None and str(val).strip() != "":
                        record[key] = val
                else:
                    record[key] = val
        records.append(record)
    return records


def _norm_key(key: str) -> str:
    """Normalize header keys for tolerant matching."""
    return re.sub(r"[^a-z0-9]", "", str(key or "").lower())


def _pick(record: dict, *candidates, default=""):
    """
    Case/format-insensitive getter for Excel-derived records.
    Example: matches `VNet`, `Vnet`, `virtual network`.
    """
    if not record:
        return default

    norm_map = {_norm_key(k): v for k, v in record.items()}
    for cand in candidates:
        val = norm_map.get(_norm_key(cand))
        if val is not None and str(val).strip() != "":
            return val
    return default


def _find_sheet(sheet_names: list, keywords: list) -> str:
    """Find a sheet whose name contains any of the keywords (case-insensitive)."""
    for name in sheet_names:
        name_lower = name.lower()
        for kw in keywords:
            if kw.lower() in name_lower:
                return name
    return None


def _find_sheet_prefer_exact(sheet_names: list, exact_names: list, fallback_keywords: list = None) -> str:
    """Prefer exact sheet names first, then fallback to keyword contains matching."""
    exact_lookup = {s.lower(): s for s in sheet_names}
    for n in exact_names:
        s = exact_lookup.get(str(n).lower())
        if s:
            return s
    if fallback_keywords:
        return _find_sheet(sheet_names, fallback_keywords)
    return None


# ---------------------------------------------------------------------------
# 1a. Azure Data Extraction
# ---------------------------------------------------------------------------

def _extract_azure_summary(wb, sheet_names: list) -> dict:
    """Extract Azure environment summary from assessment report."""
    summary = {
        "resource_groups": [],
        "networks": [],
        "compute": [],
        "databases": [],
        "storage": [],
        "aks_clusters": [],
        "vm_scale_sets": [],
        "web_apps": [],
        "static_web_apps": [],
        "container_apps": [],
        "storage_accounts": [],
        "database_services": [],
        "load_balancers": [],
        "application_gateways": [],
        "frontdoor_endpoints": [],
        "traffic_manager_profiles": [],
        "private_endpoints": [],
        "virtual_network_gateways": [],
        "local_network_gateways": [],
        "express_route_circuits": [],
        "route_tables": [],
        "network_interfaces": [],
        "security_groups": [],
        "public_ips": [],
        "other_resources": [],
        "migration_status": {},
        "resource_counts": {},
    }
    icon_private_link = get_icon_uri(
        "Networking", "Private-Link-Service",
        "https://www.azureicons.com/static/images/icons/Networking/svg/Private-Link-Service.svg",
    )

    # --- Resource Group summary ---
    rg_sheet = _find_sheet(sheet_names, ["RG"])
    if rg_sheet:
        rg_records = _sheet_to_records(wb, rg_sheet)
        for r in rg_records:
            name = r.get("Row Labels") or r.get("Resource Group") or ""
            count = r.get("Count of NAME") or r.get("Count") or 0
            if name and "grand total" not in str(name).lower():
                summary["resource_groups"].append({"name": str(name), "count": int(count or 0)})

    # --- Virtual Networks & Subnets ---
    summary["subnets"] = []
    vnet_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Virtual Networks"],
        ["VirtualNetwork", "virtualnetworks"]
    )
    if vnet_sheet:
        vnet_records = _sheet_to_records(wb, vnet_sheet, max_rows=50)
        for r in vnet_records:
            name = _pick(r, "name")
            addr = _pick(r, "addressSpace", "Address Space", "Addressprefixes")
            location = _pick(r, "location")
            peerings = _pick(r, "Virtual Network Peerings", "peerings", "Vnet Peerings")
            
            summary["networks"].append({
                "name": str(name),
                "address_space": str(addr),
                "location": str(location),
                "peerings": peerings or ""
            })

            # Try to safely extract subnets from native structures first.
            subnets_raw = _pick(r, "subnets", "Subnets")
            if isinstance(subnets_raw, str) and subnets_raw.startswith("["):
                try:
                    import ast
                    parsed_subnets = ast.literal_eval(subnets_raw)
                    if isinstance(parsed_subnets, list):
                        for sub in parsed_subnets:
                            if isinstance(sub, dict):
                                props = sub.get("properties") or {}
                                nsg_id = ""
                                if isinstance(props, dict):
                                    nsg_obj = props.get("networkSecurityGroup") or {}
                                    if isinstance(nsg_obj, dict):
                                        nsg_id = str(nsg_obj.get("id") or "")
                                summary["subnets"].append({
                                    "name": sub.get("name", "Unknown_Subnet"),
                                    "vnet": str(name),
                                    "address_prefix": sub.get("addressPrefix", ""),
                                    "nsg": nsg_id,
                                })
                except Exception:
                    pass
            elif isinstance(subnets_raw, list):
                 for sub in subnets_raw:
                      if hasattr(sub, "get"):
                           props = sub.get("properties") or {}
                           nsg_id = ""
                           if isinstance(props, dict):
                               nsg_obj = props.get("networkSecurityGroup") or {}
                               if isinstance(nsg_obj, dict):
                                   nsg_id = str(nsg_obj.get("id") or "")
                           summary["subnets"].append({
                               "name": sub.get("name", "Unknown_Subnet"),
                               "vnet": str(name),
                               "address_prefix": sub.get("addressPrefix", ""),
                               "nsg": nsg_id,
                           })
            else:
                # Fallback for flattened columns from this report format:
                # "Subnet Names", "Subnet Prefixes"
                subnet_names_raw = _pick(r, "Subnet Names", "SubnetNames")
                subnet_prefixes_raw = _pick(r, "Subnet Prefixes", "SubnetPrefixes")
                subnet_nsgs_raw = _pick(r, "Subnet NSG IDs", "Subnet NSG Ids", "SubnetNsgIds", "Subnet N S G I Ds")
                if _is_present(subnet_names_raw):
                    names = [x.strip() for x in re.split(r"[;,]", str(subnet_names_raw)) if x.strip()]
                    prefixes = [x.strip() for x in re.split(r"[;,]", str(subnet_prefixes_raw))] if _is_present(subnet_prefixes_raw) else []
                    nsgs = [x.strip() for x in re.split(r"[;,]", str(subnet_nsgs_raw))] if _is_present(subnet_nsgs_raw) else []
                    for idx, n in enumerate(names):
                        summary["subnets"].append({
                            "name": n,
                            "vnet": str(name),
                            "address_prefix": prefixes[idx] if idx < len(prefixes) else "",
                            "nsg": nsgs[idx] if idx < len(nsgs) else "",
                        })

    # --- Virtual Machines ---
    vm_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Virtual Machines"],
        ["virtualmachines"]
    )
    if vm_sheet:
        vm_records = _sheet_to_records(wb, vm_sheet, max_rows=200)
        for r in vm_records:
            name = _pick(r, "name")
            size = _pick(r, "vmSize", "VM Size", "hardwareProfile", "Vm Size")
            vnet = _pick(r, "VNet", "V Net", "Vnet", "Virtual Network")
            subnet = _pick(r, "Subnet")
            resource_id = _pick(r, "Resource Id", "ResourceId", "id", "resourceid")
            priv_ip = _pick(r, "Private IP", "Private Ip", "private_ip")
            pub_ip = _pick(r, "Public IP", "Public Ip", "public_ip")
            os_type = _pick(r, "osType", "OS Type", "Ostype")
            location = _pick(r, "location")
            rg = _pick(r, "resourceGroup", "Resource Group", "Resourcegroup")
            network_interfaces = _pick(r, "Network Interfaces", "NetworkInterfaces")
            summary["compute"].append({
                "name": str(name),
                "resource_id": str(resource_id),
                "size": str(size),
                "vnet": str(vnet),
                "subnet": str(subnet),
                "private_ip": str(priv_ip),
                "public_ip": str(pub_ip),
                "os": str(os_type),
                "location": str(location),
                "resource_group": str(rg),
                "network_interfaces": network_interfaces,
            })

    # --- Virtual Machine Scale Sets ---
    vmss_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Virtual Machine Scale Sets"],
        ["virtualmachinescalesets", "vmss", "vm scale sets", "vm scale set", "virtual machine scale set", "virtual machine scalesets", "scalesets", "vm-scale-sets"]
    )
    if vmss_sheet:
        vmss_records = _sheet_to_records(wb, vmss_sheet, max_rows=100)
        for r in vmss_records:
            # Attempt to infer VNet/Subnet from embedded Network Interfaces / IP Config JSON
            vnet_val = str(_pick(r, "VNet", "V Net", "Vnet", "Virtual Network"))
            subnet_val = str(_pick(r, "Subnet"))
            if not vnet_val:
                net_raw = _pick(
                    r,
                    "Network Interfaces",
                    "NetworkInterfaces",
                    "NetworkInterfaceConfigurations",
                    "Networkinterfaceconfigurations",
                    "Ip Configurations",
                    "IpConfigurations",
                    "Ip Config",
                    "Properties",
                    "Details",
                    "Connectivity",
                )
                parsed = None
                if isinstance(net_raw, (list, dict)):
                    parsed = net_raw
                elif isinstance(net_raw, str) and net_raw.strip():
                    txt = net_raw.strip()
                    # Try JSON then AST literal
                    try:
                        parsed = json.loads(txt)
                    except Exception:
                        try:
                            parsed = ast.literal_eval(txt)
                        except Exception:
                            parsed = None

                def _find_subnet_id(obj):
                    if obj is None:
                        return None
                    if isinstance(obj, str):
                        s = obj.strip()
                        if s.lower().startswith("/subscriptions/") and "/virtualnetworks/" in s.lower() and "/subnets/" in s.lower():
                            return s
                        return None
                    if isinstance(obj, dict):
                        for v in obj.values():
                            res = _find_subnet_id(v)
                            if res:
                                return res
                    if isinstance(obj, list):
                        for it in obj:
                            res = _find_subnet_id(it)
                            if res:
                                return res
                    return None

                sid = _find_subnet_id(parsed)
                if sid:
                    try:
                        # Extract VNet and Subnet names from resource id
                        parts = sid.split("/")
                        # find index of 'virtualNetworks' and 'subnets'
                        lw = [p.lower() for p in parts]
                        if "virtualnetworks" in lw and "subnets" in lw:
                            vi = lw.index("virtualnetworks")
                            si = lw.index("subnets")
                            if vi + 1 < len(parts):
                                vnet_val = parts[vi + 1]
                            if si + 1 < len(parts):
                                subnet_val = parts[si + 1]
                    except Exception:
                        pass

            summary["vm_scale_sets"].append({
                "name": str(_pick(r, "name")),
                "sku": str(_pick(r, "sku", "SKU")),
                "vnet": str(vnet_val),
                "subnet": str(subnet_val),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourceGroup", "Resource Group", "Resourcegroup")),
                "is_vmss": True,
                "raw_row": r,
            })

    # --- SQL Virtual Machines (map to actual VM rows via Virtual Machine Resource Id) ---
    sql_vm_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Sql Virtual Machines"],
        ["sql virtual machines", "sqlvirtualmachines", "sqlvms", "sql virtual"]
    )
    if sql_vm_sheet:
        for r in _sheet_to_records(wb, sql_vm_sheet, max_rows=200):
            vm_res = _pick(r, "Virtual Machine Resource Id", "VirtualMachineResourceId", "vmResourceId", "Virtual Machine ResourceId", "virtualmachineresourceid")
            if not vm_res:
                continue
            vm_res = str(vm_res).strip()
            # Try to match against compute rows by resource_id or by name suffix
            for vm in summary.get("compute", []):
                rid = str(vm.get("resource_id") or "").strip()
                if not rid:
                    # fallback: match by VM name appearing in resource id
                    if vm.get("name") and vm.get("name").lower() in vm_res.lower():
                        vm["is_sql_vm"] = True
                        vm["sql_row"] = r
                        break
                    continue
                if rid and (rid.lower() == vm_res.lower() or rid.lower().endswith(vm_res.split("/")[-1].lower()) or vm_res.lower().endswith(rid.split("/")[-1].lower())):
                    vm["is_sql_vm"] = True
                    vm["sql_row"] = r
                    break

    # --- AKS / Managed Clusters ---
    aks_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["AKS", "Managed Clusters"],
        ["AKS", "managed cluster", "managedclusters"]
    )
    if aks_sheet:
        seen_aks = set()
        for r in _sheet_to_records(wb, aks_sheet, max_rows=400):
            name = str(_pick(r, "name")).strip()
            if not name:
                continue
            rg = str(_pick(r, "resourcegroup", "Resource Group", "Resource  Group")).strip()
            node_rg = str(_pick(r, "Node Resource Group", "NodeResourceGroup")).strip()
            location = str(_pick(r, "location")).strip()
            network_hint = _infer_network_hint(
                _pick(r, "VNet", "V Net", "Vnet", "Virtual Network"),
                _pick(r, "Connectivity"),
                _pick(r, "Details"),
                _pick(r, "Resource Id", "ResourceId"),
            )
            key = f"{name.lower()}|{rg.lower()}"
            if key in seen_aks:
                continue
            seen_aks.add(key)
            summary["aks_clusters"].append({
                "name": name,
                "location": location,
                "resource_group": rg,
                "node_resource_group": node_rg,
                "type": "AKS",
                "network_hint": str(network_hint),
            })

    # --- Network Interfaces ---
    ni_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Network Interfaces"],
        ["NetworkInterface", "network interfaces"]
    )
    if ni_sheet:
        ni_records = _sheet_to_records(wb, ni_sheet, max_rows=2000)
        for r in ni_records:
            name = str(_pick(r, "name")).strip()
            rg = str(_pick(r, "resourcegroup", "Resource Group", "Resource  Group")).strip()
            ip_cfg = _pick(r, "Ip Configurations", "IP Configurations", "IpConfigurations")
            vnet = _infer_network_hint(
                _pick(r, "VNet", "V Net", "Vnet", "Virtual Network"),
                ip_cfg,
                _pick(r, "resourceid", "Resource Id"),
                _pick(r, "virtualmachineid", "Virtual Machine Id"),
            )
            subnet = ""
            if _is_present(ip_cfg):
                m = re.search(r"/virtualNetworks/[^/]+/subnets/([^/\"'\s]+)", str(ip_cfg), flags=re.IGNORECASE)
                if m:
                    subnet = m.group(1)
            summary["network_interfaces"].append({
                "name": name,
                "resource_group": rg,
                "vnet": str(vnet),
                "subnet": subnet,
                "ip_configurations": ip_cfg,
            })

    # --- NSGs ---
    nsg_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Network Security Groups"],
        ["NetworkSecurityGroup", "NSG", "networksecuritygroups"]
    )
    if nsg_sheet:
        nsg_records = _sheet_to_records(wb, nsg_sheet, max_rows=100)
        for r in nsg_records:
            name = _pick(r, "name")
            location = _pick(r, "location")
            network_interfaces = _pick(r, "Network Interfaces", "Networkinterfaces")
            summary["security_groups"].append({
                "name": str(name),
                "location": str(location),
                "network_interfaces": network_interfaces,
            })

    # --- Load Balancers ---
    lb_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Load Balancers"],
        ["LoadBalancer", "loadbalancers"]
    )
    if lb_sheet:
        lb_records = _sheet_to_records(wb, lb_sheet, max_rows=50)
        for r in lb_records:
            name = _pick(r, "name")
            sku = _pick(r, "sku", "SKU")
            rg = _pick(r, "resourcegroup", "Resource Group", "Resource  Group")
            network_hint = _infer_network_hint(
                _pick(r, "VNet", "Vnet", "Virtual Network"),
                _pick(r, "BackendAddressPools", "Backend Address Pools"),
                _pick(r, "BackendIPConfigurations", "Backend IP Configurations"),
                _pick(r, "FrontendIpConfigurations", "Frontend IP Configurations"),
                _pick(r, "Details"),
                _pick(r, "Connectivity"),
            )
            backend_targets = _infer_lb_backend_targets(
                _pick(r, "BackendAddressPools", "Backend Address Pools"),
                _pick(r, "BackendIPConfigurations", "Backend IP Configurations"),
                _pick(r, "Details"),
                _pick(r, "Connectivity"),
            )
            summary["load_balancers"].append({
                "name": str(name),
                "sku": str(sku),
                "resource_group": str(rg),
                "network_hint": str(network_hint),
                "backend_targets": backend_targets,
            })

    # --- Public IPs ---
    pip_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Public IP Addresses"],
        ["PublicIPAddress", "publicipaddresses"]
    )
    if pip_sheet:
        pip_records = _sheet_to_records(wb, pip_sheet, max_rows=100)
        for r in pip_records:
            name = _pick(r, "name")
            ip = _pick(r, "ipAddress", "IP Address", "Ip Address")
            summary["public_ips"].append({
                "name": str(name),
                "ip": str(ip),
            })

    # --- Application Gateways ---
    appgw_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Application Gateways"],
        ["ApplicationGateway", "application gateways", "applicationgateways"]
    )
    if appgw_sheet:
        for r in _sheet_to_records(wb, appgw_sheet, max_rows=200):
            summary["application_gateways"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- Front Door / AFD Endpoints ---
    afd_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Afdendpoints"],
        ["afdendpoint", "front door", "frontdoor", "afd"]
    )
    if afd_sheet:
        for r in _sheet_to_records(wb, afd_sheet, max_rows=200):
            summary["frontdoor_endpoints"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- Traffic Manager ---
    tm_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Trafficmanagerprofiles"],
        ["traffic manager", "trafficmanager"]
    )
    if tm_sheet:
        for r in _sheet_to_records(wb, tm_sheet, max_rows=200):
            summary["traffic_manager_profiles"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- Private Endpoints ---
    pe_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Private Endpoints"],
        ["private endpoint", "privateendpoints", "private link", "private links", "privatelink", "privatelinks"]
    )
    if pe_sheet:
        for r in _sheet_to_records(wb, pe_sheet, max_rows=300):
            # Capture raw private link connections and standard fields
            linked_raw = _pick(
                r,
                "privateLinkServiceId",
                "Linked Resource",
                "Details",
                "Private Link Service Connections",
                "PrivateLinkServiceConnections",
            )
            # Normalize linked resources into a list
            linked_list = []
            if isinstance(linked_raw, list):
                linked_list = linked_raw
            elif isinstance(linked_raw, dict):
                linked_list = [linked_raw]
            elif isinstance(linked_raw, str) and linked_raw.strip():
                txt = linked_raw.strip()
                try:
                    parsed_lr = json.loads(txt)
                    if isinstance(parsed_lr, list):
                        linked_list = parsed_lr
                    else:
                        linked_list = [parsed_lr]
                except Exception:
                    try:
                        parsed_lr = ast.literal_eval(txt)
                        if isinstance(parsed_lr, list):
                            linked_list = parsed_lr
                        else:
                            linked_list = [parsed_lr]
                    except Exception:
                        # fallback: single string value
                        linked_list = [txt]

            summary["private_endpoints"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
                "vnet": str(_pick(r, "vnet", "virtual network")),
                "subnet": str(_pick(r, "subnet", "Subnet Id", "SubnetId", "subnet id")),
                "linked_resource": str(linked_list[0]) if linked_list else "",
                "linked_resources": linked_list,
                "group_ids": str(_pick(r, "groupIds", "Group IDs")),
                "raw_row": r,
            })

    # --- Virtual Network Gateways (VPN Gateway) ---
    vng_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Virtual Network Gateways"],
        ["virtual network gateway", "virtualnetworkgateway"]
    )
    if vng_sheet:
        for r in _sheet_to_records(wb, vng_sheet, max_rows=200):
            summary["virtual_network_gateways"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
                "gateway_type": str(_pick(r, "gatewaytype", "gateway type")),
            })

    # --- Local Network Gateways ---
    lng_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Local Network Gateways"],
        ["local network gateway", "localnetworkgateway"]
    )
    if lng_sheet:
        for r in _sheet_to_records(wb, lng_sheet, max_rows=200):
            summary["local_network_gateways"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
                "gateway_ip": str(_pick(r, "gatewayipaddress", "gateway ip")),
            })

    # --- ExpressRoute Circuits ---
    er_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Express Route Circuits"],
        ["express route circuit", "expressroute"]
    )
    if er_sheet:
        for r in _sheet_to_records(wb, er_sheet, max_rows=200):
            summary["express_route_circuits"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- Route Tables ---
    rt_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Route Tables"],
        ["route table", "routetables"]
    )
    if rt_sheet:
        for r in _sheet_to_records(wb, rt_sheet, max_rows=400):
            summary["route_tables"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- Network Interfaces ---
    nic_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Network Interfaces"],
        ["network interface", "networkinterfaces"]
    )
    if nic_sheet:
        for r in _sheet_to_records(wb, nic_sheet, max_rows=600):
            summary["network_interfaces"].append({
                "name": str(_pick(r, "name")),
                "location": str(_pick(r, "location")),
                "resource_group": str(_pick(r, "resourcegroup", "resource group")),
            })

    # --- SQL / Database resources ---
    for kw in ["SqlServer", "SQL", "sqlDatabases", "CosmosDB", "DatabaseAccount"]:
        db_sheet = _find_sheet(sheet_names, [kw])
        if db_sheet:
            db_records = _sheet_to_records(wb, db_sheet, max_rows=50)
            for r in db_records:
                name = _pick(r, "name")
                kind = _pick(r, "kind", "Kind", default=kw)
                summary["databases"].append({
                    "name": str(name),
                    "type": str(kind),
                    "raw_row": r,
                })

    # --- Storage Accounts ---
    sa_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Storage Accounts"],
        ["StorageAccount", "storageaccounts", "Storage Account"]
    )
    if sa_sheet:
        sa_records = _sheet_to_records(wb, sa_sheet, max_rows=100)
        for r in sa_records:
            name = _pick(r, "name")
            kind = _pick(r, "kind", "Kind", default="StorageV2")
            summary["storage"].append({
                "name": str(name),
                "kind": str(kind),
            })

    # --- Migration Status (from all High-Level Migration Plan sheets) ---
    summary["migration_details"] = {"Can_Move": [], "Cannot_Move": [], "Action_Required": []}
    
    # Collect data from ALL sheets that look like high-level migration plans
    hl_patterns = ["High-Level Migration Plan", "HighLevel", "High_Level"]
    found_hl_sheets = [s for s in sheet_names if any(p.lower() in s.lower() for p in hl_patterns)]
    
    if found_hl_sheets:
        for hl_sheet in found_hl_sheets:
            hl_records = _sheet_to_records(wb, hl_sheet, max_rows=500)
            for r in hl_records:
                rtype = str(r.get("Resource Type") or r.get("Type") or "")
                status = str(r.get("Move Status") or r.get("Status") or "Unknown")
                count = r.get("Resource Count") or r.get("Count") or 0
                
                if not rtype: continue
                
                # Simple aggregation (add to existing counts if resource type already present)
                detail = {"type": rtype, "count": int(count)}
                
                target_list = None
                if "can move" in status.lower() and "cannot" not in status.lower():
                    target_list = summary["migration_details"]["Can_Move"]
                elif "cannot move" in status.lower():
                    target_list = summary["migration_details"]["Cannot_Move"]
                elif "action required" in status.lower() or "evaluate" in status.lower():
                    target_list = summary["migration_details"]["Action_Required"]
                
                if target_list is not None:
                    # Merge if same type exists
                    existing = next((item for item in target_list if item["type"] == rtype), None)
                    if existing:
                        existing["count"] += detail["count"]
                    else:
                        target_list.append(detail)
    else:
        # Fallback to pure dashboard numbers if HL Plan is missing
        dashboard_sheet = _find_sheet(sheet_names, ["Dashboard"])
        if dashboard_sheet:
            dash_records = _sheet_to_records(wb, dashboard_sheet, max_rows=30)
            for r in dash_records:
                metric = str(r.get("Metric") or r.get("") or "")
                count = r.get("Count") or r.get("Value") or 0
                if "can move" in metric.lower() and "cannot" not in metric.lower():
                    summary["migration_status"]["can_move"] = int(count or 0)
                elif "cannot move" in metric.lower():
                    summary["migration_status"]["cannot_move"] = int(count or 0)
                elif "action required" in metric.lower():
                    summary["migration_status"]["action_required"] = int(count or 0)
                elif "total" in metric.lower() and "resource" in metric.lower():
                    summary["migration_status"]["total"] = int(count or 0)

    # --- Resource counts from AllResources ---
    all_sheet = _find_sheet(sheet_names, ["AllResources", "All Resources"])
    if all_sheet:
        all_records = _sheet_to_records(wb, all_sheet, max_rows=5000)
        type_counts = defaultdict(int)
        seen_major = {
            "aks_clusters": set(),
            "vm_scale_sets": set(),
            "web_apps": set(),
            "static_web_apps": set(),
            "container_apps": set(),
            "storage_accounts": set(),
            "database_services": set(),
            "load_balancers": set(),
            "application_gateways": set(),
            "frontdoor_endpoints": set(),
        }
        for r in all_records:
            # IMPORTANT: Prefer provider path column (`Resource Type`) over friendly `Type`.
            rtype = _pick(r, "Resource Type", "Resource  Type", "type", "TYPE", default="Unknown")
            # Get friendly short name
            rtype_str = str(rtype)
            if "/" in rtype_str:
                rtype_str = rtype_str.split("/")[-1]
            type_counts[rtype_str] += 1

            # Major architecture classification from provider types
            rtype_l = str(rtype).lower()
            rname = str(_pick(r, "name", "Name", "NAME", default="")).strip()
            rloc = str(_pick(r, "location", "Location", default="")).strip()
            rrg = str(_pick(r, "resourcegroup", "Resource Group", "Resource  Group", "Resourcegroup", default="")).strip()

            def add_major(bucket: str):
                if not rname:
                    return
                key = f"{rname.lower()}|{rrg.lower()}|{rloc.lower()}"
                if key in seen_major[bucket]:
                    return
                seen_major[bucket].add(key)
                network_hint = _infer_network_hint(
                    _pick(r, "VNet", "Vnet", "Virtual Network", "VpcId", "SubnetId"),
                    _pick(r, "Connectivity"),
                    _pick(r, "Details"),
                    _pick(r, "ResourceId", "id"),
                )
                backend_targets = _infer_lb_backend_targets(
                    _pick(r, "BackendAddressPools", "Backend Address Pools"),
                    _pick(r, "BackendIPConfigurations", "Backend IP Configurations"),
                    _pick(r, "Connectivity"),
                    _pick(r, "Details"),
                    _pick(r, "ResourceId", "id"),
                )
                summary[bucket].append({
                    "name": rname,
                    "location": rloc,
                    "resource_group": rrg,
                    "type": str(rtype),
                    "network_hint": str(network_hint),
                    "backend_targets": backend_targets,
                })

            if "/managedclusters" in rtype_l:
                add_major("aks_clusters")
            elif "/virtualmachinescalesets" in rtype_l:
                add_major("vm_scale_sets")
            elif "/staticsites" in rtype_l:
                add_major("static_web_apps")
            elif "/containerapps" in rtype_l:
                add_major("container_apps")
            elif "/sites" in rtype_l:
                # Includes Web App / Function App; keeping under web apps bucket for architecture view.
                add_major("web_apps")
            elif "/storageaccounts" in rtype_l:
                add_major("storage_accounts")
            elif (
                "microsoft.sql/" in rtype_l
                or "microsoft.dbfor" in rtype_l
                or "microsoft.documentdb/" in rtype_l
                or "microsoft.cache/redis" in rtype_l
            ):
                add_major("database_services")
            elif "/loadbalancers" in rtype_l:
                add_major("load_balancers")
            elif "/applicationgateways" in rtype_l:
                add_major("application_gateways")
            elif "microsoft.cdn/profiles/afdendpoints" in rtype_l or "/afdendpoints" in rtype_l:
                add_major("frontdoor_endpoints")
        summary["resource_counts"] = dict(sorted(type_counts.items(), key=lambda x: -x[1]))

    # Add sheet names for context
    summary["_available_sheets"] = sheet_names

    # --- AKS network integration inference ---
    # Many assessments do not populate VNet/Subnet fields directly in AKS sheet.
    # Infer from NICs in AKS node resource group when possible.
    if summary.get("aks_clusters") and summary.get("network_interfaces"):
        rg_to_vnet_counts: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
        rg_to_subnet_counts: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
        for ni in summary.get("network_interfaces", []):
            rg = str(ni.get("resource_group") or "").strip().lower()
            vnet = str(ni.get("vnet") or "").strip()
            subnet = str(ni.get("subnet") or "").strip()
            if not rg:
                continue
            if vnet:
                rg_to_vnet_counts[rg][vnet] += 1
            if subnet:
                rg_to_subnet_counts[rg][subnet] += 1

        for aks in summary.get("aks_clusters", []):
            node_rg = str(aks.get("node_resource_group") or "").strip().lower()
            if not node_rg:
                continue

            # Fill network_hint if empty
            if not _is_present(aks.get("network_hint")) and rg_to_vnet_counts.get(node_rg):
                vnet = max(rg_to_vnet_counts[node_rg].items(), key=lambda kv: kv[1])[0]
                aks["network_hint"] = vnet

            # Add subnet hint for placement in detailed diagrams
            if not _is_present(aks.get("subnet_hint")) and rg_to_subnet_counts.get(node_rg):
                subnet = max(rg_to_subnet_counts[node_rg].items(), key=lambda kv: kv[1])[0]
                aks["subnet_hint"] = subnet

    def _dedupe_by_name_rg(items: List[dict]) -> List[dict]:
        out = []
        seen = set()
        for item in items or []:
            key = (
                str(item.get("name") or "").strip().lower(),
                str(item.get("resource_group") or "").strip().lower(),
            )
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
        return out

    for bucket in [
        "aks_clusters",
        "vm_scale_sets",
        "web_apps",
        "static_web_apps",
        "container_apps",
        "storage_accounts",
        "database_services",
        "load_balancers",
        "application_gateways",
        "frontdoor_endpoints",
    ]:
        summary[bucket] = _dedupe_by_name_rg(summary.get(bucket, []))

    # ---- Subscription mapping (for subscription swim-lanes in the diagram) ----
    # AllResources maps every resource name → its subscription ID; a few sheets
    # (e.g. SQL VMs) also carry a friendly "Subscription Name" we can reuse.
    name_to_sub = {}       # resource name (lower) → subscription id
    sub_id_to_name = {}    # subscription id (lower) → friendly name
    all_sheet = _find_sheet_prefer_exact(sheet_names, ["AllResources"], ["allresources"])
    if all_sheet:
        for r in _sheet_to_records(wb, all_sheet, max_rows=5000):
            nm = str(_pick(r, "Name", "name") or "").strip().lower()
            sid = str(_pick(r, "Subscription", "subscriptionId", "Subscriptionid") or "").strip().lower()
            if nm and sid:
                name_to_sub.setdefault(nm, sid)
    # Harvest id→name from any sheet that carries both columns.
    for sh in sheet_names:
        try:
            recs = _sheet_to_records(wb, sh, max_rows=300)
        except Exception:
            continue
        got = False
        for r in recs:
            sid = str(_pick(r, "Subscriptionid", "subscriptionId", "Subscription Id") or "").strip().lower()
            snm = str(_pick(r, "Subscription Name", "subscriptionName") or "").strip()
            if sid and len(sid) >= 30 and snm:
                sub_id_to_name.setdefault(sid, snm)
                got = True
        # keep scanning; cheap enough, and different sheets cover different subs
    summary["subscription_names"] = sub_id_to_name
    for net in summary.get("networks", []):
        sid = name_to_sub.get(str(net.get("name") or "").strip().lower(), "")
        net["subscription_id"] = sid
        net["subscription"] = sub_id_to_name.get(sid, "") or (f"Subscription {sid[:8]}" if sid else "Unknown subscription")

    return summary


# ---------------------------------------------------------------------------
# 1b. AWS Data Extraction
# ---------------------------------------------------------------------------

def _extract_aws_summary(wb, sheet_names: list) -> dict:
    """Extract AWS environment summary."""
    summary = {
        "resource_groups": [],
        "networks": [],
        "compute": [],
        "databases": [],
        "storage": [],
        "load_balancers": [],
        "internet_gateways": [],
        "vpc_endpoints": [],
        "route_tables": [],
        "security_groups": [],
        "other_resources": [],
        "resource_counts": {},
    }

    # --- Summary / Resource counts ---
    sum_sheet = _find_sheet_prefer_exact(sheet_names, ["Summary"], ["Summary"])
    if sum_sheet:
        sum_records = _sheet_to_records(wb, sum_sheet)
        for r in sum_records:
            name = _pick(r, "Friendly Name", "Resource Type")
            count = _pick(r, "Count", default=0)
            if name and count:
                summary["resource_counts"][str(name)] = int(count)

    # --- EC2 Instances ---
    ec2_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 Instances"],
        ["EC2 Instance"]
    )
    if ec2_sheet:
        ec2_records = _sheet_to_records(wb, ec2_sheet, max_rows=200)
        for r in ec2_records:
            name = _pick(r, "ResourceName", "Name")
            itype = _pick(r, "InstanceType", "Azure SKU")
            vpc = _pick(r, "VpcId")
            subnet = _pick(r, "SubnetId")
            priv_ip = _pick(r, "PrivateIpAddress", "PrivateIp")
            pub_ip = _pick(r, "PublicIpAddress", "PublicIp")
            conn = _pick(r, "Connectivity")
            summary["compute"].append({
                "name": str(name),
                "type": str(itype),
                "vpc": str(vpc),
                "subnet": str(subnet),
                "private_ip": str(priv_ip),
                "public_ip": str(pub_ip),
                "connectivity": str(conn),
            })

    # --- VPCs ---
    vpc_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["VPCs"],
        ["VPC"]
    )
    if vpc_sheet:
        vpc_records = _sheet_to_records(wb, vpc_sheet, max_rows=50)
        for r in vpc_records:
            name = _pick(r, "ResourceName", "Name", "VpcId")
            cidr = _pick(r, "CidrBlock", "CIDR")
            summary["networks"].append({
                "name": str(name),
                "vpc_id": str(_pick(r, "VpcId")),
                "cidr": str(cidr),
            })

    # --- Subnets ---
    subnet_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 Subnets"],
        ["Subnet"]
    )
    if subnet_sheet:
        sub_records = _sheet_to_records(wb, subnet_sheet, max_rows=100)
        for r in sub_records:
            # Capture BOTH name and subnet_id. EC2 instances reference subnets
            # by ID (subnet-xxxxx) but the friendly Name column is what we
            # want to display, so we store both and let the renderer match
            # on subnet_id while labelling with name.
            name = _pick(r, "ResourceName", "Name")
            subnet_id = _pick(r, "SubnetId")
            if not name:
                name = subnet_id  # fallback when no name tag
            cidr = _pick(r, "CidrBlock")
            vpc = _pick(r, "VpcId")
            az = _pick(r, "AvailabilityZone")
            summary.setdefault("subnets", []).append({
                "name": str(name),
                "subnet_id": str(subnet_id),
                "cidr": str(cidr),
                "vpc": str(vpc),
                "az": str(az),
            })

    # --- Security Groups ---
    sg_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Security Groups"],
        ["Security Group"]
    )
    if sg_sheet:
        sg_records = _sheet_to_records(wb, sg_sheet, max_rows=100)
        for r in sg_records:
            name = _pick(r, "ResourceName", "Name", "GroupName")
            vpc = _pick(r, "VpcId")
            summary["security_groups"].append({
                "name": str(name),
                "vpc": str(vpc),
            })

    # --- RDS ---
    rds_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["RDS", "RDS Instances"],
        ["RDS", "RDS Database"]
    )
    if rds_sheet:
        rds_records = _sheet_to_records(wb, rds_sheet, max_rows=50)
        for r in rds_records:
            name = _pick(r, "ResourceName", "Name")
            engine = _pick(r, "Engine")
            multi_az = _pick(r, "MultiAZ")
            summary["databases"].append({
                "name": str(name),
                "engine": str(engine),
                "multi_az": str(multi_az),
            })

    # --- S3 ---
    s3_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["S3 Buckets"],
        ["S3", "S3 Bucket"]
    )
    if s3_sheet:
        s3_records = _sheet_to_records(wb, s3_sheet, max_rows=100)
        for r in s3_records:
            name = _pick(r, "ResourceName", "Name", "BucketName")
            summary["storage"].append({"name": str(name), "type": "S3"})

    # --- Load Balancers ---
    lb_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Load Balancers"],
        ["Load Balancer"]
    )
    if lb_sheet:
        lb_records = _sheet_to_records(wb, lb_sheet, max_rows=50)
        for r in lb_records:
            name = _pick(r, "ResourceName", "Name", "LoadBalancerName")
            # Capture topology-relevant fields so the renderer can place the
            # LB inside its VPC and draw arrows to its target instances.
            summary["load_balancers"].append({
                "name": str(name),
                "type": str(_pick(r, "Type")),
                "scheme": str(_pick(r, "Scheme")),
                "vpc": str(_pick(r, "VpcId", "Vpc Id")),
                "dns_name": str(_pick(r, "DNSName", "DnsName", "DNS Name")),
                "availability_zones": str(_pick(r, "AvailabilityZones", "Availability Zones")),
                # Target info — column names vary by extraction tool. We grab
                # the most common ones and let the renderer parse instance IDs.
                "targets": str(_pick(
                    r, "Targets", "TargetInstanceIds", "BackendInstances",
                    "Instances", "TargetGroupArns", "TargetGroups",
                )),
                "raw_row": r,  # fallback: regex-scan the row for i-* IDs
            })

    # --- Internet Gateways ---
    igw_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 Internet Gateways"],
        ["Internet Gateway", "EC2 Internet Gateway"]
    )
    if igw_sheet:
        igw_records = _sheet_to_records(wb, igw_sheet, max_rows=300)
        for r in igw_records:
            summary["internet_gateways"].append({
                "name": str(_pick(r, "ResourceName", "Name", "InternetGatewayId")),
                "location": str(_pick(r, "Location")),
            })

    # --- VPC Endpoints ---
    vpce_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 VPC Endpoints"],
        ["VPC Endpoint", "VPCEndpoint"]
    )
    if vpce_sheet:
        vpce_records = _sheet_to_records(wb, vpce_sheet, max_rows=400)
        for r in vpce_records:
            summary["vpc_endpoints"].append({
                "name": str(_pick(r, "ResourceName", "Name", "Id")),
                "vpc": str(_pick(r, "VpcId")),
                "service_name": str(_pick(r, "ServiceName")),
            })

    # --- Route Tables ---
    rt_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 Route Tables"],
        ["Route Table", "Route Tables"]
    )
    if rt_sheet:
        rt_records = _sheet_to_records(wb, rt_sheet, max_rows=400)
        for r in rt_records:
            summary["route_tables"].append({
                "name": str(_pick(r, "ResourceName", "Name", "RouteTableId")),
                "vpc": str(_pick(r, "VpcId")),
            })

    # --- NAT Gateways ---
    natgw_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 NAT Gateways", "NAT Gateways", "Nat Gateways"],
        ["NAT Gateway", "NatGateway"]
    )
    if natgw_sheet:
        natgw_records = _sheet_to_records(wb, natgw_sheet, max_rows=200)
        summary.setdefault("nat_gateways", [])
        for r in natgw_records:
            summary["nat_gateways"].append({
                "name": str(_pick(r, "ResourceName", "Name", "NatGatewayId")),
                "vpc": str(_pick(r, "VpcId")),
                "subnet": str(_pick(r, "SubnetId")),
                "state": str(_pick(r, "State", "ConnectivityType")),
                "public_ip": str(_pick(r, "PublicIp")),
            })

    # --- Auto Scaling Groups ---
    asg_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["Auto Scaling Groups", "AutoScaling Groups"],
        ["Auto Scaling Group", "AutoScalingGroup", "AutoScaling"]
    )
    if asg_sheet:
        asg_records = _sheet_to_records(wb, asg_sheet, max_rows=200)
        summary.setdefault("auto_scaling_groups", [])
        for r in asg_records:
            summary["auto_scaling_groups"].append({
                "name": str(_pick(r, "ResourceName", "Name", "AutoScalingGroupName")),
                "min_size": str(_pick(r, "MinSize")),
                "max_size": str(_pick(r, "MaxSize")),
                "desired": str(_pick(r, "DesiredCapacity")),
                "launch_template": str(_pick(r, "LaunchTemplate", "LaunchConfigurationName")),
            })

    # --- Elastic IPs ---
    eip_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EC2 Elastic IPs"],
        ["Elastic IP", "ElasticIP", "EC2 EIP"]
    )
    if eip_sheet:
        eip_records = _sheet_to_records(wb, eip_sheet, max_rows=500)
        summary.setdefault("elastic_ips", [])
        for r in eip_records:
            summary["elastic_ips"].append({
                "name": str(_pick(r, "ResourceName", "Name", "PublicIp", "AllocationId")),
                "public_ip": str(_pick(r, "PublicIp")),
                "attached_to": str(_pick(r, "Attached To", "AttachedTo", "InstanceId", "NetworkInterfaceId")),
                "attached": str(_pick(r, "Attached")),
            })

    # --- EBS Volumes ---
    vol_sheet = _find_sheet_prefer_exact(
        sheet_names,
        ["EBS Volumes"],
        ["EBS Volume", "Volume"]
    )
    if vol_sheet:
        vol_records = _sheet_to_records(wb, vol_sheet, max_rows=2000)
        summary.setdefault("volumes", [])
        for r in vol_records:
            summary["volumes"].append({
                "name": str(_pick(r, "ResourceName", "Name", "VolumeId")),
                "size_gb": str(_pick(r, "Size", "SizeGB")),
                "type": str(_pick(r, "VolumeType")),
                "az": str(_pick(r, "AvailabilityZone")),
                "state": str(_pick(r, "State")),
            })

    # --- ECS Clusters / Services ---
    ecs_sheet = _find_sheet_prefer_exact(sheet_names, ["ECS Clusters"], ["ECS Cluster"])
    if ecs_sheet:
        ecs_records = _sheet_to_records(wb, ecs_sheet, max_rows=200)
        summary.setdefault("ecs_clusters", [])
        for r in ecs_records:
            summary["ecs_clusters"].append({
                "name": str(_pick(r, "ResourceName", "Name", "clusterName")),
                "status": str(_pick(r, "Status")),
                "active_services_count": str(_pick(r, "ActiveServicesCount")),
            })

    ecs_svc_sheet = _find_sheet_prefer_exact(sheet_names, ["ECS Services"], ["ECS Service"])
    if ecs_svc_sheet:
        ecs_svc_records = _sheet_to_records(wb, ecs_svc_sheet, max_rows=300)
        summary.setdefault("ecs_services", [])
        for r in ecs_svc_records:
            summary["ecs_services"].append({
                "name": str(_pick(r, "ResourceName", "Name", "serviceName")),
                "cluster": str(_pick(r, "ClusterArn", "Cluster")),
                "launch_type": str(_pick(r, "LaunchType")),
            })

    # --- EKS Clusters ---
    eks_sheet = _find_sheet_prefer_exact(sheet_names, ["EKS Clusters"], ["EKS Cluster"])
    if eks_sheet:
        eks_records = _sheet_to_records(wb, eks_sheet, max_rows=100)
        summary.setdefault("eks_clusters", [])
        for r in eks_records:
            summary["eks_clusters"].append({
                "name": str(_pick(r, "ResourceName", "Name")),
                "version": str(_pick(r, "Version")),
                "status": str(_pick(r, "Status")),
            })

    # --- ECR Repositories ---
    ecr_sheet = _find_sheet_prefer_exact(sheet_names, ["ECR Repositories"], ["ECR Repository"])
    if ecr_sheet:
        ecr_records = _sheet_to_records(wb, ecr_sheet, max_rows=500)
        summary.setdefault("ecr_repositories", [])
        for r in ecr_records:
            summary["ecr_repositories"].append({
                "name": str(_pick(r, "ResourceName", "Name", "repositoryName")),
            })

    # --- Lambda Functions ---
    lambda_sheet = _find_sheet_prefer_exact(sheet_names, ["Lambda Functions"], ["Lambda Function", "Lambda"])
    if lambda_sheet:
        lambda_records = _sheet_to_records(wb, lambda_sheet, max_rows=500)
        summary.setdefault("lambda_functions", [])
        for r in lambda_records:
            summary["lambda_functions"].append({
                "name": str(_pick(r, "ResourceName", "Name", "FunctionName")),
                "runtime": str(_pick(r, "Runtime")),
            })

    # --- RDS DB Clusters (Aurora/DocDB/Neptune) ---
    rds_cluster_sheet = _find_sheet_prefer_exact(
        sheet_names, ["RDS Clusters", "RDS clusters"], ["RDS Cluster", "DBCluster"]
    )
    if rds_cluster_sheet:
        rdsc_records = _sheet_to_records(wb, rds_cluster_sheet, max_rows=200)
        summary.setdefault("rds_clusters", [])
        for r in rdsc_records:
            summary["rds_clusters"].append({
                "name": str(_pick(r, "ResourceName", "Name", "DBClusterIdentifier")),
                "engine": str(_pick(r, "Engine")),
                "status": str(_pick(r, "Status")),
            })

    # --- DynamoDB Tables ---
    ddb_sheet = _find_sheet_prefer_exact(sheet_names, ["DynamoDB Tables"], ["DynamoDB", "DDB"])
    if ddb_sheet:
        ddb_records = _sheet_to_records(wb, ddb_sheet, max_rows=500)
        summary.setdefault("dynamodb_tables", [])
        for r in ddb_records:
            summary["dynamodb_tables"].append({
                "name": str(_pick(r, "ResourceName", "Name", "TableName")),
            })

    # --- Classic Load Balancers (ELB v1) ---
    clb_sheet = _find_sheet_prefer_exact(sheet_names, ["Classic Load Balancers"], ["Classic Load Balancer"])
    if clb_sheet:
        clb_records = _sheet_to_records(wb, clb_sheet, max_rows=200)
        for r in clb_records:
            summary["load_balancers"].append({
                "name": str(_pick(r, "ResourceName", "Name", "LoadBalancerName")),
                "type": "Classic",
                "scheme": str(_pick(r, "Scheme")),
            })

    return summary


# ---------------------------------------------------------------------------
# 1c. GCP Data Extraction
# ---------------------------------------------------------------------------

def _extract_gcp_summary(wb, sheet_names: list) -> dict:
    """Extract GCP environment summary."""
    summary = {
        "resource_groups": [],
        "networks": [],
        "compute": [],
        "databases": [],
        "storage": [],
        "other_resources": [],
        "resource_counts": {},
    }

    # --- Summary ---
    sum_sheet = _find_sheet(sheet_names, ["Summary"])
    if sum_sheet:
        sum_records = _sheet_to_records(wb, sum_sheet)
        for r in sum_records:
            name = r.get("Metric") or r.get("Resource Type") or ""
            count = r.get("Value") or r.get("Count") or 0
            if name and count and "---" not in str(name):
                try:
                    summary["resource_counts"][str(name)] = int(count)
                except (ValueError, TypeError):
                    pass

    # --- All Resources ---
    all_sheet = _find_sheet(sheet_names, ["All Resources"])
    if all_sheet:
        all_records = _sheet_to_records(wb, all_sheet, max_rows=500)
        for r in all_records:
            rtype = r.get("Resource Type") or ""
            name = r.get("Name") or ""
            machine = r.get("Machine Type / Class") or ""
            location = r.get("Location") or ""
            azure_svc = r.get("Azure Service") or ""
            azure_sku = r.get("Azure VM SKU") or ""

            entry = {
                "name": str(name),
                "type": str(rtype),
                "location": str(location),
                "azure_service": str(azure_svc),
            }

            if rtype == "Instance":
                entry["machine_type"] = str(machine)
                entry["azure_sku"] = str(azure_sku)
                summary["compute"].append(entry)
            elif rtype == "Network":
                summary["networks"].append(entry)
            elif rtype == "Bucket":
                summary["storage"].append(entry)
            elif rtype in ("Instance", "Cluster") and "sql" in str(r.get("GCP Service", "")).lower():
                summary["databases"].append(entry)
            else:
                summary["other_resources"].append(entry)

    return summary


# ---------------------------------------------------------------------------
# 1d. Generic Fallback
# ---------------------------------------------------------------------------

def _extract_generic_summary(wb, sheet_names: list) -> dict:
    """Fallback: read Summary + AllResources sheets."""
    summary = {"resource_counts": {}, "networks": [], "compute": [], "other_resources": []}
    sum_sheet = _find_sheet(sheet_names, ["Summary", "Dashboard"])
    if sum_sheet:
        for r in _sheet_to_records(wb, sum_sheet, max_rows=50):
            vals = list(r.values())
            if len(vals) >= 2 and vals[1]:
                try:
                    summary["resource_counts"][str(vals[0])] = int(vals[1])
                except (ValueError, TypeError):
                    pass
    return summary


# ---------------------------------------------------------------------------
# 2. Compact Summary — Trim for token budget
# ---------------------------------------------------------------------------

def compact_summary(summary: dict, max_items: int = 80) -> dict:
    """
    Trim the summary to fit within GPT token limits.
    For large environments, group similar resources.
    """
    compact = {
        "cloud": summary.get("cloud", "unknown"),
        "resource_counts": summary.get("resource_counts", {}),
        "migration_status": summary.get("migration_status", {})
    }
    
    if "migration_details" in summary:
        compact["migration_details"] = summary.get("migration_details", {})

    # Networks: keep all (usually < 20)
    compact["networks"] = summary.get("networks", [])[:20]

    # Compute: summarize if > max_items
    compute = summary.get("compute", [])
    if len(compute) > max_items:
        # Group by subnet/vnet
        by_group = defaultdict(list)
        for vm in compute:
            key = vm.get("vnet") or vm.get("vpc") or vm.get("subnet") or "default"
            by_group[key].append(vm)

        compact["compute_summary"] = []
        for group, vms in by_group.items():
            sizes = defaultdict(int)
            for vm in vms:
                sizes[vm.get("size") or vm.get("type") or "unknown"] += 1
            compact["compute_summary"].append({
                "group": group,
                "count": len(vms),
                "sizes": dict(sizes),
                "sample_names": [vm["name"] for vm in vms[:3]],
            })
    else:
        compact["compute"] = compute

    # Databases, storage, LBs, SGs: keep concise
    for key in [
        "databases",
        "storage",
        "load_balancers",
        "internet_gateways",
        "vpc_endpoints",
        "application_gateways",
        "frontdoor_endpoints",
        "traffic_manager_profiles",
        "private_endpoints",
        "virtual_network_gateways",
        "local_network_gateways",
        "express_route_circuits",
        "route_tables",
        "network_interfaces",
        "security_groups",
        "public_ips",
        "subnets",
    ]:
        items = summary.get(key, [])
        if items:
            compact[key] = items[:50]

    # Resource groups
    rgs = summary.get("resource_groups", [])
    if rgs:
        compact["resource_groups"] = rgs[:30]

    return compact


# ---------------------------------------------------------------------------
# 3. GPT Prompt Builder
# ---------------------------------------------------------------------------

NETWORK_PROMPT = """You are a cloud architecture diagram expert. Based on the following {cloud} environment data, generate a Mermaid diagram showing the ACTUAL NETWORK TOPOLOGY architecture.

Please mimic a modern Enterprise Architecture diagram based ONLY on the provided JSON data.

Rules:
1. STRICTLY use ONLY the resources listed in the JSON. DO NOT invent, hallucinate, or add example Load Balancers, Databases, or Storage if they are not in the data!
2. Use `graph LR` (Left-to-Right) layout.
3. PRIORITIZE a VNet/VPC-centric layout (similar to enterprise network diagrams), not a Public-IP-centric layout.
4. Structure the hierarchy explicitly using `subgraph` blocks:
   - Create an explicit node for the User/Internet on the far left: `User["👤 User"] --> Internet["🌐 Internet"]`
   - Create an outer subgraph called "Cloud Environment".
   - Inside Cloud Environment, create subgraphs for each ACTUAL VNet/VPC present in the data.
   - Inside VNets/VPCs, create subgraphs for ACTUAL Subnets.
   - Place Compute (VMs) inside their respective Subnets.
   - Group Storage and Databases into a standalone subgraph positioned to the right (ONLY if they exist in the data).
5. Show realistic connections based on the data:
   - VNET Peerings: Use labeled edge syntax, e.g. `VnetA ---|Peering| VnetB`.
   - Subnet/VNet hierarchy MUST be visually primary.
   - Compute: Link VMs to their subnets.
   - External ingress: Link Internet to edge entry points (e.g., Load Balancer/App Gateway/Public IP) only when present.
   - If a VM has a public IP, represent it as `Internet --> PublicIP --> VM` (Public IP as an attachment to VM), not as the main organizing structure.
   - If public IP data is missing/unclear, omit Public IP nodes and keep focus on VNet/Subnet/VM + Peering.
   - Do NOT force Internet/Public IP links for every VM.
   - Node IDs MUST be a single ALPHANUMERIC word (e.g., V1, SN1, HubVnet).
   - ALWAYS use the ID["🖥️ Label"] syntax for every node.
   - ALWAYS wrap labels in double quotes.
   - Use Mermaid label syntax for edges: `ID1 -->|Label| ID2` or `ID1 ---|Label| ID2`.
   - ONE connection per line ONLY (NO chaining like A -- B -- C).
   - Mermaid keywords (graph, subgraph, end) MUST be in lowercase.
   - Always end each relationship line with a literal newline character.
6. Peering emphasis:
   - Always render VNet peering links when peering information exists in the data.
   - Keep peering links visually distinct using labels like `|Peering|`.
7. Resource Labeling:
   - For large groups (>5), summarize: `node1["🖥️ web-vm-01..05 (5x D4s_v5)"]`.
   - Prefix labels with emojis: 🖥️ (Compute), 🗄️ (Databases), 🪣 (Storage), ⚖️ (Load Balancers), 🛡️ (Security Groups).
8. Keep the diagram readable — max 80 nodes.
9. Do NOT include any markdown formatting or code fences — output ONLY the raw Mermaid code.

Environment Data:
```json
{data}
```

Generate the Mermaid diagram now:"""

RESOURCE_GROUP_PROMPT = """You are a cloud architecture diagram expert. Based on the following {cloud} environment data, generate a Mermaid diagram showing the RESOURCE GROUP layout.

Please mimic a modern Enterprise Architecture diagram.

Rules:
1. Use `graph LR` (Left-to-Right) layout.
2. The core layout should have a central root node `CloudOverview["☁️ Cloud Subscription"]` pointing to the respective Resource Groups.
3. Each Resource Group must be an explicit `subgraph` named after the Resource Group.
4. Inside each RG subgraph, generate individual nodes for each type of resource (Compute, Storage, Network, etc).
5. Append the count for each resource type directly inside the label and prefix with an emoji:
   - 🖥️ for VMs 
   - 🗄️ for Databases
   - 🪣 for Storage Accounts
   - 🌐 for Virtual Networks / Public IPs
6. CRITICAL: Node labels containing parentheses, brackets, or special characters MUST be wrapped in double quotes. 
   CORRECT: `node1["🖥️ VMs (12)"]`
   INCORRECT: `node1[🖥️ VMs (12)]`
7. Add proper styling using `classDef`:
   - `classDef rg fill:#333333,stroke:#555555,color:white` (Apply to subgraphs if possible, otherwise skip)
8. Do NOT include any markdown formatting or code fences — output ONLY the raw Mermaid code.

Environment Data:
```json
{data}
```

Generate the Mermaid diagram now:"""

MIGRATION_PROMPT = """You are a cloud migration expert. Based on the following {cloud} environment data with migration status, generate a Mermaid diagram showing MIGRATION READINESS.

Please mimic a modern Enterprise Dashboard diagram.

Rules:
1. Use `graph TD` (Top-Down) layout.
2. Add a central root node: `Root["📊 Total Resources Analyzed"]`.
3. Create 3 distinct `subgraph` blocks to visually categorize readiness:
   - `subgraph GreenPhase ["✅ Ready to Move"]`
   - `subgraph YellowPhase ["⚠️ Action Required"]`
   - `subgraph RedPhase ["❌ Cannot Move"]`
4. The Root node should point to each of the 3 subgraphs (or dummy entry nodes inside them).
5. STRICTLY Use the data from `migration_details` array to populate each subgraph! For example, if "Can_Move" has "Virtual Machines": 12, then draw `node["🖥️ Virtual Machines (12)"]` inside the GreenPhase subgraph. DO NOT guess or hallucinate resources.
6. Prefix labels with emojis:
   - 🖥️ for Compute, 🗄️ for Databases, 🌐 for Networking, 🪣 for Storage, ⚙️ for generic resources.
7. CRITICAL: Node labels containing parentheses, brackets, or special characters MUST be wrapped in double quotes. 
   CORRECT: `node1["🖥️ VMs (12)"]`
   INCORRECT: `node1[🖥️ VMs (12)]`
8. Use `classDef` to strictly apply the correct status colors to the nodes:
   - `classDef ready fill:#10b981,stroke:#047857,color:white` (apply to GreenPhase nodes)
   - `classDef action fill:#f59e0b,stroke:#b45309,color:white` (apply to YellowPhase nodes)
   - `classDef block fill:#ef4444,stroke:#b91c1c,color:white` (apply to RedPhase nodes)
9. Keep it balanced and visually impactful. Limit to max 60 total nodes.
10. Do NOT include any markdown formatting or code fences — output ONLY the raw Mermaid code.

Environment Data:
```json
{data}
```

Generate the Mermaid diagram now:"""

PROMPTS = {
    "network": NETWORK_PROMPT,
    "resource_groups": RESOURCE_GROUP_PROMPT,
    "migration": MIGRATION_PROMPT,
}


def build_prompt(summary: dict, diagram_type: str = "network") -> str:
    """Build the GPT prompt for diagram generation."""
    template = PROMPTS.get(diagram_type, NETWORK_PROMPT)
    compact = compact_summary(summary)
    data_json = json.dumps(compact, indent=2, default=str)

    # Truncate if too long (keep under ~6000 tokens ≈ 24KB)
    if len(data_json) > 20000:
        data_json = data_json[:20000] + "\n... (truncated for token limit)"

    return template.format(
        cloud=summary.get("cloud", "cloud").upper(),
        data=data_json,
    )


# ---------------------------------------------------------------------------
# 4. GPT Diagram Generation
# ---------------------------------------------------------------------------

def generate_architecture_diagram(
    file_path: str,
    diagram_type: str = "network",
    cloud_type: str = None,
) -> dict:
    """
    Main entry point: Extract data from Excel → build native draw.io XML.

    draw.io is the ONLY supported output. Mermaid generation has been removed.

    Returns:
        {
            "drawio_xml": "<draw.io mxGraph XML>",
            "summary": { ... },
            "diagram_type": "network",
            "cloud": "azure"
        }
    """
    # 1. Extract environment summary
    print(f"🏗️ Extracting environment data from {os.path.basename(file_path)}...")
    summary = extract_environment_summary(file_path, cloud_type)

    # 2. Build native draw.io XML. Two Azure views are supported:
    #      - "resource_groups" → resources grouped by Resource Group
    #      - anything else ("network"/"complete") → full nested architecture
    #        (Region → VNet → Subnet → resources) WITH VNet peering arrows.
    normalized_type = str(diagram_type or "network").strip().lower().replace("-", "_")
    cloud = str(summary.get("cloud") or "").lower()
    drawio_xml = ""
    try:
        if cloud == "azure" and normalized_type in {"resource_groups", "resourcegroups", "rg", "rgwise", "resource_group"}:
            print("🧭 Building draw.io resource-group view...")
            drawio_xml = _build_azure_drawio_rgwise(summary)
        elif cloud == "azure":
            print("🧭 Building complete draw.io architecture (with peering arrows)...")
            drawio_xml = _build_azure_drawio_nested(summary, arrows=True)
        else:
            print("🧭 Building draw.io network topology diagram...")
            drawio_xml = _build_drawio_network_xml(summary)
    except Exception as e:
        print(f"⚠️ draw.io generation failed: {e}")
        drawio_xml = ""

    return {
        "drawio_xml": drawio_xml,
        "summary": compact_summary(summary),
        "diagram_type": diagram_type,
        "cloud": summary.get("cloud", "unknown"),
    }


def _safe_mermaid_id(value: str, fallback: str) -> str:
    """Create Mermaid-safe alphanumeric node IDs."""
    raw = re.sub(r"[^A-Za-z0-9]", "", str(value or ""))
    if not raw:
        raw = fallback
    if not raw[0].isalpha():
        raw = f"N{raw}"
    return raw


def _is_present(value) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return text != "" and text.lower() not in {"none", "nan", "null", "n/a", "-"}


def _parse_peering_targets(peerings_value) -> List[str]:
    """Extract likely peered VNet names from mixed text/list payloads."""
    if not _is_present(peerings_value):
        return []

    text = str(peerings_value)

    # Try common serialized list/dict shapes first.
    names = re.findall(r"(?:remoteVirtualNetwork|remoteVnet|remoteNetwork|name)\s*['\":= ]+\s*([A-Za-z0-9._-]+)", text)
    if names:
        return list(dict.fromkeys(names))

    # Fallback: split by separators.
    parts = re.split(r"[,\n;|]+", text)
    cleaned = []
    for p in parts:
        token = p.strip().strip("'\"")
        if token and len(token) > 1:
            cleaned.append(token)
    return list(dict.fromkeys(cleaned))


def _has_azure_dns(summary: Dict) -> bool:
    """Best-effort check whether Azure DNS resources exist in the assessment."""
    for s in summary.get("_available_sheets", []) or []:
        ls = str(s).lower()
        if "dns zone" in ls or "private dns" in ls:
            return True
    for k in (summary.get("resource_counts") or {}).keys():
        if "dns" in str(k).lower():
            return True
    return False


def _infer_network_hint(*values) -> str:
    """
    Infer VNet/VPC name/id from mixed text fields.
    Supports Azure resource IDs and generic Connectivity/Details text.
    """
    text = " | ".join([str(v) for v in values if v is not None and str(v).strip() != ""])
    if not text:
        return ""

    # Azure subnet/vnet resource id
    m = re.search(r"/virtualNetworks/([^/]+)/subnets/", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    # Azure explicit tokens
    m = re.search(r"\bVNet\s*[:=]\s*([A-Za-z0-9._-]+)", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    # AWS explicit tokens
    m = re.search(r"\bVPC\s*[:=]\s*([A-Za-z0-9._-]+)", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    # Raw vnet/vpc ids
    m = re.search(r"\b(vnet-[A-Za-z0-9._-]+|vpc-[A-Za-z0-9a-f]+)\b", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)

    return ""


def _infer_lb_backend_targets(*values) -> List[str]:
    """
    Extract likely VM backend targets from LB-related payloads.
    """
    text = " | ".join([str(v) for v in values if v is not None and str(v).strip() != ""])
    if not text:
        return []

    targets: List[str] = []
    # Azure VMSS and VM resource IDs.
    targets += re.findall(r"/virtualMachineScaleSets/([^/\"'\s]+)", text, flags=re.IGNORECASE)
    targets += re.findall(r"/virtualMachines/([^/\"'\s]+)", text, flags=re.IGNORECASE)
    # Azure NIC ids often carry node names for AKS pools.
    targets += re.findall(r"/networkInterfaces/([^/\"'\s]+)", text, flags=re.IGNORECASE)
    # Generic key-value hints.
    targets += re.findall(r"\b(?:vm|virtualmachine|target|backend)\s*[:=]\s*([A-Za-z0-9._-]+)", text, flags=re.IGNORECASE)

    seen = set()
    out = []
    for t in targets:
        raw = str(t).strip()
        key = raw.lower()
        if not key or re.fullmatch(r"\d+", key):
            continue
        if key and key not in seen:
            seen.add(key)
            out.append(raw)
    return out


def _append_capped_nodes(
    lines: List[str],
    items: List[dict],
    id_prefix: str,
    emoji: str,
    label_key: str = "name",
    max_items: int = 40,
) -> List[str]:
    """
    Append nodes with a cap to keep Mermaid render stable on large environments.
    Returns created node IDs.
    """
    created = []
    for idx, item in enumerate(items[:max_items], start=1):
        name = str(item.get(label_key) or f"{id_prefix}-{idx}")
        nid = _safe_mermaid_id(f"{id_prefix}{idx}{name}", f"{id_prefix}{idx}")
        lines.append(f'    {nid}["{emoji} {name}"]')
        created.append(nid)
    remaining = len(items) - max_items
    if remaining > 0:
        more_id = _safe_mermaid_id(f"{id_prefix}More", f"{id_prefix}More")
        lines.append(f'    {more_id}["... +{remaining} more"]')
        created.append(more_id)
    return created


def _build_aws_network_diagram(summary: Dict) -> str:
    """
    Deterministic AWS topology in the same simplified style:
    User -> Internet -> Global DNS -> CloudEntry -> VPCs/Subnets/EC2.
    """
    networks = summary.get("networks", []) or []
    subnets = summary.get("subnets", []) or []
    compute = summary.get("compute", []) or []
    load_balancers = summary.get("load_balancers", []) or []
    internet_gateways = summary.get("internet_gateways", []) or []
    vpc_endpoints = summary.get("vpc_endpoints", []) or []
    route_tables = summary.get("route_tables", []) or []

    # Canonical VPC resolution map: support both display name and VPC id.
    vpc_alias_to_name: Dict[str, str] = {}
    for v in networks:
        name = str(v.get("name") or "").strip()
        vpc_id = str(v.get("vpc_id") or "").strip()
        if name:
            vpc_alias_to_name[name.lower()] = name
        if vpc_id:
            vpc_alias_to_name[vpc_id.lower()] = name or vpc_id

    # Ensure VPCs referenced by instances/subnets are present.
    def ensure_vpc_exists(vpc_raw: str):
        raw = str(vpc_raw or "").strip()
        if not raw:
            return
        if raw.lower() in vpc_alias_to_name:
            return
        networks.append({"name": raw, "vpc_id": raw, "cidr": ""})
        vpc_alias_to_name[raw.lower()] = raw

    for s in subnets:
        ensure_vpc_exists(s.get("vpc"))
    for vm in compute:
        ensure_vpc_exists(vm.get("vpc"))

    def resolve_vpc_name(vpc_raw: str) -> str:
        raw = str(vpc_raw or "").strip()
        if not raw:
            return raw
        return vpc_alias_to_name.get(raw.lower(), raw)

    lines: List[str] = []
    lines.append("graph LR")
    lines.append('User["👤 User"] --> Internet["🌐 Internet"]')
    lines.append('DNS["🌍 Global DNS"]')
    lines.append("Internet --> DNS")
    lines.append('subgraph CloudEnv["Cloud Environment"]')
    lines.append('  CloudEntry["☁ Workload Entry"]')

    vpc_to_subnets: Dict[str, List[dict]] = defaultdict(list)
    for s in subnets:
        vpc_name = resolve_vpc_name(s.get("vpc"))
        vpc_to_subnets[vpc_name.lower()].append(s)

    vpc_nodes: Dict[str, str] = {}
    vm_nodes: Dict[str, str] = {}

    for i, v in enumerate(networks, start=1):
        vpc_name = str(v.get("name") or f"vpc-{i}").strip()
        vpc_id = _safe_mermaid_id(vpc_name, f"VPC{i}")
        vpc_nodes[vpc_name] = vpc_id
        lines.append(f'  subgraph {vpc_id}["🧱 {vpc_name}"]')

        vpc_subnets = vpc_to_subnets.get(vpc_name.lower(), [])
        vms_in_vpc = [
            x for x in compute
            if resolve_vpc_name(x.get("vpc")).lower() == vpc_name.lower()
        ]

        subnet_names = {str(s.get("name") or "").strip() for s in vpc_subnets if str(s.get("name") or "").strip()}
        for vm in vms_in_vpc:
            sname = str(vm.get("subnet") or "").strip() or "default"
            subnet_names.add(sname)
        if not subnet_names:
            subnet_names.add("default")

        for j, sname in enumerate(sorted(subnet_names), start=1):
            sid = _safe_mermaid_id(f"{vpc_id}{sname}", f"SN{i}{j}")
            lines.append(f'    subgraph {sid}["Subnet: {sname}"]')
            subnet_vms = [
                x for x in vms_in_vpc
                if (str(x.get("subnet") or "").strip() or "default").lower() == sname.lower()
            ]
            # Group by name to avoid duplicate-node collisions for repeated names.
            name_counts: Dict[str, int] = defaultdict(int)
            for vm in subnet_vms:
                name_counts[str(vm.get("name") or "ec2-instance")] += 1

            for k, (vm_name, count) in enumerate(sorted(name_counts.items()), start=1):
                label = f"{vm_name} ({count}x)" if count > 1 else vm_name
                vm_id = _safe_mermaid_id(f"{vm_name}{k}", f"EC2{i}{j}{k}")
                vm_nodes[label] = vm_id
                lines.append(f'      {vm_id}["🖥️ {label}"]')
            lines.append("    end")

        lines.append("  end")

    # Instances without VPC mapping
    mapped_vm_names = set(vm_nodes.keys())
    remaining = [vm for vm in compute if str(vm.get("name") or "") not in mapped_vm_names]
    if remaining:
        lines.append('  subgraph UnassignedNet["Unassigned / Unknown Network"]')
        for i, vm in enumerate(remaining, start=1):
            vm_name = str(vm.get("name") or f"unmapped-ec2-{i}")
            vm_id = _safe_mermaid_id(vm_name, f"UEC2{i}")
            vm_nodes[vm_name] = vm_id
            lines.append(f'    {vm_id}["🖥️ {vm_name}"]')
        lines.append("  end")

    # Optional service groups (no noisy many-to-many arrows)
    if load_balancers or internet_gateways:
        lines.append('  subgraph EdgeIngress["Edge / Ingress"]')
        _append_capped_nodes(lines, load_balancers, "LB", "⚖️", max_items=60)
        _append_capped_nodes(lines, internet_gateways, "IGW", "🌉", max_items=60)
        lines.append("  end")

    if vpc_endpoints or route_tables:
        lines.append('  subgraph NetworkServices["Network Services"]')
        _append_capped_nodes(lines, vpc_endpoints, "VPCE", "🔒", max_items=100)
        _append_capped_nodes(lines, route_tables, "RT", "🧭", max_items=100)
        lines.append("  end")

    lines.append("end")

    # Single clean external flow
    lines.append("DNS --> CloudEntry")
    if vpc_nodes:
        lines.append(f"CloudEntry --> {next(iter(vpc_nodes.values()))}")

    return _clean_mermaid_output("\n".join(lines))


def _build_azure_network_diagram(summary: Dict) -> str:
    """
    Deterministic Azure topology:
    User -> Internet -> DNS -> VM (no Public IP nodes in topology)
    VNet -> Subnet -> VM hierarchy
    VNet peering shown where available.
    """
    networks = summary.get("networks", []) or []
    subnets = summary.get("subnets", []) or []
    compute = summary.get("compute", []) or []
    network_interfaces = summary.get("network_interfaces", []) or []
    load_balancers = summary.get("load_balancers", []) or []
    app_gateways = summary.get("application_gateways", []) or []
    frontdoors = summary.get("frontdoor_endpoints", []) or []
    traffic_mgr = summary.get("traffic_manager_profiles", []) or []
    private_endpoints = summary.get("private_endpoints", []) or []
    vnet_gateways = summary.get("virtual_network_gateways", []) or []
    local_gateways = summary.get("local_network_gateways", []) or []
    express_routes = summary.get("express_route_circuits", []) or []
    route_tables = summary.get("route_tables", []) or []

    # Major architecture services
    aks_clusters = summary.get("aks_clusters", []) or []
    vm_scale_sets = summary.get("vm_scale_sets", []) or []
    web_apps = summary.get("web_apps", []) or []
    static_web_apps = summary.get("static_web_apps", []) or []
    container_apps = summary.get("container_apps", []) or []
    storage_accounts = summary.get("storage_accounts", []) or summary.get("storage", []) or []
    database_services = summary.get("database_services", []) or summary.get("databases", []) or []
    # Force compact visualization for large enterprise assessments so everything is readable in one tab.
    compact_view = True

    def _safe_label_text(text: str) -> str:
        return str(text or "").replace('"', "'").replace("\n", " ")

    def _short_name(text: str, max_len: int = 18) -> str:
        s = str(text or "").strip()
        if len(s) <= max_len:
            return s
        return s[:max_len - 1] + "…"

    def _compact_group_node(
        lines_ref: List[str],
        id_seed: str,
        emoji: str,
        title: str,
        items: List[dict],
    ) -> str:
        if not items:
            return ""
        nid = _safe_mermaid_id(id_seed, id_seed)
        lines_ref.append(f'    {nid}["{emoji} {_safe_label_text(title)} ({len(items)})"]')
        return nid

    lines: List[str] = []
    lines.append("graph LR")
    lines.append('User["👤 User"] --> Internet["🌐 Internet"]')
    dns_label = "Azure DNS" if _has_azure_dns(summary) else "Global DNS"
    lines.append(f'DNS["🌍 {dns_label}"]')
    lines.append("Internet --> DNS")
    lines.append('subgraph CloudEnv["Cloud Environment"]')

    # Ensure VNets referenced only from VM rows are still rendered.
    known_vnets = {str(v.get("name") or "").strip().lower() for v in networks}
    for vm in compute:
        vname = str(vm.get("vnet") or "").strip()
        if vname and vname.lower() not in known_vnets:
            networks.append({"name": vname, "peerings": ""})
            known_vnets.add(vname.lower())
    # Ensure VNets referenced only from VM Scale Sets are also rendered
    for vmss in vm_scale_sets:
        vname = str(vmss.get("vnet") or "").strip()
        if vname and vname.lower() not in known_vnets:
            networks.append({"name": vname, "peerings": ""})
            known_vnets.add(vname.lower())

    # Make VMSS act like compute rows so they are included in subnet VM grouping
    existing_compute_names = {str(c.get("name") or "").strip() for c in compute}
    for vmss in vm_scale_sets:
        vmss_name = str(vmss.get("name") or "").strip()
        if vmss_name and vmss_name not in existing_compute_names:
            compute.append(vmss)

    # Index subnets by vnet.
    vnet_to_subnets: Dict[str, List[dict]] = defaultdict(list)
    for s in subnets:
        vnet_to_subnets[str(s.get("vnet") or "").strip().lower()].append(s)

    vm_nodes: Dict[str, str] = {}
    vnet_nodes: Dict[str, str] = {}
    vnet_nodes_by_lower: Dict[str, str] = {}

    # Single cloud entry point to avoid cluttered external arrows.
    lines.append('  CloudEntry["☁ Workload Entry"]')

    # Build VNet > Subnet > VM hierarchy.
    for i, v in enumerate(networks, start=1):
        vname = str(v.get("name") or f"VNet{i}").strip()
        vid = _safe_mermaid_id(vname, f"VNet{i}")
        vnet_nodes[vname] = vid
        vnet_nodes_by_lower[vname.lower()] = vid

        lines.append(f'  subgraph {vid}["🔷 {vname}"]')

        v_subnets = vnet_to_subnets.get(vname.lower(), [])
        vms_in_vnet = [x for x in compute if str(x.get("vnet") or "").strip().lower() == vname.lower()]
        shown_vm_in_vnet = 0
        max_vm_icons_per_vnet = 4

        # Union of subnet names from VNet sheet and VM rows to avoid dropping VMs.
        subnet_names = {str(s.get("name") or "").strip() for s in v_subnets if str(s.get("name") or "").strip()}
        for vm in vms_in_vnet:
            sname = str(vm.get("subnet") or "").strip() or "default"
            subnet_names.add(sname)
        if not subnet_names:
            subnet_names.add("default")

        for j, sname in enumerate(sorted(subnet_names), start=1):
            sid = _safe_mermaid_id(f"{vid}{sname}", f"SN{i}{j}")
            lines.append(f'    subgraph {sid}["Subnet: {sname}"]')
            subnet_vms = [
                x for x in vms_in_vnet
                if (str(x.get("subnet") or "").strip() or "default").lower() == sname.lower()
            ]
            if compact_view:
                remaining_slots = max_vm_icons_per_vnet - shown_vm_in_vnet
                if remaining_slots > 0:
                    for k, vm in enumerate(subnet_vms[:remaining_slots], start=1):
                        vm_name = str(vm.get("name") or f"vm{i}{j}{k}")
                        vm_id = _safe_mermaid_id(vm_name, f"VM{i}{j}{k}")
                        vm_nodes[vm_name] = vm_id
                        lines.append(f'      {vm_id}["🖥️ {_safe_label_text(_short_name(vm_name, 16))}"]')
                        shown_vm_in_vnet += 1
            else:
                for k, vm in enumerate(subnet_vms, start=1):
                    vm_name = str(vm.get("name") or f"vm{i}{j}{k}")
                    vm_id = _safe_mermaid_id(vm_name, f"VM{i}{j}{k}")
                    vm_nodes[vm_name] = vm_id
                    lines.append(f'      {vm_id}["🖥️ {_safe_label_text(_short_name(vm_name, 16))}"]')
            lines.append("    end")

        if compact_view:
            vm_total = len(vms_in_vnet)
            more = vm_total - shown_vm_in_vnet
            vm_summary = f"🖥️ Total VMs: {vm_total}"
            if more > 0:
                vm_summary += f" (showing {shown_vm_in_vnet}, +{more} more)"
            vm_summary_id = _safe_mermaid_id(f"{vid}VmTotal", f"VMSUM{i}")
            lines.append(f'    {vm_summary_id}["{_safe_label_text(vm_summary)}"]')

        lines.append("  end")

    # VMs not mapped to known VNets still shown in an Unassigned area.
    mapped_vm_names = set(vm_nodes.keys())
    remaining_vms = [vm for vm in compute if str(vm.get("name") or "") not in mapped_vm_names]
    if remaining_vms:
        lines.append('  subgraph UnassignedNet["Unassigned / Unknown Network"]')
        for i, vm in enumerate(remaining_vms, start=1):
            vm_name = str(vm.get("name") or f"unmappedvm{i}")
            vm_id = _safe_mermaid_id(vm_name, f"UVM{i}")
            vm_nodes[vm_name] = vm_id
            lines.append(f'    {vm_id}["🖥️ {vm_name}"]')
        lines.append("  end")

    # Network services & edge components
    edge_nodes: List[str] = []
    edge_node_hints: Dict[str, str] = {}
    lb_node_backend_targets: Dict[str, List[str]] = {}
    lb_node_resource_groups: Dict[str, str] = {}
    lb_nodes: List[str] = []
    needs_inferred_aks_lb = bool(aks_clusters) and not bool(load_balancers)
    if frontdoors or traffic_mgr or app_gateways or load_balancers or needs_inferred_aks_lb:
        lines.append('  subgraph EdgeIngress["Edge / Ingress"]')
        if compact_view:
            for prefix, emoji, title, items in [
                ("AFD", "🌍", "Front Door Endpoints", frontdoors),
                ("TM", "🧭", "Traffic Manager Profiles", traffic_mgr),
                ("AGW", "🚪", "Application Gateways", app_gateways),
            ]:
                nid = _compact_group_node(lines, f"{prefix}Group", emoji, title, items)
                if nid:
                    edge_nodes.append(nid)
            # Keep LBs detailed enough to preserve LB->AKS mapping.
            for idx, item in enumerate(load_balancers[:12], start=1):
                name = str(item.get("name") or f"LB-{idx}")
                nid = _safe_mermaid_id(f"LB{idx}{name}", f"LB{idx}")
                lines.append(f'    {nid}["⚖️ {name}"]')
                edge_nodes.append(nid)
                lb_nodes.append(nid)
                lb_node_resource_groups[nid] = str(item.get("resource_group") or "").strip().lower()
                lb_node_backend_targets[nid] = list(item.get("backend_targets") or [])
            if len(load_balancers) > 12:
                more_id = _safe_mermaid_id("LBMore", "LBMore")
                lines.append(f'    {more_id}["⚖️ Load Balancers (+{len(load_balancers) - 12} more)"]')
                edge_nodes.append(more_id)
                lb_nodes.append(more_id)
        else:
            for prefix, emoji, items, cap in [
                ("AFD", "🌍", frontdoors, 30),
                ("TM", "🧭", traffic_mgr, 30),
                ("AGW", "🚪", app_gateways, 50),
                ("LB", "⚖️", load_balancers, 60),
            ]:
                for idx, item in enumerate(items[:cap], start=1):
                    name = str(item.get("name") or f"{prefix}-{idx}")
                    nid = _safe_mermaid_id(f"{prefix}{idx}{name}", f"{prefix}{idx}")
                    lines.append(f'    {nid}["{emoji} {name}"]')
                    edge_nodes.append(nid)
                    hint = str(item.get("network_hint") or "").strip()
                    if hint:
                        edge_node_hints[nid] = hint
                    if prefix == "LB":
                        lb_nodes.append(nid)
                        lb_node_resource_groups[nid] = str(item.get("resource_group") or "").strip().lower()
                        lb_node_backend_targets[nid] = list(item.get("backend_targets") or [])
                remaining = len(items) - cap
                if remaining > 0:
                    more_id = _safe_mermaid_id(f"{prefix}More", f"{prefix}More")
                    lines.append(f'    {more_id}["... +{remaining} more"]')
                    edge_nodes.append(more_id)
                    if prefix == "LB":
                        lb_nodes.append(more_id)

        # If AKS exists but explicit load balancer rows are missing in assessment,
        # show inferred managed LB so architecture remains realistic.
        if needs_inferred_aks_lb:
            inferred_lb_id = _safe_mermaid_id("AKSManagedLoadBalancer", "AKSLB")
            lines.append(f'    {inferred_lb_id}["⚖️ AKS Managed Load Balancer (inferred)"]')
            edge_nodes.append(inferred_lb_id)
            lb_nodes.append(inferred_lb_id)
            lb_node_backend_targets[inferred_lb_id] = []
        lines.append("  end")

    network_service_nodes = []
    if vnet_gateways or local_gateways or express_routes or private_endpoints:
        lines.append('  subgraph NetworkServices["Network Services"]')
        if compact_view:
            for seed, emoji, title, items in [
                ("VNGGroup", "🔐", "Virtual Network Gateways", vnet_gateways),
                ("LNGGroup", "🏢", "Local Network Gateways", local_gateways),
                ("ERGroup", "🔌", "Express Route Circuits", express_routes),
                ("PEGroup", "🔒", "Private Endpoints", private_endpoints),
            ]:
                nid = _compact_group_node(lines, seed, emoji, title, items)
                if nid:
                    network_service_nodes.append(nid)
        else:
            network_service_nodes += _append_capped_nodes(lines, vnet_gateways, "VNG", "🔐", max_items=50)
            network_service_nodes += _append_capped_nodes(lines, local_gateways, "LNG", "🏢", max_items=50)
            network_service_nodes += _append_capped_nodes(lines, express_routes, "ER", "🔌", max_items=50)
            network_service_nodes += _append_capped_nodes(lines, private_endpoints, "PE", "🔒", max_items=80)
        lines.append("  end")

    if route_tables:
        lines.append('  subgraph PolicySecurity["Policy / Security"]')
        if compact_view:
            _compact_group_node(lines, "RTGroup", "🧭", "Route Tables", route_tables)
        else:
            _append_capped_nodes(lines, route_tables, "RT", "🧭", max_items=80)
        lines.append("  end")

    platform_nodes: List[str] = []
    aks_nodes: List[str] = []
    aks_node_by_name_norm: Dict[str, str] = {}
    aks_node_by_node_rg: Dict[str, str] = {}
    aks_node_rg_by_id: Dict[str, str] = {}
    aks_node_hints: Dict[str, str] = {}
    if (aks_clusters or vm_scale_sets or web_apps or static_web_apps or container_apps):
        lines.append('  subgraph PlatformTier["Platform / App Services"]')
        if compact_view:
            # Keep AKS detailed (few nodes, important topology anchor).
            for idx, item in enumerate((aks_clusters or [])[:20], start=1):
                name = str(item.get("name") or f"AKS-{idx}").strip()
                nid = _safe_mermaid_id(f"AKS{idx}{name}", f"AKS{idx}")
                lines.append(f'    {nid}["☸️ {name}"]')
                aks_nodes.append(nid)
                platform_nodes.append(nid)
                aks_node_by_name_norm[re.sub(r"[^a-z0-9]", "", name.lower())] = nid
                node_rg = str(item.get("node_resource_group") or "").strip().lower()
                if node_rg:
                    aks_node_by_node_rg[node_rg] = nid
                    aks_node_rg_by_id[nid] = node_rg
                hint = str(item.get("network_hint") or "").strip()
                if hint:
                    aks_node_hints[nid] = hint
            for seed, emoji, title, items in [
                ("VMSSGroup", "🧩", "VM Scale Sets", vm_scale_sets),
                ("WEBGroup", "🌐", "Web Apps", web_apps),
                ("SWAGroup", "📄", "Static Web Apps", static_web_apps),
                ("CAPPGroup", "📦", "Container Apps", container_apps),
            ]:
                nid = _compact_group_node(lines, seed, emoji, title, items)
                if nid:
                    platform_nodes.append(nid)
        # Fallback: ensure VMSS count node is visible even if compact grouping failed
        if vm_scale_sets and not any('VMSS' in p for p in platform_nodes):
            vmss_count = len(vm_scale_sets)
            nid = _safe_mermaid_id(f"VMSSGroupFallback{vmss_count}", f"VMSSF{vmss_count}")
            lines.append(f'    {nid}["🧩 VM Scale Sets ({vmss_count})"]')
            platform_nodes.append(nid)
        else:
            for idx, item in enumerate((aks_clusters or [])[:60], start=1):
                name = str(item.get("name") or f"AKS-{idx}").strip()
                nid = _safe_mermaid_id(f"AKS{idx}{name}", f"AKS{idx}")
                lines.append(f'    {nid}["☸️ {name}"]')
                aks_nodes.append(nid)
                platform_nodes.append(nid)
                aks_node_by_name_norm[re.sub(r"[^a-z0-9]", "", name.lower())] = nid
                node_rg = str(item.get("node_resource_group") or "").strip().lower()
                if node_rg:
                    aks_node_by_node_rg[node_rg] = nid
                    aks_node_rg_by_id[nid] = node_rg
                hint = str(item.get("network_hint") or "").strip()
                if hint:
                    aks_node_hints[nid] = hint
            if len(aks_clusters) > 60:
                more_id = _safe_mermaid_id("AKSMore", "AKSMore")
                lines.append(f'    {more_id}["... +{len(aks_clusters) - 60} more"]')
                aks_nodes.append(more_id)
                platform_nodes.append(more_id)
            platform_nodes += _append_capped_nodes(lines, vm_scale_sets, "VMSS", "🧩", max_items=60)
            platform_nodes += _append_capped_nodes(lines, web_apps, "WEB", "🌐", max_items=80)
            platform_nodes += _append_capped_nodes(lines, static_web_apps, "SWA", "📄", max_items=60)
            platform_nodes += _append_capped_nodes(lines, container_apps, "CAPP", "📦", max_items=80)
        lines.append("  end")

    data_nodes: List[str] = []
    if storage_accounts or database_services:
        lines.append('  subgraph DataTier["Data Services"]')
        if compact_view:
            st = _compact_group_node(lines, "STGroup", "🪣", "Storage Accounts", storage_accounts)
            db = _compact_group_node(lines, "DBGroup", "🗄️", "Database Services", database_services)
            if st:
                data_nodes.append(st)
            if db:
                data_nodes.append(db)
        else:
            data_nodes += _append_capped_nodes(lines, storage_accounts, "ST", "🪣", max_items=100)
            data_nodes += _append_capped_nodes(lines, database_services, "DB", "🗄️", max_items=100)
        lines.append("  end")

    lines.append("end")

    def _match_vnet_id_from_hint(hint: str) -> str:
        hl = str(hint or "").strip().lower()
        if not hl:
            return ""
        if hl in vnet_nodes_by_lower:
            return vnet_nodes_by_lower[hl]
        for vname_l, vid in vnet_nodes_by_lower.items():
            if hl in vname_l or vname_l in hl:
                return vid
        return ""

    # Build RG -> VNet hints from observed NICs and VMs.
    rg_to_vnet: Dict[str, str] = {}
    for vm in compute:
        rg = str(vm.get("resource_group") or "").strip().lower()
        vn = str(vm.get("vnet") or "").strip().lower()
        if rg and vn and vn in vnet_nodes_by_lower and rg not in rg_to_vnet:
            rg_to_vnet[rg] = vnet_nodes_by_lower[vn]
    for ni in network_interfaces:
        rg = str(ni.get("resource_group") or "").strip().lower()
        vnet_hint = str(ni.get("vnet") or "").strip()
        vid = _match_vnet_id_from_hint(vnet_hint)
        if rg and vid and rg not in rg_to_vnet:
            rg_to_vnet[rg] = vid

    # Single external flow path.
    lines.append("DNS --> CloudEntry")

    # Optional minimal anchoring: one line from entry to first VNet to keep left-to-right flow.
    representative_vnet = ""
    if vnet_nodes:
        representative_vnet = next(iter(vnet_nodes.values()))
        lines.append(f"CloudEntry --> {representative_vnet}")
        # Attach edge nodes to their inferred VNet when possible.
        for n in edge_nodes:
            is_lb = n in lb_nodes
            target = _match_vnet_id_from_hint(edge_node_hints.get(n, ""))
            if not target and is_lb:
                rg = lb_node_resource_groups.get(n, "")
                target = rg_to_vnet.get(rg, "")
            if target:
                lines.append(f"{n} --> {target}")
            elif (not is_lb) and representative_vnet:
                # Non-LB edge resources can safely anchor to representative VNet for layout.
                lines.append(f"{n} --> {representative_vnet}")
    # Keep one clean ingress line; avoid noisy direct fan-out from CloudEntry.

    # Place AKS against its actual VNet where hinting is available.
    aks_vnet_targets: Dict[str, str] = {}
    for aks_id in aks_nodes:
        target = _match_vnet_id_from_hint(aks_node_hints.get(aks_id, ""))
        if not target:
            target = rg_to_vnet.get(aks_node_rg_by_id.get(aks_id, ""), "")
        if target:
            aks_vnet_targets[aks_id] = target

    # Heuristic: if most AKS clusters map to one VNet and a minority are missing, use that.
    if aks_nodes and len(aks_vnet_targets) < len(aks_nodes):
        mapped_values = list(aks_vnet_targets.values())
        if mapped_values:
            dominant = max(set(mapped_values), key=mapped_values.count)
            for aks_id in aks_nodes:
                if aks_id not in aks_vnet_targets:
                    aks_vnet_targets[aks_id] = dominant

    for aks_id, target in aks_vnet_targets.items():
        lines.append(f"{target} --> {aks_id}")

    # AKS commonly fronts through managed LB; map by Node RG first.
    for lb_id in lb_nodes:
        rg = lb_node_resource_groups.get(lb_id, "")
        target_aks = aks_node_by_node_rg.get(rg)
        if target_aks:
            lines.append(f"{lb_id} --> {target_aks}")
        elif aks_nodes and len(aks_nodes) == 1:
            lines.append(f"{lb_id} --> {aks_nodes[0]}")

    # Explicit LB -> VM links where backend targets are discoverable.
    def _norm_name(x: str) -> str:
        return re.sub(r"[^a-z0-9]", "", str(x or "").lower())

    vm_lookup = {_norm_name(vm_name): vm_id for vm_name, vm_id in vm_nodes.items()}
    if not compact_view:
        for lb_id, targets in lb_node_backend_targets.items():
            for t in targets:
                nt = _norm_name(t)
                if not nt:
                    continue
                # Exact normalized match first
                vm_id = vm_lookup.get(nt)
                if vm_id:
                    lines.append(f"{lb_id} --> {vm_id}")
                    continue
                # Fallback contains matching
                for vmn, vid in vm_lookup.items():
                    if nt in vmn or vmn in nt:
                        lines.append(f"{lb_id} --> {vid}")
                        break

    # VNet peering links.
    used_peering_edges = set()
    for v in networks:
        src_name = str(v.get("name") or "").strip()
        src_id = vnet_nodes.get(src_name)
        if not src_id:
            continue
        targets = _parse_peering_targets(v.get("peerings"))
        for t in targets:
            # Match by exact name first, then contains.
            tgt_name = None
            if t in vnet_nodes:
                tgt_name = t
            else:
                tl = t.lower()
                for candidate in vnet_nodes:
                    if tl == candidate.lower() or tl in candidate.lower() or candidate.lower() in tl:
                        tgt_name = candidate
                        break
            if not tgt_name or tgt_name == src_name:
                continue
            a = src_id
            b = vnet_nodes[tgt_name]
            edge_key = tuple(sorted([a, b]))
            if edge_key in used_peering_edges:
                continue
            used_peering_edges.add(edge_key)
            lines.append(f"{a} ---|Peering| {b}")

    return _clean_mermaid_output("\n".join(lines))


# ---------------------------------------------------------------------------
# AWS draw.io renderer
# ---------------------------------------------------------------------------
# Uses drawio's built-in AWS4 stencil library (no external files needed).
# AWS resource icons are referenced via `shape=mxgraph.aws4.resourceIcon;
# resIcon=mxgraph.aws4.<resource>`. Container boxes (VPCs, Subnets) use
# `shape=mxgraph.aws4.group` so they render with AWS's official dashed
# borders and labels.

# AWS service category colors (matches official AWS Architecture Center).
_AWS_COLOR_COMPUTE = "#ED7100"   # orange
_AWS_COLOR_NETWORK = "#8C4FFF"   # purple
_AWS_COLOR_STORAGE = "#7AA116"   # green
_AWS_COLOR_DB = "#C925D1"        # magenta
_AWS_COLOR_SECURITY = "#DD344C"  # red
_AWS_COLOR_INTEGRATION = "#E7157B"


def _aws_icon_style(res_icon: str, color: str) -> str:
    return (
        f"sketch=0;outlineConnect=0;fontColor=#232F3E;gradientColor=none;"
        f"fillColor={color};strokeColor=#232F3E;dashed=0;verticalLabelPosition=bottom;"
        f"verticalAlign=top;align=center;html=1;fontSize=10;fontStyle=0;aspect=fixed;"
        f"shape=mxgraph.aws4.resourceIcon;resIcon=mxgraph.aws4.{res_icon};"
    )


def _build_aws_drawio_network_xml(summary: Dict) -> str:
    """AWS network topology in drawio XML — VPC → Subnet → EC2 layout."""
    networks = summary.get("networks", []) or []
    subnets = summary.get("subnets", []) or []
    compute = summary.get("compute", []) or []
    storage = summary.get("storage", []) or []
    databases = summary.get("databases", []) or []
    internet_gateways = summary.get("internet_gateways", []) or []
    vpc_endpoints = summary.get("vpc_endpoints", []) or []
    load_balancers = summary.get("load_balancers", []) or []
    nat_gateways = summary.get("nat_gateways", []) or []
    auto_scaling_groups = summary.get("auto_scaling_groups", []) or []
    elastic_ips = summary.get("elastic_ips", []) or []
    volumes = summary.get("volumes", []) or []
    ecs_clusters = summary.get("ecs_clusters", []) or []
    ecs_services = summary.get("ecs_services", []) or []
    eks_clusters = summary.get("eks_clusters", []) or []
    ecr_repos = summary.get("ecr_repositories", []) or []
    lambda_functions = summary.get("lambda_functions", []) or []
    rds_clusters = summary.get("rds_clusters", []) or []
    dynamodb_tables = summary.get("dynamodb_tables", []) or []

    # --- mxfile skeleton -----------------------------------------------------
    mxfile = ET.Element("mxfile", {"host": "app.diagrams.net", "type": "device"})
    diagram = ET.SubElement(mxfile, "diagram", {"id": "aws-network", "name": "AWS Network"})
    model = ET.SubElement(diagram, "mxGraphModel", {
        "dx": "1600", "dy": "900", "grid": "1", "gridSize": "10",
        "guides": "1", "tooltips": "1", "connect": "1", "arrows": "1",
        "fold": "1", "page": "1", "pageScale": "1",
        "pageWidth": "2400", "pageHeight": "1600", "math": "0", "shadow": "0",
    })
    root = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", {"id": "0"})
    ET.SubElement(root, "mxCell", {"id": "1", "parent": "0"})

    cell_seq = [1]  # mutable counter

    def add_vertex(value, x, y, w, h, style, parent="1"):
        cell_seq[0] += 1
        nid = f"n{cell_seq[0]}"
        c = ET.SubElement(root, "mxCell", {
            "id": nid, "value": str(value or ""),
            "style": style, "vertex": "1", "parent": parent,
        })
        ET.SubElement(c, "mxGeometry", {
            "x": str(int(x)), "y": str(int(y)),
            "width": str(int(w)), "height": str(int(h)), "as": "geometry",
        })
        return nid

    def add_edge(src, tgt, style="endArrow=block;strokeColor=#5d6d7e;"):
        cell_seq[0] += 1
        nid = f"e{cell_seq[0]}"
        c = ET.SubElement(root, "mxCell", {
            "id": nid, "style": style, "edge": "1",
            "parent": "1", "source": src, "target": tgt,
        })
        ET.SubElement(c, "mxGeometry", {"relative": "1", "as": "geometry"})
        return nid

    # --- Styles --------------------------------------------------------------
    style_vpc_container = (
        "rounded=0;whiteSpace=wrap;html=1;dashed=1;dashPattern=1 4;"
        "strokeColor=#248814;fillColor=#E9F3E6;strokeWidth=2;verticalAlign=top;"
        "fontSize=14;fontStyle=1;fontColor=#248814;align=left;spacingLeft=12;"
    )
    style_subnet_public = (
        "rounded=0;whiteSpace=wrap;html=1;dashed=0;"
        "strokeColor=#7AA116;fillColor=#F2F8E8;strokeWidth=1;verticalAlign=top;"
        "fontSize=12;fontStyle=1;fontColor=#3F6F00;align=left;spacingLeft=8;"
    )
    style_subnet_private = (
        "rounded=0;whiteSpace=wrap;html=1;dashed=0;"
        "strokeColor=#147EBA;fillColor=#E6F2F8;strokeWidth=1;verticalAlign=top;"
        "fontSize=12;fontStyle=1;fontColor=#0E5A86;align=left;spacingLeft=8;"
    )
    style_outer = (
        "rounded=0;whiteSpace=wrap;html=1;strokeColor=#7f8c8d;"
        "fillColor=#f8f9fa;strokeWidth=1;verticalAlign=top;fontSize=12;fontStyle=1;"
        "align=left;spacingLeft=8;"
    )
    style_panel = (
        "rounded=0;whiteSpace=wrap;html=1;strokeColor=#95a5a6;"
        "fillColor=#ffffff;strokeWidth=1;verticalAlign=top;fontSize=12;fontStyle=1;"
        "align=left;spacingLeft=8;"
    )

    # --- Index data ----------------------------------------------------------
    # VPC id ↔ name resolution.
    vpc_id_to_name: Dict[str, str] = {}
    for v in networks:
        name = str(v.get("name") or "").strip()
        vid = str(v.get("vpc_id") or "").strip()
        if vid:
            vpc_id_to_name[vid.lower()] = name or vid
        if name:
            vpc_id_to_name[name.lower()] = name

    def _resolve_vpc(raw):
        s = str(raw or "").strip().lower()
        return vpc_id_to_name.get(s, raw or "")

    # Subnet id → record (for label/cidr lookup when EC2 references by ID).
    subnet_id_to_record: Dict[str, dict] = {}
    for sn in subnets:
        sid = str(sn.get("subnet_id") or "").strip().lower()
        if sid:
            subnet_id_to_record[sid] = sn

    # Group compute by (VPC, subnet-key).
    # subnet-key is the subnet ID when available, else the subnet name field.
    vpc_subnet_to_compute: Dict[tuple, List[dict]] = defaultdict(list)
    for c in compute:
        vpc = _resolve_vpc(c.get("vpc"))
        snet = str(c.get("subnet") or "").strip()
        if not vpc:
            vpc = "Unassigned"
        if not snet:
            snet = "default"
        vpc_subnet_to_compute[(vpc, snet)].append(c)

    # All subnet keys per VPC.
    vpc_to_subnet_keys: Dict[str, List[str]] = defaultdict(list)
    for (vpc, snet) in vpc_subnet_to_compute.keys():
        if snet not in vpc_to_subnet_keys[vpc]:
            vpc_to_subnet_keys[vpc].append(snet)
    # Also include subnets that have no compute (so they still render).
    for sn in subnets:
        vpc = _resolve_vpc(sn.get("vpc"))
        sid = str(sn.get("subnet_id") or "").strip()
        sname = str(sn.get("name") or "").strip()
        key = sid or sname or "default"
        if vpc and key not in vpc_to_subnet_keys[vpc]:
            vpc_to_subnet_keys[vpc].append(key)

    # Ordered list of VPCs to render (deterministic).
    vpc_names_to_render = [str(v.get("name") or "").strip() for v in networks if v.get("name")]
    # Add any VPCs found in compute but not in networks list.
    for (vpc, _snet) in vpc_subnet_to_compute.keys():
        if vpc and vpc not in vpc_names_to_render and vpc != "Unassigned":
            vpc_names_to_render.append(vpc)
    if "Unassigned" in (v for (v, _) in vpc_subnet_to_compute.keys()):
        vpc_names_to_render.append("Unassigned")

    # --- Layout constants ----------------------------------------------------
    env_x, env_y = 60, 60
    ingress_y = env_y + 30
    vpc_w_each = 760
    vpc_gap_x = 40
    vpc_top_y = env_y + 200

    # Per-subnet capacity / cell sizing.
    vm_max_per_subnet = 6  # AWS often has more EC2 per subnet; allow more than Azure (4)
    cell_w = 130
    cell_h = 76
    cell_gap = 16
    grid_cols = 3  # 3-column grid → fits 6 cells in 2 rows nicely

    # --- Ingress (User → Internet → Route 53) --------------------------------
    user_id = add_vertex("User", env_x + 20, ingress_y, 48, 48,
                         _aws_icon_style("users", _AWS_COLOR_SECURITY))
    inet_id = add_vertex("Internet", env_x + 110, ingress_y, 48, 48,
                         _aws_icon_style("internet_gateway", _AWS_COLOR_NETWORK))
    r53_id = add_vertex("Route 53", env_x + 210, ingress_y, 48, 48,
                        _aws_icon_style("route_53", _AWS_COLOR_NETWORK))
    add_edge(user_id, inet_id)
    add_edge(inet_id, r53_id)

    # Map EC2 instance-id → its drawio node id so the LB pass can draw
    # connectivity arrows. EC2 instance IDs live in the `connectivity` text
    # of each compute record (e.g., "Attached to: i-011a1a99958420bae | ...").
    instance_id_to_node: Dict[str, str] = {}
    _iid_re = re.compile(r"\bi-[0-9a-f]{8,}\b", re.IGNORECASE)

    # --- Render each VPC -----------------------------------------------------
    rendered_vpc_ids: List[str] = []  # for ingress arrow to first VPC
    vpc_box_max_right = env_x
    vpc_box_max_bottom = vpc_top_y

    for vi, vpc_name in enumerate(vpc_names_to_render):
        vpc_x = env_x + 20 + vi * (vpc_w_each + vpc_gap_x)

        subnet_keys = vpc_to_subnet_keys.get(vpc_name, [])
        if not subnet_keys:
            continue

        # Compute VPC height: each subnet box height = header + cells.
        per_subnet_heights = []
        for sk in subnet_keys:
            cnt = len(vpc_subnet_to_compute.get((vpc_name, sk), []))
            shown = min(cnt, vm_max_per_subnet)
            rows = max(1, (shown + grid_cols - 1) // grid_cols)
            h = 60 + rows * (cell_h + cell_gap)
            per_subnet_heights.append(h)

        vpc_h = 60 + sum(per_subnet_heights) + 30 * len(subnet_keys)
        vpc_id = add_vertex(f"VPC: {vpc_name}", vpc_x, vpc_top_y, vpc_w_each, vpc_h,
                            style_vpc_container)
        rendered_vpc_ids.append(vpc_id)
        vpc_box_max_right = max(vpc_box_max_right, vpc_x + vpc_w_each)
        vpc_box_max_bottom = max(vpc_box_max_bottom, vpc_top_y + vpc_h)

        # Subnets stacked vertically inside the VPC.
        subnet_y_cursor = vpc_top_y + 40
        for si, sk in enumerate(subnet_keys):
            # Resolve subnet display: prefer name + CIDR if we have it.
            sn_rec = subnet_id_to_record.get(sk.lower())
            if sn_rec:
                sn_label = sn_rec.get("name") or sk
                cidr = sn_rec.get("cidr") or ""
                az = sn_rec.get("az") or ""
            else:
                sn_label = sk
                cidr = ""
                az = ""
            # Naive public/private guess from name.
            is_public = "public" in str(sn_label).lower() or "pub" in str(sn_label).lower()
            sn_style = style_subnet_public if is_public else style_subnet_private
            sn_h = per_subnet_heights[si]

            badge = f"{sn_label}"
            if cidr:
                badge += f"  ·  {cidr}"
            if az:
                badge += f"  ·  {az}"

            subnet_x = vpc_x + 20
            subnet_w = vpc_w_each - 40
            sn_box_id = add_vertex(badge, subnet_x, subnet_y_cursor, subnet_w, sn_h, sn_style)

            # Render EC2 grid inside subnet.
            ec2_items = vpc_subnet_to_compute.get((vpc_name, sk), [])
            # Prioritize ASG-tagged items (mirrors VMSS handling on Azure side).
            asg_items = [c for c in ec2_items if c.get("is_asg") or c.get("is_vmss")]
            other_items = [c for c in ec2_items if c not in asg_items]
            remaining = max(0, vm_max_per_subnet - len(asg_items))
            show_items = asg_items + other_items[:remaining]
            overflow = len(ec2_items) - len(show_items)

            grid_x0 = subnet_x + 20
            grid_y0 = subnet_y_cursor + 36
            for idx, ec2 in enumerate(show_items):
                row = idx // grid_cols
                col = idx % grid_cols
                cx = grid_x0 + col * (cell_w + cell_gap)
                cy = grid_y0 + row * (cell_h + cell_gap)
                ec2_name = str(ec2.get("name") or f"ec2-{idx+1}")
                is_asg = ec2.get("is_asg") or ec2.get("is_vmss")
                icon = "auto_scaling2" if is_asg else "ec2"
                ec2_node_id = add_vertex(ec2_name[:24], cx, cy, 50, 56,
                                         _aws_icon_style(icon, _AWS_COLOR_COMPUTE))
                # Index by every instance-id found in the connectivity blob
                # so the LB pass can target this node by ID.
                conn_text = str(ec2.get("connectivity") or "")
                for iid in _iid_re.findall(conn_text):
                    instance_id_to_node[iid.lower()] = ec2_node_id

            if overflow > 0:
                # Small badge bottom-right showing how many were hidden.
                add_vertex(f"+{overflow} more", subnet_x + subnet_w - 100,
                           subnet_y_cursor + sn_h - 26, 90, 18,
                           "rounded=1;whiteSpace=wrap;html=1;strokeColor=#95a5a6;"
                           "fillColor=#ecf0f1;fontSize=10;align=center;")

            subnet_y_cursor += sn_h + 16

    # --- Side panels ---------------------------------------------------------
    panel_x = vpc_box_max_right + 40
    panel_y = vpc_top_y

    def _render_panel(title, items, icon_name, color, max_items=20):
        nonlocal panel_y
        if not items:
            return
        panel_w = 240
        rows = min(len(items), max_items)
        cell_per_row = 2
        n_rows = (rows + cell_per_row - 1) // cell_per_row
        panel_h = 60 + n_rows * 74
        pid = add_vertex(f"{title} ({len(items)})", panel_x, panel_y, panel_w, panel_h,
                         style_panel)
        for idx, it in enumerate(items[:max_items]):
            r = idx // cell_per_row
            c = idx % cell_per_row
            x = panel_x + 30 + c * 100
            y = panel_y + 40 + r * 74
            label = str(it.get("name") or "")[:18]
            add_vertex(label, x, y, 50, 56, _aws_icon_style(icon_name, color))
        panel_y += panel_h + 30

    # ---- Load Balancers: place inside their VPC + draw arrows to targets ----
    # LBs are part of the network topology; we render them inside the VPC box
    # (top edge) and connect them to EC2 targets when instance IDs are listed
    # in any of the LB row's columns or fallback raw_row.
    def _lb_target_instance_ids(lb: dict) -> List[str]:
        ids: List[str] = []
        ids.extend(_iid_re.findall(str(lb.get("targets") or "")))
        raw = lb.get("raw_row") or {}
        if isinstance(raw, dict):
            joined = " ".join(str(x) for x in raw.values() if x is not None)
            ids.extend(_iid_re.findall(joined))
        seen, out = set(), []
        for i in ids:
            il = i.lower()
            if il not in seen:
                seen.add(il)
                out.append(il)
        return out

    # Track LB y-offset per VPC so multiple LBs stack neatly.
    vpc_to_lb_count: Dict[str, int] = defaultdict(int)
    # Get rendered VPC box positions: scan for vertex cells whose value starts
    # with "VPC: ". The geometry x/y/width gives us the placement.
    vpc_box_geom: Dict[str, dict] = {}
    for cell in root.findall("mxCell"):
        v = cell.get("value") or ""
        if v.startswith("VPC: "):
            vname = v[len("VPC: "):].strip()
            geom = cell.find("mxGeometry")
            if geom is not None:
                vpc_box_geom[vname] = {
                    "id": cell.get("id"),
                    "x": int(geom.get("x", "0")),
                    "y": int(geom.get("y", "0")),
                    "w": int(geom.get("width", "0")),
                    "h": int(geom.get("height", "0")),
                }

    lb_unplaced: List[dict] = []
    for lb in load_balancers:
        lb_vpc = _resolve_vpc(lb.get("vpc"))
        if not lb_vpc:
            lb_unplaced.append(lb)
            continue
        geom = vpc_box_geom.get(lb_vpc)
        if not geom:
            lb_unplaced.append(lb)
            continue
        # Place LB icon just inside the VPC box, top-right corner, stacking
        # horizontally then wrapping if there are many.
        slot = vpc_to_lb_count[lb_vpc]
        vpc_to_lb_count[lb_vpc] += 1
        lb_x = geom["x"] + geom["w"] - 70 - (slot * 70)
        lb_y = geom["y"] + 8
        lb_label = str(lb.get("name") or "lb")[:20]
        scheme = str(lb.get("scheme") or "").strip()
        if scheme:
            lb_label = f"{lb_label} ({scheme})"
        lb_node = add_vertex(lb_label, lb_x, lb_y, 50, 56,
                             _aws_icon_style("elastic_load_balancing", _AWS_COLOR_NETWORK))
        # Draw arrows to target instances when we can resolve them.
        for tgt_id in _lb_target_instance_ids(lb):
            tgt_node = instance_id_to_node.get(tgt_id)
            if tgt_node:
                add_edge(lb_node, tgt_node,
                         style="endArrow=block;endFill=1;strokeColor=#8C4FFF;"
                               "strokeWidth=1;dashed=0;")

    # ---- Side panels for topology-relevant resources only ----
    # (EIPs, EBS Volumes, ECR Repos removed — they're inventory, not topology.)

    # Compute / scaling
    _render_panel("Auto Scaling Groups", auto_scaling_groups, "auto_scaling2", _AWS_COLOR_COMPUTE)
    _render_panel("Lambda Functions", lambda_functions, "lambda", _AWS_COLOR_COMPUTE)
    _render_panel("ECS Clusters", ecs_clusters, "elastic_container_service", _AWS_COLOR_COMPUTE)
    _render_panel("ECS Services", ecs_services, "elastic_container_service", _AWS_COLOR_COMPUTE, max_items=12)
    _render_panel("EKS Clusters", eks_clusters, "elastic_kubernetes_service", _AWS_COLOR_COMPUTE)

    # Network — only LBs without a resolvable VPC fall back to a side panel
    if lb_unplaced:
        _render_panel("Load Balancers (no VPC)", lb_unplaced, "elastic_load_balancing", _AWS_COLOR_NETWORK)
    _render_panel("Internet Gateways", internet_gateways, "internet_gateway", _AWS_COLOR_NETWORK)
    _render_panel("NAT Gateways", nat_gateways, "nat_gateway", _AWS_COLOR_NETWORK)
    _render_panel("VPC Endpoints", vpc_endpoints, "endpoints", _AWS_COLOR_NETWORK)

    # Storage — S3 is the only one with topology meaning (VPC endpoints connect
    # to it). EBS volumes are per-EC2 disks; not a network entity.
    _render_panel("S3 Buckets", storage, "simple_storage_service", _AWS_COLOR_STORAGE)

    # Database
    _render_panel("RDS Instances", databases, "rds", _AWS_COLOR_DB)
    _render_panel("RDS Clusters", rds_clusters, "rds", _AWS_COLOR_DB)
    _render_panel("DynamoDB Tables", dynamodb_tables, "dynamodb", _AWS_COLOR_DB)

    # Connect ingress to the first VPC.
    if rendered_vpc_ids:
        add_edge(r53_id, rendered_vpc_ids[0],
                 style="endArrow=block;endFill=1;strokeColor=#5d6d7e;strokeWidth=1;")

    return ET.tostring(mxfile, encoding="unicode")


def _xml_escape(value: str) -> str:
    """Escape a string for use inside an XML attribute value."""
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _clean_cidr(value: str) -> str:
    """Turn an address space like "['10.0.0.0/24']" into "10.0.0.0/24"."""
    s = str(value or "").strip()
    if not s:
        return ""
    s = s.strip("[]")
    parts = [p.strip().strip("'\"") for p in s.split(",") if p.strip().strip("'\"")]
    return ", ".join(parts)


def _shelf_pack(sizes, max_w, gap):
    """Pack (w, h) boxes left-to-right, wrapping when the row exceeds max_w.

    Returns (positions, content_w, content_h) where positions[i] = (x, y) of
    box i relative to the content origin (0, 0). Guarantees boxes never
    overlap because each is advanced past the previous one plus `gap`.
    """
    positions = []
    x = y = row_h = content_w = 0
    for (w, h) in sizes:
        if x > 0 and (x + w) > max_w:
            x = 0
            y += row_h + gap
            row_h = 0
        positions.append((x, y))
        x += w + gap
        row_h = max(row_h, h)
        content_w = max(content_w, x - gap)
    content_h = y + row_h
    return positions, content_w, content_h


def _build_azure_drawio_nested(summary: Dict, arrows: bool = False) -> str:
    """Deterministic nested draw.io layout for Azure.

    Hierarchy: Region → VNet → Subnet → resource icons. Every container is
    auto-sized to exactly fit its children (packed with fixed gaps), so boxes
    never overlap regardless of how many subnets/resources a VNet has. This
    mirrors the curated IIFL-HFC reference layout.
    """
    # ---- layout constants (tuned to the IIFL-HFC reference template) ----
    ICON_W, ICON_H = 36, 36            # native azure2 icon size (matches template)
    CELL_W, CELL_H = 84, 58            # packing footprint per icon (label room)
    G = 8                              # gap between siblings (→ ~92px column pitch)
    SUB_PAD, SUB_HEADER = 8, 28        # subnet inner padding / title band
    SUB_COLS = 6                       # max icons per row inside a subnet
    MIN_SUBNET_W = 108
    VNET_PAD, VNET_HEADER = 10, 30
    VNET_MAXW = 1600                   # wrap subnets past this inner width
    MIN_VNET_W = 150
    REGION_PAD, REGION_HEADER = 12, 36
    REGION_MAXW = 4200                 # wrap VNets past this inner width
    MARGIN, REGION_GAP = 40, 44
    MAX_ICONS_PER_BOX = 1000           # effectively uncapped — show every resource

    networks = summary.get("networks", []) or []
    subnets = summary.get("subnets", []) or []
    compute = summary.get("compute", []) or []
    vm_scale_sets = summary.get("vm_scale_sets", []) or []
    aks_clusters = summary.get("aks_clusters", []) or []
    load_balancers = summary.get("load_balancers", []) or []
    application_gateways = summary.get("application_gateways", []) or []
    private_endpoints = summary.get("private_endpoints", []) or []
    databases = summary.get("database_services", []) or summary.get("databases", []) or []
    storage_accounts = summary.get("storage_accounts", []) or summary.get("storage", []) or []
    web_apps = summary.get("web_apps", []) or []
    network_interfaces = summary.get("network_interfaces", []) or []

    # Native draw.io Azure (azure2) stencil library paths — same as the template.
    # These render as crisp built-in shapes (no embedded SVG) inside draw.io.
    AZ2 = "img/lib/azure2"
    icon = {
        "vm": f"{AZ2}/compute/Virtual_Machine.svg",
        "vmss": f"{AZ2}/compute/VM_Scale_Sets.svg",
        "aks": f"{AZ2}/compute/Kubernetes_Services.svg",
        "lb": f"{AZ2}/networking/Load_Balancers.svg",
        "appgw": f"{AZ2}/networking/Application_Gateways.svg",
        "pe": f"{AZ2}/networking/Private_Link.svg",
        "sql": f"{AZ2}/databases/SQL_Database.svg",
        "storage": f"{AZ2}/storage/Storage_Accounts.svg",
        "web": f"{AZ2}/compute/App_Services.svg",
        "vnet": f"{AZ2}/networking/Virtual_Networks.svg",
    }

    # ---- container styles (colors taken from the reference template) ----
    _CBASE = ("rounded=1;arcSize=2;whiteSpace=wrap;html=1;verticalAlign=top;align=left;"
              "spacingLeft=8;spacingTop=4;container=1;collapsible=0;")
    S_REGION = _CBASE + "fillColor=#F7F7F7;strokeColor=#8A8A8A;fontStyle=1;fontSize=14;"
    S_VNET = _CBASE + "fillColor=#EAF3FC;strokeColor=#1F6FC5;fontStyle=1;fontSize=12;"
    S_SUBNET = _CBASE + "fillColor=#FFFFFF;strokeColor=#7FA9D9;fontSize=11;"
    S_INV_REGION = _CBASE + "fillColor=#FFFFFF;strokeColor=#9E9E9E;fontStyle=1;fontSize=14;"
    S_AKS_REGION = _CBASE + "fillColor=#FFFCF0;strokeColor=#D9A400;fontStyle=1;fontSize=14;"
    S_AKS_BOX = ("rounded=1;arcSize=6;whiteSpace=wrap;html=1;container=1;collapsible=1;collapsed=1;"
                 "verticalAlign=top;align=left;spacingLeft=8;fontSize=11;fontStyle=1;fillColor=#FFF7E6;strokeColor=#D9A400;")
    # Native azure2 image shape (matches the template's icon styling exactly).
    S_ICON = ("image;aspect=fixed;html=1;points=[];align=center;fontSize=9;"
              "verticalLabelPosition=bottom;verticalAlign=top;labelBackgroundColor=none;whiteSpace=wrap;")
    # Edge styles for connections.
    E_PEER = "edgeStyle=none;rounded=1;html=1;endArrow=none;dashed=1;strokeColor=#C55A11;strokeWidth=1.5;"
    E_CONN = "edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;endArrow=block;endFill=1;strokeColor=#16A085;strokeWidth=1.2;"

    cells = []
    _cid = {"n": 0}

    def new_id(prefix="c"):
        _cid["n"] += 1
        return f"{prefix}{_cid['n']}"

    def emit(cid, value, style, x, y, w, h, parent):
        cells.append(
            f'<mxCell id="{cid}" value="{_xml_escape(value)}" style="{_xml_escape(style)}" '
            f'vertex="1" parent="{parent}">'
            f'<mxGeometry x="{int(x)}" y="{int(y)}" width="{int(w)}" height="{int(h)}" as="geometry"/>'
            f'</mxCell>'
        )

    def emit_collapsed(cid, value, style, x, y, cw, ch, ew, eh, parent):
        # A draw.io collapsed container: on-screen size is (cw, ch) (just the
        # header); alternateBounds holds the expanded size (ew, eh) used when
        # the user double-clicks to expand. Children are hidden until expanded.
        cells.append(
            f'<mxCell id="{cid}" value="{_xml_escape(value)}" style="{_xml_escape(style)}" '
            f'vertex="1" parent="{parent}">'
            f'<mxGeometry x="{int(x)}" y="{int(y)}" width="{int(cw)}" height="{int(ch)}" as="geometry">'
            f'<mxRectangle width="{int(ew)}" height="{int(eh)}" as="alternateBounds"/>'
            f'</mxGeometry></mxCell>'
        )

    icon_id_by_name = {}   # resource name (lower) → emitted icon cell id (for arrows)

    def box_size_for_icons(n):
        """Inner content size (w, h) for n icons packed SUB_COLS-wide."""
        if n <= 0:
            return CELL_W, 0
        sizes = [(CELL_W, CELL_H)] * n
        max_row_w = SUB_COLS * (CELL_W + G)
        _, cw, ch = _shelf_pack(sizes, max_row_w, G)
        return cw, ch

    def emit_icons(items, parent_id):
        shown = items[:MAX_ICONS_PER_BOX]
        sizes = [(CELL_W, CELL_H)] * len(shown)
        max_row_w = SUB_COLS * (CELL_W + G)
        positions, _, _ = _shelf_pack(sizes, max_row_w, G)
        off = (CELL_W - ICON_W) // 2   # center the 36px icon in its footprint cell
        for (label, kind), (px, py) in zip(shown, positions):
            iid = new_id("i")
            emit(iid, label, S_ICON + f"image={icon.get(kind, icon['vm'])};",
                 SUB_PAD + px + off, SUB_HEADER + py, ICON_W, ICON_H, parent_id)
            key = str(label).strip().lower()
            if key and key not in icon_id_by_name:
                icon_id_by_name[key] = iid

    # ---- group VNets by region, build subnet→resource membership ----
    # VMs (and VMSS) placed into their subnet; everything else → inventory.
    vnet_by_name = {}
    for idx, v in enumerate(networks):
        nm = str(v.get("name") or f"vnet-{idx+1}").strip()
        vnet_by_name[nm.lower()] = {
            "name": nm,
            "region": (str(v.get("location") or "").strip() or "Unknown Region"),
            "subscription": (str(v.get("subscription") or "").strip() or "Unknown subscription"),
            "cidr": _clean_cidr(v.get("address_space")),
            "subnets": {},  # subnet_lower -> {name, cidr, nsg, icons:[]}
        }

    def ensure_vnet(vlower, fallback_name=""):
        if vlower not in vnet_by_name:
            vnet_by_name[vlower] = {
                "name": fallback_name or vlower or "unknown-vnet",
                "region": "Unknown Region",
                "subscription": "Unknown subscription",
                "cidr": "",
                "subnets": {},
            }
        return vnet_by_name[vlower]

    def ensure_subnet(v, sname, cidr="", nsg=""):
        sl = str(sname or "default").strip().lower() or "default"
        if sl not in v["subnets"]:
            v["subnets"][sl] = {"name": str(sname or "default").strip() or "default",
                                "cidr": cidr, "nsg": nsg, "icons": []}
        return v["subnets"][sl]

    for s in subnets:
        vl = str(s.get("vnet") or "").strip().lower()
        if not vl:
            continue
        v = ensure_vnet(vl)
        ensure_subnet(v, s.get("name"), _clean_cidr(s.get("address_prefix")), str(s.get("nsg") or ""))

    # AKS node-pool VMs (names like aks-* and/or living in an MC_* node resource
    # group) are managed infrastructure. They would otherwise swamp their subnet
    # with dozens of identical nodes, so we divert them into a collapsed
    # "AKS Nodes" box per cluster instead of placing them in the topology.
    node_rg_to_cluster = {
        str(a.get("node_resource_group") or "").strip().lower(): str(a.get("name") or "").strip()
        for a in aks_clusters if str(a.get("node_resource_group") or "").strip()
    }
    aks_node_buckets = defaultdict(list)   # bucket label → [node names]

    def _is_aks_node(r):
        nm = str(r.get("name") or "").strip().lower()
        rg = str(r.get("resource_group") or "").strip().lower()
        return nm.startswith("aks-") or rg.startswith("mc_")

    def place_compute(items, kind, collapse_aks_nodes=False):
        for r in items:
            if collapse_aks_nodes and _is_aks_node(r):
                rg = str(r.get("resource_group") or "").strip().lower()
                label = node_rg_to_cluster.get(rg) or (str(r.get("resource_group") or "").strip() or "AKS nodes")
                aks_node_buckets[label].append(str(r.get("name") or "aks-node").strip())
                continue
            vl = str(r.get("vnet") or "").strip().lower()
            if not vl:
                continue
            v = ensure_vnet(vl)
            sub = ensure_subnet(v, r.get("subnet") or "default")
            sub["icons"].append((str(r.get("name") or kind).strip(), kind))

    place_compute(compute, "vm", collapse_aks_nodes=True)
    place_compute(vm_scale_sets, "vmss", collapse_aks_nodes=True)

    # Private endpoints carry their own vnet+subnet → place them in that subnet
    # when known; the rest fall back to the PaaS inventory so none are lost.
    unplaced_pe = []
    for pe in private_endpoints:
        vl = str(pe.get("vnet") or "").strip().lower()
        if not vl or vl not in vnet_by_name:
            unplaced_pe.append(pe)
            continue
        v = ensure_vnet(vl)
        sub = ensure_subnet(v, pe.get("subnet") or "default")
        sub["icons"].append((str(pe.get("name") or "private-endpoint").strip(), "pe"))

    # AKS: infer the VNet/subnet it is integrated into from the NICs created in
    # its node resource group (MC_*), so a VNet-integrated cluster shows inside
    # its subnet rather than in the PaaS inventory.
    nic_rg_vnet = defaultdict(Counter)
    nic_rg_subnet = defaultdict(Counter)
    for ni in network_interfaces:
        rg = str(ni.get("resource_group") or "").strip().lower()
        vn = str(ni.get("vnet") or "").strip()
        sn = str(ni.get("subnet") or "").strip()
        if rg and vn:
            nic_rg_vnet[rg][vn] += 1
        if rg and sn:
            nic_rg_subnet[rg][sn] += 1

    def _dominant(counter):
        return counter.most_common(1)[0][0] if counter else ""

    placed_aks = []   # AKS objects successfully placed in a subnet
    for aks in aks_clusters:
        nrg = str(aks.get("node_resource_group") or "").strip().lower()
        rg = str(aks.get("resource_group") or "").strip().lower()
        hint = str(aks.get("network_hint") or "").strip().lower()
        vn = hint or _dominant(nic_rg_vnet.get(nrg)) or _dominant(nic_rg_vnet.get(rg))
        sn = _dominant(nic_rg_subnet.get(nrg)) or _dominant(nic_rg_subnet.get(rg))
        # Fallback: a subnet whose name signals AKS/Kubernetes.
        if not vn:
            for vl, vd in vnet_by_name.items():
                for sl in vd["subnets"]:
                    if "aks" in sl or "kubernet" in sl:
                        vn, sn = vd["name"], vd["subnets"][sl]["name"]
                        break
                if vn:
                    break
        if vn:
            v = ensure_vnet(vn.lower())
            sub = ensure_subnet(v, sn or "aks-subnet")
            sub["icons"].append((str(aks.get("name") or "aks").strip(), "aks"))
            placed_aks.append(aks)

    # ---- compute sizes bottom-up, emit Subscription → Region → VNet → Subnet ----
    cursor_y = MARGIN
    vnet_cell_id = {}     # vnet name (lower) → vnet cell id (peering edges)
    subnet_cell_id = {}   # (vnet_lower, subnet_lower) → subnet cell id (flow edges)

    S_SUBSCRIPTION = _CBASE + "fillColor=#FFFCF0;strokeColor=#D9A400;fontStyle=1;fontSize=15;"

    def size_vnet(v):
        sub_items = sorted(v["subnets"].values(), key=lambda d: d["name"].lower())
        if not sub_items:
            sub_items = [{"name": "(no subnets)", "cidr": "", "nsg": "", "icons": []}]
        sub_sizes, sub_meta = [], []
        for sub in sub_items:
            cw, ch = box_size_for_icons(len(sub["icons"]))
            sw = max(MIN_SUBNET_W, SUB_PAD * 2 + cw)
            sh = SUB_HEADER + (SUB_PAD + ch + SUB_PAD if ch > 0 else SUB_PAD)
            sub_sizes.append((sw, sh))
            sub_meta.append(sub)
        positions, cw, ch = _shelf_pack(sub_sizes, VNET_MAXW, G)
        vnet_w = max(MIN_VNET_W, VNET_PAD * 2 + cw)
        vnet_h = VNET_HEADER + VNET_PAD + ch + VNET_PAD
        return vnet_w, vnet_h, list(zip(sub_meta, sub_sizes, positions))

    def size_region(vnets):
        renders = []
        for v in sorted(vnets, key=lambda d: d["name"].lower()):
            vw, vh, sl = size_vnet(v)
            renders.append((v, vw, vh, sl))
        positions, cw, ch = _shelf_pack([(vw, vh) for (_, vw, vh, _) in renders], REGION_MAXW, G)
        region_w = max(MIN_VNET_W, REGION_PAD * 2 + cw)
        region_h = REGION_HEADER + REGION_PAD + ch + REGION_PAD
        return region_w, region_h, list(zip(renders, positions))

    def emit_vnet(vdata, vw, vh, sub_layout, ox, oy, parent_id):
        vlabel = vdata["name"] + (f"  {vdata['cidr']}" if vdata["cidr"] else "")
        vnet_id = new_id("v")
        vnet_cell_id[vdata["name"].lower()] = vnet_id
        emit(vnet_id, vlabel, S_VNET, ox, oy, vw, vh, parent_id)
        for sub, (sw, sh), (spx, spy) in sub_layout:
            slabel = sub["name"] + (f"  {sub['cidr']}" if sub["cidr"] else "")
            sid = new_id("s")
            subnet_cell_id[(vdata["name"].lower(), sub["name"].lower())] = sid
            emit(sid, slabel, S_SUBNET, VNET_PAD + spx, VNET_HEADER + spy, sw, sh, vnet_id)
            emit_icons(sub["icons"], sid)

    def emit_region(region_name, rw, rh, vnet_layout, ox, oy, parent_id):
        rid = new_id("r")
        emit(rid, region_name, S_REGION, ox, oy, rw, rh, parent_id)
        for (vdata, vw, vh, sl), (vpx, vpy) in vnet_layout:
            emit_vnet(vdata, vw, vh, sl, REGION_PAD + vpx, REGION_HEADER + vpy, rid)

    # Group VNets: Subscription → Region → [VNets].
    by_sub = defaultdict(lambda: defaultdict(list))
    for v in vnet_by_name.values():
        by_sub[v["subscription"]][v["region"]].append(v)

    for sub_name in sorted(by_sub.keys()):
        region_sizes, region_layouts = [], []
        for region_name in sorted(by_sub[sub_name].keys()):
            rw, rh, vl = size_region(by_sub[sub_name][region_name])
            region_sizes.append((rw, rh))
            region_layouts.append((region_name, rw, rh, vl))
        # Stack regions vertically inside the subscription (one per row).
        rpos, cw, ch = _shelf_pack(region_sizes, 1, REGION_GAP)
        sub_w = max(MIN_VNET_W, REGION_PAD * 2 + cw)
        sub_h = REGION_HEADER + REGION_PAD + ch + REGION_PAD
        sub_id = new_id("sub")
        emit(sub_id, f"SUBSCRIPTION · {sub_name}", S_SUBSCRIPTION, MARGIN, cursor_y, sub_w, sub_h, "1")
        for (region_name, rw, rh, vl), (rx, ry) in zip(region_layouts, rpos):
            emit_region(region_name, rw, rh, vl, REGION_PAD + rx, REGION_HEADER + ry, sub_id)
        cursor_y += sub_h + REGION_GAP

    # ---- inventory region for non-VNet-bound resources (PaaS / data / edge) ----
    # Private endpoints and VNet-integrated AKS were already placed into their
    # subnets above, so they are excluded here to avoid double-counting.
    placed_aks_keys = {
        f"{str(a.get('name') or '').strip().lower()}|{str(a.get('resource_group') or '').strip().lower()}"
        for a in placed_aks
    }
    unplaced_aks = [
        a for a in aks_clusters
        if f"{str(a.get('name') or '').strip().lower()}|{str(a.get('resource_group') or '').strip().lower()}"
        not in placed_aks_keys
    ]
    inv_groups = [
        ("Web Apps / App Services", web_apps, "web"),
        ("Databases", databases, "sql"),
        ("Storage Accounts", storage_accounts, "storage"),
        ("Private Endpoints", unplaced_pe, "pe"),
        ("Load Balancers", load_balancers, "lb"),
        ("Application Gateways", application_gateways, "appgw"),
        ("AKS Clusters", unplaced_aks, "aks"),
    ]
    inv_present = [(t, items, k) for (t, items, k) in inv_groups if items]
    if inv_present:
        box_sizes = []
        box_meta = []
        for title, items, kind in inv_present:
            icons = [(str(r.get("name") or title).strip(), kind) for r in items]
            cw, ch = box_size_for_icons(len(icons))
            bw = max(MIN_SUBNET_W, SUB_PAD * 2 + cw)
            bh = SUB_HEADER + (SUB_PAD + ch + SUB_PAD if ch > 0 else SUB_PAD)
            box_sizes.append((bw, bh))
            box_meta.append((title, icons))
        positions, content_w, content_h = _shelf_pack(box_sizes, REGION_MAXW, G)
        region_w = max(MIN_VNET_W, REGION_PAD * 2 + content_w)
        region_h = REGION_HEADER + REGION_PAD + content_h + REGION_PAD
        region_id = new_id("r")
        emit(region_id, "Managed Services & PaaS (inventory)", S_INV_REGION,
             MARGIN, cursor_y, region_w, region_h, "1")
        for (title, icons), (bw, bh), (bpx, bpy) in zip(box_meta, box_sizes, positions):
            extra = f"  (+{len(icons) - MAX_ICONS_PER_BOX} more)" if len(icons) > MAX_ICONS_PER_BOX else ""
            box_id = new_id("s")
            emit(box_id, f"{title} ({len(icons)}){extra}", S_SUBNET,
                 REGION_PAD + bpx, REGION_HEADER + bpy, bw, bh, region_id)
            emit_icons(icons, box_id)
        cursor_y += region_h + REGION_GAP

    # ---- AKS node pools: one collapsed box per cluster (nodes hidden until expanded) ----
    if aks_node_buckets:
        NODE_COLS = SUB_COLS
        boxes = []
        for label in sorted(aks_node_buckets.keys()):
            names = aks_node_buckets[label]
            sizes = [(CELL_W, CELL_H)] * len(names)
            positions, cw, ch = _shelf_pack(sizes, NODE_COLS * (CELL_W + G), G)
            ew = max(MIN_SUBNET_W, SUB_PAD * 2 + cw)
            eh = SUB_HEADER + SUB_PAD + ch + SUB_PAD
            title = f"AKS Nodes: {label} ({len(names)})"
            cw_collapsed = max(220, min(380, 9 * len(title)))
            ch_collapsed = 30
            boxes.append((title, names, positions, ew, eh, cw_collapsed, ch_collapsed))

        csizes = [(b[5], b[6]) for b in boxes]
        cpos, content_w, content_h = _shelf_pack(csizes, REGION_MAXW, G)
        region_w = max(MIN_VNET_W, REGION_PAD * 2 + content_w)
        region_h = REGION_HEADER + REGION_PAD + content_h + REGION_PAD
        region_id = new_id("r")
        emit(region_id, "AKS Node Pools  (collapsed — double-click a box to expand)",
             S_AKS_REGION, MARGIN, cursor_y, region_w, region_h, "1")
        for (title, names, positions, ew, eh, cwth, chth), (bx, by) in zip(boxes, cpos):
            box_id = new_id("k")
            emit_collapsed(box_id, title, S_AKS_BOX,
                           REGION_PAD + bx, REGION_HEADER + by, cwth, chth, ew, eh, region_id)
            off = (CELL_W - ICON_W) // 2
            for nm, (px, py) in zip(names, positions):
                emit(new_id("i"), nm, S_ICON + f"image={icon['vm']};",
                     SUB_PAD + px + off, SUB_HEADER + py, ICON_W, ICON_H, box_id)
        cursor_y += region_h + REGION_GAP

    # ---- edges: VNet peering + resource-to-resource connections ----
    if arrows:
        import ast as _ast, re as _re

        def _edge(src, tgt, style, label=""):
            eid = new_id("e")
            lbl = f' value="{_xml_escape(label)}"' if label else ' value=""'
            cells.append(
                f'<mxCell id="{eid}"{lbl} style="{_xml_escape(style)}" edge="1" parent="1" '
                f'source="{src}" target="{tgt}">'
                f'<mxGeometry relative="1" as="geometry"/></mxCell>'
            )

        def _as_list(raw):
            if isinstance(raw, list):
                return raw
            if isinstance(raw, str) and raw.strip().startswith("["):
                try:
                    return _ast.literal_eval(raw)
                except Exception:
                    return []
            return []

        # NIC name → owning VM name, parsed from each VM's network_interfaces blob.
        # Lets us resolve load-balancer backends that reference NICs, not VMs.
        nic_to_vm = {}
        for v in compute:
            owner = str(v.get("name") or "").strip()
            for nicname in _re.findall(r"networkinterfaces/([^/'\"\]}]+)", str(v.get("network_interfaces") or ""), _re.I):
                nic_to_vm[nicname.strip().lower()] = owner

        def _resolve_icon(name):
            k = str(name or "").strip().lower()
            if not k:
                return None
            if k in icon_id_by_name:
                return icon_id_by_name[k]
            # strip common NIC/interface suffixes (e.g. "vm1-eth2", "vm_nic0")
            k2 = _re.sub(r"[-_](eth\d+|nic\d*|ipconfig\d*|primary|internal|external)$", "", k)
            if k2 in icon_id_by_name:
                return icon_id_by_name[k2]
            # NIC → owning VM
            vm = nic_to_vm.get(k) or nic_to_vm.get(k2)
            if vm and vm.strip().lower() in icon_id_by_name:
                return icon_id_by_name[vm.strip().lower()]
            return None

        # 1) VNet peering (dashed orange, no arrowhead) — matches the template.
        seen_pairs = set()
        for v in networks:
            a = str(v.get("name") or "").strip().lower()
            if a not in vnet_cell_id:
                continue
            for p in _as_list(v.get("peerings")):
                if not isinstance(p, dict):
                    continue
                props = p.get("properties") or {}
                rvn = (props.get("remoteVirtualNetwork") or {}) if isinstance(props, dict) else {}
                rid = str(rvn.get("id") or "")
                if "/virtualnetworks/" not in rid.lower():
                    continue
                b = rid.rstrip("/").split("/")[-1].strip().lower()
                if b not in vnet_cell_id or b == a:
                    continue
                pair = tuple(sorted((a, b)))
                if pair in seen_pairs:
                    continue
                seen_pairs.add(pair)
                _edge(vnet_cell_id[a], vnet_cell_id[b], E_PEER)

        # 2) Resource connections (solid green arrows), capped to stay readable.
        seen_conn = set()
        conn_count = 0
        CONN_CAP = 500

        def _conn(src, tgt):
            nonlocal conn_count
            if not src or not tgt or src == tgt or conn_count >= CONN_CAP:
                return
            key = (src, tgt)
            if key in seen_conn:
                return
            seen_conn.add(key)
            _edge(src, tgt, E_CONN)
            conn_count += 1

        # Load balancer → backend members
        for lb in load_balancers:
            sid = icon_id_by_name.get(str(lb.get("name") or "").strip().lower())
            if not sid:
                continue
            for t in _as_list(lb.get("backend_targets")):
                _conn(sid, _resolve_icon(t))

        # Private endpoint → linked target resource
        for pe in private_endpoints:
            sid = icon_id_by_name.get(str(pe.get("name") or "").strip().lower())
            if not sid:
                continue
            blob = str(pe.get("linked_resource") or pe.get("linked_resources") or "")
            targets = _re.findall(r"/providers/[^/]+/[^/]+/([^/'\"\]}]+)", blob)
            for nm in targets[:1]:
                _conn(sid, _resolve_icon(nm))

        # AKS internal load balancers → their AKS cluster(s).
        aks_icon_ids = [icon_id_by_name[k] for k in (
            str(a.get("name") or "").strip().lower() for a in aks_clusters
        ) if k in icon_id_by_name]
        for lb in load_balancers:
            if "kubernetes" in str(lb.get("name") or "").lower():
                sid = icon_id_by_name.get(str(lb.get("name") or "").strip().lower())
                for aid in aks_icon_ids:
                    _conn(sid, aid)

        # 3) Logical flow arrows (blue) inferred from subnet tiers:
        #    edge/public → web/app → app → db, within each VNet. These are
        #    inferred from naming (not verified dependencies) but convey flow.
        E_FLOW = ("edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;endArrow=block;endFill=1;"
                  "strokeColor=#2E75B6;strokeWidth=1.3;")

        def _tier(sname):
            s = str(sname).lower()
            if any(k in s for k in ("db", "data", "sql", "postgre", "mysql", "mongo", "cache", "redis")):
                return "db"
            if any(k in s for k in ("appgw", "gateway", "waf", "ngfw", "firewall", "dmz",
                                    "perimeter", "public", "frontend", "edge", "bastion")):
                return "edge"
            if "web" in s:
                return "web"
            if any(k in s for k in ("app", "api", "aks", "kubernet", "compute", "svc", "service", "mid")):
                return "app"
            return ""

        flow = [0]
        FLOW_CAP = 400

        def _flow(a_list, b_list):
            for a in a_list:
                for b in b_list:
                    if a != b and flow[0] < FLOW_CAP:
                        _edge(a, b, E_FLOW)
                        flow[0] += 1

        for vl in vnet_by_name.keys():
            tiers = {"edge": [], "web": [], "app": [], "db": []}
            for (vv, sl), sid in subnet_cell_id.items():
                if vv == vl:
                    t = _tier(sl)
                    if t:
                        tiers[t].append(sid)
            _flow(tiers["edge"], tiers["web"] or tiers["app"])
            _flow(tiers["web"], tiers["app"])
            _flow(tiers["app"], tiers["db"])

    body = "".join(cells)
    total_w = 2000
    total_h = cursor_y + MARGIN
    return (
        '<mxfile host="app.diagrams.net">'
        '<diagram name="Azure Architecture" id="azure-arch">'
        f'<mxGraphModel dx="1200" dy="800" grid="1" gridSize="10" guides="1" '
        f'tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" '
        f'pageWidth="{total_w}" pageHeight="{int(total_h)}" math="0" shadow="0">'
        '<root><mxCell id="0"/><mxCell id="1" parent="0"/>'
        f'{body}</root></mxGraphModel></diagram></mxfile>'
    )


def _azure_icon_map():
    """Shared Azure icon URI map (resolved from local Azure-Icons/)."""
    AZ = "https://www.azureicons.com/static/images/icons"
    return {
        "vm": get_icon_uri("Compute", "Virtual-Machine", f"{AZ}/Compute/svg/Virtual-Machine.svg"),
        "vmss": get_icon_uri("Compute", "VM-Scale-Sets", f"{AZ}/Compute/svg/VM-Scale-Sets.svg"),
        "aks": get_icon_uri("Containers", "Kubernetes-Services", f"{AZ}/Containers/svg/Kubernetes-Services.svg"),
        "lb": get_icon_uri("Networking", "Load-Balancers", f"{AZ}/Networking/svg/Load-Balancers.svg"),
        "appgw": get_icon_uri("Networking", "Application-Gateways", f"{AZ}/Networking/svg/Application-Gateways.svg"),
        "pe": get_icon_uri("Networking", "Private-Link", f"{AZ}/Networking/svg/Private-Link.svg"),
        "sql": get_icon_uri("Databases", "Azure-SQL", f"{AZ}/Databases/svg/Azure-SQL.svg"),
        "storage": get_icon_uri("Storage", "Storage-Accounts", f"{AZ}/Storage/svg/Storage-Accounts.svg"),
        "web": get_icon_uri("App-Services", "App-Services", f"{AZ}/App%20Services/svg/App-Services.svg"),
        "vnet": get_icon_uri("Networking", "Virtual-Networks", f"{AZ}/Networking/svg/Virtual-Networks.svg"),
    }


def _build_azure_drawio_rgwise(summary: Dict) -> str:
    """Resource-group-oriented draw.io layout for Azure.

    Groups every resource under its Resource Group. Each RG is an auto-sized
    container holding all of its resources as icons, packed in a grid so
    nothing overlaps. RG boxes are themselves packed into rows.
    """
    ICON_W, ICON_H = 60, 74
    G = 14
    PAD, HEADER = 12, 42
    RG_ICON_COLS = 10            # icons per row inside an RG box
    MIN_RG_W = 200
    GRID_MAXW = 3600             # wrap RG boxes past this width
    MARGIN, RG_GAP = 40, 40

    icon = _azure_icon_map()
    S_RG = "rounded=1;arcSize=3;whiteSpace=wrap;html=1;verticalAlign=top;fontStyle=1;fontSize=13;fillColor=#eef2f7;strokeColor=#6c8ebf;strokeWidth=1;spacingTop=4;"
    S_ICON = (
        "shape=image;verticalLabelPosition=bottom;verticalAlign=top;align=center;"
        "imageAspect=0;aspect=fixed;whiteSpace=wrap;html=1;labelBackgroundColor=none;"
        "fontSize=8;strokeColor=none;fillColor=none;"
    )

    buckets = defaultdict(list)   # rg name -> [(label, kind)]

    def add(items, kind):
        for r in items:
            if not isinstance(r, dict):
                continue
            rg = str(r.get("resource_group") or "").strip() or "(no resource group)"
            buckets[rg].append((str(r.get("name") or kind).strip(), kind))

    add(summary.get("compute", []) or [], "vm")
    add(summary.get("vm_scale_sets", []) or [], "vmss")
    add(summary.get("aks_clusters", []) or [], "aks")
    add(summary.get("web_apps", []) or [], "web")
    add(summary.get("database_services", []) or [], "sql")
    add(summary.get("storage_accounts", []) or [], "storage")
    add(summary.get("private_endpoints", []) or [], "pe")
    add(summary.get("load_balancers", []) or [], "lb")
    add(summary.get("application_gateways", []) or [], "appgw")

    cells = []
    _cid = {"n": 0}

    def new_id(prefix="c"):
        _cid["n"] += 1
        return f"{prefix}{_cid['n']}"

    def emit(cid, value, style, x, y, w, h, parent, edge=False):
        cells.append(
            f'<mxCell id="{cid}" value="{_xml_escape(value)}" style="{_xml_escape(style)}" '
            f'vertex="1" parent="{parent}">'
            f'<mxGeometry x="{int(x)}" y="{int(y)}" width="{int(w)}" height="{int(h)}" as="geometry"/>'
            f'</mxCell>'
        )

    max_row_w = RG_ICON_COLS * (ICON_W + G)
    rg_render = []  # (rg_name, icons, rg_w, rg_h, positions)
    for rg_name in sorted(buckets.keys()):
        icons = buckets[rg_name]
        sizes = [(ICON_W, ICON_H)] * len(icons)
        positions, cw, ch = _shelf_pack(sizes, max_row_w, G)
        rg_w = max(MIN_RG_W, PAD * 2 + (cw if icons else ICON_W))
        rg_h = HEADER + PAD + (ch if icons else 0) + PAD
        rg_render.append((rg_name, icons, rg_w, rg_h, positions))

    rg_sizes = [(w, h) for (_, _, w, h, _) in rg_render]
    rg_positions, _, _ = _shelf_pack(rg_sizes, GRID_MAXW, RG_GAP)

    for (rg_name, icons, rg_w, rg_h, positions), (rx, ry) in zip(rg_render, rg_positions):
        rg_id = new_id("g")
        emit(rg_id, f"{rg_name} ({len(icons)})", S_RG, MARGIN + rx, MARGIN + ry, rg_w, rg_h, "1")
        for (label, kind), (px, py) in zip(icons, positions):
            emit(new_id("i"), label, S_ICON + f"image={icon.get(kind, icon['vm'])};",
                 PAD + px, HEADER + py, ICON_W, ICON_H, rg_id)

    total_h = MARGIN + (max((ry + h) for (_, _, _, h, _), (_, ry) in zip(rg_render, rg_positions)) if rg_render else 0) + MARGIN
    body = "".join(cells)
    return (
        '<mxfile host="app.diagrams.net">'
        '<diagram name="Azure by Resource Group" id="azure-rg">'
        f'<mxGraphModel dx="1200" dy="800" grid="1" gridSize="10" guides="1" '
        f'tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" '
        f'pageWidth="2000" pageHeight="{int(total_h)}" math="0" shadow="0">'
        '<root><mxCell id="0"/><mxCell id="1" parent="0"/>'
        f'{body}</root></mxGraphModel></diagram></mxfile>'
    )


def _build_drawio_network_xml(summary: Dict) -> str:
    """
    Build native draw.io XML for a clean network architecture layout.
    Dispatches by cloud — Azure uses the deterministic nested layout
    (Region → VNet → Subnet → resources), AWS uses drawio's built-in AWS4
    stencil library (no external files needed).
    """
    cloud = str(summary.get("cloud") or "").lower()
    if cloud == "aws":
        return _build_aws_drawio_network_xml(summary)
    if cloud != "azure":
        return ""

    # Azure uses the deterministic nested auto-sizing layout.
    return _build_azure_drawio_nested(summary)

    networks = summary.get("networks", []) or []
    subnets = summary.get("subnets", []) or []
    compute = summary.get("compute", []) or []
    aks_clusters = summary.get("aks_clusters", []) or []
    load_balancers = summary.get("load_balancers", []) or []
    network_interfaces = summary.get("network_interfaces", []) or []
    storage_accounts = summary.get("storage_accounts", []) or summary.get("storage", []) or []
    databases = summary.get("database_services", []) or summary.get("databases", []) or []
    private_endpoints = summary.get("private_endpoints", []) or []
    application_gateways = summary.get("application_gateways", []) or []
    vm_scale_sets = summary.get("vm_scale_sets", []) or []

    # Index by VNet.
    vnet_to_subnets: Dict[str, List[dict]] = defaultdict(list)
    for s in subnets:
        vnet_to_subnets[str(s.get("vnet") or "").strip().lower()].append(s)

    vnet_to_vms: Dict[str, List[dict]] = defaultdict(list)
    for vm in compute:
        vnet_to_vms[str(vm.get("vnet") or "").strip().lower()].append(vm)
    for vmss in vm_scale_sets:
        vnet_to_vms[str(vmss.get("vnet") or "").strip().lower()].append(vmss)

    # If any VMSS lacked a VNet mapping, try to infer it from the raw_row subnet id
    # and reconcile with known `subnets` so the VMSS is placed into the correct VNet.
    def _extract_subnet_id_from_raw(raw):
        if not raw:
            return None
        # Search recursively for a string containing both virtualNetworks and /subnets/
        def _search(obj):
            if obj is None:
                return None
            if isinstance(obj, str):
                s = obj.strip()
                if s.lower().startswith("/subscriptions/") and "/virtualnetworks/" in s.lower() and "/subnets/" in s.lower():
                    return s
                return None
            if isinstance(obj, dict):
                for v in obj.values():
                    res = _search(v)
                    if res:
                        return res
            if isinstance(obj, list):
                for it in obj:
                    res = _search(it)
                    if res:
                        return res
            return None
        return _search(raw)

    # Build a quick lookup of subnet IDs (normalized lower) -> subnet record
    subnet_id_map: Dict[str, dict] = {}
    for s in subnets:
        sid = str(s.get("id") or "").strip().lower()
        if sid:
            subnet_id_map[sid] = s

    # Reassign vmss without vnet using subnet id from their raw_row
    unassigned = list(vnet_to_vms.get("", []))
    for vmss in unassigned:
        if vmss.get("vnet"):
            continue
        raw = vmss.get("raw_row")
        sid = _extract_subnet_id_from_raw(raw)
        if not sid:
            continue
        sid_norm = sid.strip().lower()
        matched = None
        # Exact match first
        if sid_norm in subnet_id_map:
            matched = subnet_id_map[sid_norm]
        else:
            # Partial-match by subnet name suffix
            for k, srec in subnet_id_map.items():
                if k.endswith(sid_norm) or sid_norm.endswith(k):
                    matched = srec
                    break
        if matched:
            vmss["vnet"] = str(matched.get("vnet") or "").strip()
            vmss["subnet"] = str(matched.get("name") or "").strip()
            # Move entry to proper vnet bucket
            try:
                # remove from empty bucket and add to correct bucket
                vnet_to_vms[""].remove(vmss)
            except Exception:
                pass
            vnet_to_vms[str(vmss.get("vnet") or "").strip().lower()].append(vmss)

    # Infer RG -> VNet/Subnet from NIC data to place AKS and managed LB accurately.
    rg_to_vnet_counts: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    rg_to_subnet_counts: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for ni in network_interfaces:
        rg = str(ni.get("resource_group") or "").strip().lower()
        vnet = str(ni.get("vnet") or "").strip()
        subnet = str(ni.get("subnet") or "").strip()
        if not rg:
            continue
        if vnet:
            rg_to_vnet_counts[rg][vnet] += 1
        if subnet:
            rg_to_subnet_counts[rg][subnet] += 1

    def _dominant_from_counts(counts: Dict[str, int]) -> str:
        if not counts:
            return ""
        return max(counts.items(), key=lambda kv: kv[1])[0]

    # Pre-compute layout dimensions so boxes don't overflow.
    max_per_row = 2
    vnet_w = 460
    gap_x = 40
    gap_y = 30
    subnet_h = 198
    vm_max_per_subnet = 4

    # Extra height added to any subnet whose name contains 'aks' or 'kubernetes'
    # to safely accommodate the AKS cluster tile below the VM grid.
    aks_subnet_extra_h = 170   # pixels added on top of the normal subnet_h
    # Pre-scan all subnet names so we can tag AKS subnets before building vnet_layout.
    def _is_aks_subnet(name: str) -> bool:
        nl = str(name or "").lower()
        return "aks" in nl or "kubernetes" in nl

    vnet_layout = []
    for idx, v in enumerate(networks):
        vname = str(v.get("name") or f"vnet-{idx+1}").strip()
        v_subnets = vnet_to_subnets.get(vname.lower(), [])
        if v_subnets:
            subnet_entries = []
            for s in v_subnets:
                subnet_entries.append({
                    "name": str(s.get("name") or "default").strip() or "default",
                    "nsg": str(s.get("nsg") or "").strip(),
                })
        else:
            vm_subnets = {
                str(vm.get("subnet") or "").strip() or "default"
                for vm in vnet_to_vms.get(vname.lower(), [])
            }
            subnet_names = sorted(vm_subnets) if vm_subnets else ["default"]
            subnet_entries = [{"name": n, "nsg": ""} for n in subnet_names]

        # De-dup by name while preserving first NSG.
        seen_sn = set()
        dedup_subnets = []
        for s in subnet_entries:
            key = str(s.get("name") or "").lower()
            if key in seen_sn:
                continue
            seen_sn.add(key)
            dedup_subnets.append(s)

        # Per-subnet height: AKS subnets get extra room for the AKS cluster tile.
        per_subnet_h = [
            (subnet_h + aks_subnet_extra_h) if _is_aks_subnet(s.get("name", "")) else subnet_h
            for s in dedup_subnets
        ]
        vnet_h = 74 + sum(per_subnet_h) + 24
        vnet_layout.append({
            "index": idx,
            "name": vname,
            "subnets": dedup_subnets,
            "height": vnet_h,
            "per_subnet_h": per_subnet_h,
        })

    vnet_names_by_lower = {str(v["name"]).strip().lower(): str(v["name"]).strip() for v in vnet_layout}

    def _match_vnet_name(vnet_hint: str) -> str:
        hint = str(vnet_hint or "").strip().lower()
        if not hint:
            return ""
        if hint in vnet_names_by_lower:
            return vnet_names_by_lower[hint]
        for k, actual in vnet_names_by_lower.items():
            if hint in k or k in hint:
                return actual
        return ""

    vnet_to_subnet_names: Dict[str, List[str]] = defaultdict(list)
    for v in vnet_layout:
        vk = str(v.get("name") or "").strip().lower()
        for s in (v.get("subnets") or []):
            sn = str(s.get("name") or "").strip()
            if sn:
                vnet_to_subnet_names[vk].append(sn)

    def _match_subnet_name(vnet_lower: str, subnet_hint: str) -> str:
        names = vnet_to_subnet_names.get(vnet_lower, [])
        if not names:
            return ""
        hint = str(subnet_hint or "").strip().lower()
        if hint:
            for n in names:
                if hint == n.lower():
                    return n
            for n in names:
                nl = n.lower()
                if hint in nl or nl in hint:
                    return n
        return names[0]

    subnet_to_aks: Dict[str, List[dict]] = defaultdict(list)
    for aks in aks_clusters:
        node_rg = str(aks.get("node_resource_group") or "").strip().lower()
        rg = str(aks.get("resource_group") or "").strip().lower()
        vnet_hint = str(aks.get("network_hint") or "").strip()
        subnet_hint = str(aks.get("subnet_hint") or "").strip()
        if not vnet_hint and node_rg:
            vnet_hint = _dominant_from_counts(rg_to_vnet_counts.get(node_rg, {}))
        if not vnet_hint and rg:
            vnet_hint = _dominant_from_counts(rg_to_vnet_counts.get(rg, {}))
        if not subnet_hint and node_rg:
            subnet_hint = _dominant_from_counts(rg_to_subnet_counts.get(node_rg, {}))
        if not subnet_hint and rg:
            subnet_hint = _dominant_from_counts(rg_to_subnet_counts.get(rg, {}))
        vname = _match_vnet_name(vnet_hint)
        if not vname:
            continue
        sname = _match_subnet_name(vname.lower(), subnet_hint)
        key = f"{vname.lower()}::{sname.lower()}"
        subnet_to_aks[key].append(aks)

    # Fallback: if no AKS clusters were mapped via hint, assign ALL to the first
    # subnet whose name contains 'aks' or 'kubernetes' (works for snet-*-platform-aks).
    if not subnet_to_aks and aks_clusters:
        for vl in vnet_layout:
            vk = str(vl.get("name") or "").strip().lower()
            for s in (vl.get("subnets") or []):
                sn = str(s.get("name") or "").strip()
                if _is_aks_subnet(sn):
                    key = f"{vk}::{sn.lower()}"
                    subnet_to_aks[key] = list(aks_clusters)
                    break
            if subnet_to_aks:
                break

    # If mapping is partial, append remaining AKS clusters to the primary AKS subnet
    # so the diagram still shows all detected clusters.
    mapped_aks_keys = set()
    for items in subnet_to_aks.values():
        for aks in items:
            mapped_aks_keys.add(
                f"{str(aks.get('name') or '').strip().lower()}|{str(aks.get('resource_group') or '').strip().lower()}"
            )
    unmapped_aks = []
    for aks in aks_clusters:
        aks_key = f"{str(aks.get('name') or '').strip().lower()}|{str(aks.get('resource_group') or '').strip().lower()}"
        if aks_key not in mapped_aks_keys:
            unmapped_aks.append(aks)
    if unmapped_aks:
        target_key = ""
        for skey in subnet_to_aks.keys():
            if "aks" in skey or "kubernetes" in skey:
                target_key = skey
                break
        if not target_key:
            for vl in vnet_layout:
                vk = str(vl.get("name") or "").strip().lower()
                for s in (vl.get("subnets") or []):
                    sn = str(s.get("name") or "").strip()
                    if _is_aks_subnet(sn):
                        target_key = f"{vk}::{sn.lower()}"
                        break
                if target_key:
                    break
        if target_key:
            subnet_to_aks[target_key].extend(unmapped_aks)

    subnet_to_lbs: Dict[str, List[dict]] = defaultdict(list)
    for lb in load_balancers:
        rg = str(lb.get("resource_group") or "").strip().lower()
        vnet_hint = str(lb.get("network_hint") or "").strip()
        subnet_hint = ""
        if not vnet_hint and rg:
            vnet_hint = _dominant_from_counts(rg_to_vnet_counts.get(rg, {}))
        if rg:
            subnet_hint = _dominant_from_counts(rg_to_subnet_counts.get(rg, {}))
        vname = _match_vnet_name(vnet_hint)
        if not vname:
            continue
        sname = _match_subnet_name(vname.lower(), subnet_hint)
        key = f"{vname.lower()}::{sname.lower()}"
        subnet_to_lbs[key].append(lb)

    # Application Gateways ─ always deployed in a dedicated subnet (*appgw* / ApplicationGatewaySubnet).
    subnet_to_appgw: Dict[str, List[dict]] = defaultdict(list)
    for appgw in application_gateways:
        rg = str(appgw.get("resource_group") or "").strip().lower()
        vnet_hint = str(appgw.get("network_hint") or "").strip()
        if not vnet_hint and rg:
            vnet_hint = _dominant_from_counts(rg_to_vnet_counts.get(rg, {}))
        vname_ag = _match_vnet_name(vnet_hint)
        if not vname_ag:
            continue
        # Prefer a subnet whose name contains 'appgw' or 'applicationgateway'
        ag_subnet_names = vnet_to_subnet_names.get(vname_ag.lower(), [])
        sname_ag = ""
        for n in ag_subnet_names:
            nl = n.lower()
            if any(kw in nl for kw in ["appgw", "applicationgateway", "app-gw", "webgw", "gateway", "agw", "gw"]):
                sname_ag = n
                break
        if not sname_ag and rg:
            subnet_hint_ag = _dominant_from_counts(rg_to_subnet_counts.get(rg, {}))
            sname_ag = _match_subnet_name(vname_ag.lower(), subnet_hint_ag)
        if not sname_ag:
            continue
        key = f"{vname_ag.lower()}::{sname_ag.lower()}"
        subnet_to_appgw[key].append(appgw)

    row_heights = []
    for i in range(0, len(vnet_layout), max_per_row):
        row_items = vnet_layout[i:i + max_per_row]
        row_heights.append(max(item["height"] for item in row_items))
    total_vnet_h = sum(row_heights) + (max(0, len(row_heights) - 1) * gap_y)

    env_x = 430
    env_y = 20
    core_x = 470
    core_y = 70
    core_w = 1160
    core_h = max(620, total_vnet_h + 120)

    storage_x = core_x + core_w + 30
    storage_y = 70
    storage_w = 430
    storage_items = min(len(storage_accounts or []), 16) + min(len(databases or []), 8)
    storage_rows = max(1, (storage_items + 1) // 2)
    storage_h = max(380, 90 + (storage_rows * 96))

    env_w = (storage_x + storage_w + 40) - env_x
    env_h = max(core_h + 120, storage_h + 120)

    ingress_y = env_y + 230

    # Root XML structure.
    mxfile = ET.Element(
        "mxfile",
        {
            "host": "app.diagrams.net",
            "modified": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
            "agent": "AzureAssessment",
            "version": "24.7.17",
            "type": "device",
        },
    )
    diagram = ET.SubElement(mxfile, "diagram", {"id": "arch", "name": "Architecture"})
    model = ET.SubElement(
        diagram,
        "mxGraphModel",
        {
            "dx": "1200",
            "dy": "800",
            "grid": "1",
            "gridSize": "10",
            "guides": "1",
            "tooltips": "1",
            "connect": "1",
            "arrows": "1",
            "fold": "1",
            "page": "1",
            "pageScale": "1",
            "pageWidth": str(max(2200, env_w + 600)),
            "pageHeight": str(max(1400, env_h + 160)),
            "math": "0",
            "shadow": "0",
        },
    )
    root = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", {"id": "0"})
    ET.SubElement(root, "mxCell", {"id": "1", "parent": "0"})

    id_counter = 2

    def next_id() -> str:
        nonlocal id_counter
        cid = str(id_counter)
        id_counter += 1
        return cid

    # Track geometry of created cells so edges can be routed around boxes.
    cell_positions: Dict[str, Dict[str, int]] = {}

    def add_vertex(value: str, x: int, y: int, w: int, h: int, style: str, parent: str = "1") -> str:
        cid = next_id()
        c = ET.SubElement(
            root,
            "mxCell",
            {"id": cid, "value": value, "style": style, "vertex": "1", "parent": parent},
        )
        ET.SubElement(c, "mxGeometry", {"x": str(x), "y": str(y), "width": str(w), "height": str(h), "as": "geometry"})
        # Record position for routing
        try:
            cell_positions[cid] = {"x": int(x), "y": int(y), "w": int(w), "h": int(h)}
        except Exception:
            pass
        return cid

    def add_edge(source: str, target: str, style: str = "endArrow=block;endFill=1;strokeWidth=1;"):
        """Add an edge; if the straight path would cross any VNet/subnet boxes, add orthogonal waypoints to route it outside.
        This uses a heuristic: compute a safe X coordinate outside the union of VNet boxes and route via two waypoints.
        """
        cid = next_id()
        # Build vnet bounding boxes from subnet boxes (lazy grouping)
        vnet_boxes: List[Dict[str, int]] = []
        try:
            vnets = defaultdict(lambda: {"minx": 10**9, "miny": 10**9, "maxx": 0, "maxy": 0})
            for sk, pos in subnet_box_positions.items():
                vname = sk.split("::")[0]
                x = int(pos.get("x", 0))
                y = int(pos.get("y", 0))
                w = int(pos.get("w", 0))
                h = int(pos.get("h", 0))
                vnets[vname]["minx"] = min(vnets[vname]["minx"], x)
                vnets[vname]["miny"] = min(vnets[vname]["miny"], y)
                vnets[vname]["maxx"] = max(vnets[vname]["maxx"], x + w)
                vnets[vname]["maxy"] = max(vnets[vname]["maxy"], y + h)
            for v in vnets.values():
                if v["maxx"] > 0:
                    vnet_boxes.append({"x": v["minx"], "y": v["miny"], "w": v["maxx"] - v["minx"], "h": v["maxy"] - v["miny"]})
        except Exception:
            vnet_boxes = []

        def bbox_intersects(a, b):
            return not (a["x"] + a["w"] < b["x"] or b["x"] + b["w"] < a["x"] or a["y"] + a["h"] < b["y"] or b["y"] + b["h"] < a["y"])

        # Get source/target centers if available
        src = cell_positions.get(str(source))
        tgt = cell_positions.get(str(target))
        needs_routing = False
        if src and tgt and vnet_boxes:
            sx = src["x"] + (src.get("w", 0) / 2)
            sy = src["y"] + (src.get("h", 0) / 2)
            tx = tgt["x"] + (tgt.get("w", 0) / 2)
            ty = tgt["y"] + (tgt.get("h", 0) / 2)
            seg_bbox = {"x": int(min(sx, tx)), "y": int(min(sy, ty)), "w": int(abs(tx - sx)), "h": int(abs(ty - sy))}
            for vb in vnet_boxes:
                if bbox_intersects(seg_bbox, vb):
                    needs_routing = True
                    break

        # If routing not needed, add a simple edge
        if not needs_routing:
            c = ET.SubElement(
                root,
                "mxCell",
                {"id": cid, "edge": "1", "source": source, "target": target, "parent": "1", "style": style},
            )
            ET.SubElement(c, "mxGeometry", {"relative": "1", "as": "geometry"})
            return cid

        # Routing: choose an X coordinate outside all vnets (left or right) and create two orthogonal waypoints.
        try:
            all_minx = min(vb["x"] for vb in vnet_boxes)
            all_maxx = max(vb["x"] + vb["w"] for vb in vnet_boxes)
            # choose side closer to source
            sx = src["x"] + (src.get("w", 0) / 2)
            tx = tgt["x"] + (tgt.get("w", 0) / 2)
            left_out = int(all_minx - 60)
            right_out = int(all_maxx + 60)
            outside_x = left_out if abs(sx - left_out) < abs(sx - right_out) else right_out
            # Build edge with orthogonal style and waypoints
            routed_style = style + ";edgeStyle=orthogonalEdgeStyle;rounded=0;jettySize=auto;"
            c = ET.SubElement(
                root,
                "mxCell",
                {"id": cid, "edge": "1", "source": source, "target": target, "parent": "1", "style": routed_style},
            )
            geom = ET.SubElement(c, "mxGeometry", {"relative": "1", "as": "geometry"})
            # Waypoints must live inside an Array as="points" for draw.io
            points_arr = ET.SubElement(geom, "Array", {"as": "points"})
            ET.SubElement(points_arr, "mxPoint", {"x": str(outside_x), "y": str(int(src["y"] + (src.get("h", 0) / 2)))})
            ET.SubElement(points_arr, "mxPoint", {"x": str(outside_x), "y": str(int(tgt["y"] + (tgt.get("h", 0) / 2)))})
            return cid
        except Exception:
            # fallback
            c = ET.SubElement(
                root,
                "mxCell",
                {"id": cid, "edge": "1", "source": source, "target": target, "parent": "1", "style": style},
            )
            ET.SubElement(c, "mxGeometry", {"relative": "1", "as": "geometry"})
            return cid

    # Common styles.
    style_label = "rounded=1;whiteSpace=wrap;html=1;strokeColor=#6c8ebf;fillColor=#dae8fc;fontSize=11;"
    style_outer = "rounded=0;whiteSpace=wrap;html=1;strokeColor=#7f8c8d;fillColor=#f8f9fa;strokeWidth=1;"
    style_cloud = "rounded=0;whiteSpace=wrap;html=1;strokeColor=#95a5a6;fillColor=#ffffff;strokeWidth=1;"
    style_vnet = "rounded=0;dashed=1;whiteSpace=wrap;html=1;strokeColor=#5d6d7e;fillColor=#f4f6f7;"
    style_subnet = "rounded=0;whiteSpace=wrap;html=1;strokeColor=#aab7b8;fillColor=#ecf0f1;"
    style_count = "rounded=1;whiteSpace=wrap;html=1;strokeColor=#85929e;fillColor=#f2f4f4;fontSize=9;"
    style_storage = "rounded=0;whiteSpace=wrap;html=1;strokeColor=#7f8c8d;fillColor=#fcfcfc;"
    style_icon_base = (
        "shape=image;verticalLabelPosition=bottom;verticalAlign=top;align=center;"
        "imageAspect=0;aspect=fixed;whiteSpace=wrap;html=1;labelBackgroundColor=none;"
        "fontSize=10;strokeColor=none;fillColor=none;"
    )

    # Icons are resolved from the local Azure-Icons/ folder and embedded as
    # base64 data URIs by get_icon_uri(). The remote URL is kept as a
    # documented fallback that only kicks in if the local file is missing.
    AZ = "https://www.azureicons.com/static/images/icons"
    TW = "https://raw.githubusercontent.com/twitter/twemoji/master/assets/svg"

    icon_vm = get_icon_uri("Compute", "Virtual-Machine", f"{AZ}/Compute/svg/Virtual-Machine.svg")
    icon_vnet = get_icon_uri("Networking", "Virtual-Networks", f"{AZ}/Networking/svg/Virtual-Networks.svg")
    icon_subnet = get_icon_uri("Networking", "Subnet", f"{AZ}/Networking/svg/Subnet.svg")
    icon_storage = get_icon_uri("Storage", "Storage-Accounts", f"{AZ}/Storage/svg/Storage-Accounts.svg")
    icon_database = get_icon_uri("Databases", "Azure-SQL", f"{AZ}/Databases/svg/Azure-SQL.svg")
    icon_sql_vm = get_icon_uri("Databases", "Azure-SQL-VM", f"{AZ}/Databases/svg/Azure-SQL-VM.svg")
    icon_mysql = get_icon_uri("Databases", "Azure-Database-MySQL-Server", f"{AZ}/Databases/svg/Azure-Database-MySQL-Server.svg")
    icon_postgres = get_icon_uri("Databases", "Azure-Database-PostgreSQL-Server", f"{AZ}/Databases/svg/Azure-Database-PostgreSQL-Server.svg")
    icon_cosmos = get_icon_uri("Databases", "Azure-Cosmos-DB", f"{AZ}/Databases/svg/Azure-Cosmos-DB.svg")
    icon_dns = get_icon_uri("Networking", "DNS-Zones", f"{AZ}/Networking/svg/DNS-Zones.svg")
    icon_nsg = get_icon_uri("Networking", "Network-Security-Groups", f"{AZ}/Networking/svg/Network-Security-Groups.svg")
    # AKS lives under Compute/ locally (override map handles the lookup).
    icon_aks = get_icon_uri("Containers", "Kubernetes-Services", f"{AZ}/Containers/svg/Kubernetes-Services.svg")
    icon_lb = get_icon_uri("Networking", "Load-Balancers", f"{AZ}/Networking/svg/Load-Balancers.svg")
    icon_appgw = get_icon_uri("Networking", "Application-Gateways", f"{AZ}/Networking/svg/Application-Gateways.svg")
    icon_pe = get_icon_uri("Networking", "Private-Link", f"{AZ}/Networking/svg/Private-Link.svg")
    icon_private_link = get_icon_uri("Networking", "Private-Link-Service", f"{AZ}/Networking/svg/Private-Link-Service.svg")
    icon_vmss = get_icon_uri("Compute", "VM-Scale-Sets", f"{AZ}/Compute/svg/VM-Scale-Sets.svg")
    # Previously Twemoji glyphs — now resolved to Azure-native icons via the
    # override map (Identity/Users, Networking/Public-IP-Addresses, Other/Peerings).
    icon_peering = get_icon_uri("Twemoji", "Peering", f"{TW}/1f517.svg")
    icon_user = get_icon_uri("Twemoji", "User", f"{TW}/1f464.svg")
    icon_internet = get_icon_uri("Twemoji", "Internet", f"{TW}/1f310.svg")

    def short_name(value: str, limit: int = 16) -> str:
        # Keep full resource names exactly as discovered (no truncation).
        return str(value or "").strip()

    def add_icon_vertex(value: str, x: int, y: int, w: int, h: int, image_url: str, parent: str = "1") -> str:
        style = f"{style_icon_base}image={image_url};"
        return add_vertex(value, x, y, w, h, style, parent=parent)

    # Left ingress.
    user_id = add_icon_vertex("User", 40, ingress_y - 10, 52, 58, icon_user)
    internet_id = add_icon_vertex("Internet", 130, ingress_y - 10, 52, 58, icon_internet)
    dns_label = "Azure DNS" if _has_azure_dns(summary) else "Global DNS"
    dns_id = add_icon_vertex(dns_label, 270, ingress_y - 18, 70, 72, icon_dns)
    add_edge(user_id, internet_id)
    add_edge(internet_id, dns_id)

    # Outer environment and inner cloud containers.
    env_id = add_vertex("Cloud Environment", env_x, env_y, env_w, env_h, style_outer)
    core_id = add_vertex("Network Core", core_x, core_y, core_w, core_h, style_cloud)
    storage_box_id = add_vertex("Storage / Data", storage_x, storage_y, storage_w, storage_h, style_storage)
    add_edge(dns_id, env_id)

    # VNet placement with explicit subnet sub-boxes.
    # Each subnet shows up to 4 VM icons from that subnet.
    base_x = core_x + 30
    base_y = core_y + 40
    subnet_header_h = 30

    row_y = base_y
    i = 0
    vnet_pos: Dict[str, Dict[str, float]] = {}
    pending_subnet_icons: List[dict] = []
    pending_vnet_icons: List[dict] = []
    pending_aks_lb_edges: List[dict] = []
    subnet_box_ids: Dict[str, str] = {}          # "vnet::subnet" → draw.io cell id
    subnet_box_positions: Dict[str, Dict] = {}   # "vnet::subnet" → {x,y,w,h}
    while i < len(vnet_layout):
        row_items = vnet_layout[i:i + max_per_row]
        row_max_h = max(item["height"] for item in row_items)
        for col, item in enumerate(row_items):
            vname = item["name"]
            subnet_entries = item["subnets"]
            per_subnet_h_list = item.get("per_subnet_h") or [subnet_h] * len(subnet_entries)
            vnet_h = item["height"]
            vx = base_x + col * (vnet_w + gap_x)
            vy = row_y

            vnet_id = add_vertex("", vx, vy, vnet_w, vnet_h, style_vnet)
            vnet_pos[vname.lower()] = {
                "id": vnet_id,
                "cx": vx + (vnet_w / 2),
                "cy": vy + (vnet_h / 2),
            }

            # Queue VNet icon (render in final pass to keep in front).
            vnet_icon_w = 48
            vnet_icon_h = 51
            pending_vnet_icons.append({
                "name": short_name(vname, 20),
                "x": int(vx + vnet_w - (vnet_icon_w / 2)),
                "y": int(vy + vnet_h - (vnet_icon_h / 2)),
                "w": vnet_icon_w,
                "h": vnet_icon_h,
            })

            # Group VMs by subnet within this VNet.
            subnet_to_vms: Dict[str, List[dict]] = defaultdict(list)
            for vm in vnet_to_vms.get(vname.lower(), []):
                sname = str(vm.get("subnet") or "").strip() or "default"
                subnet_to_vms[sname.lower()].append(vm)

            # Map to track subnet cell IDs for cross-subnet arrows (e.g., AppGW -> Business Subnet)
            vnet_subnet_boxes: Dict[str, str] = {}
            vnet_ag_nids: List[Tuple[str, str]] = []
            for sidx, subnet_info in enumerate(subnet_entries):
                sname = str(subnet_info.get("name") or "default")
                nsg_id = str(subnet_info.get("nsg") or "")
                this_subnet_h = per_subnet_h_list[sidx] if sidx < len(per_subnet_h_list) else subnet_h
                sx = vx + 16
                sy = vy + 58 + sum(per_subnet_h_list[:sidx])
                sw = vnet_w - 32
                sh = this_subnet_h - 10
                subnet_cell_id = add_vertex("", sx, sy, sw, sh, style_subnet)
                _skey = f"{vname.lower()}::{sname.lower()}"
                vnet_subnet_boxes[_skey] = subnet_cell_id
                subnet_box_ids[_skey] = subnet_cell_id
                subnet_box_positions[_skey] = {"x": sx, "y": sy, "w": sw, "h": sh}
                # Subnet icon centered on the bottom-left corner of subnet box.
                subnet_icon_w = 42
                subnet_icon_h = 46
                icon_x = int(sx - (subnet_icon_w / 2))
                icon_y = int(sy + sh - (subnet_icon_h / 2))
                # Render subnet icons in a final pass so they stay in front.
                pending_subnet_icons.append({
                    "name": short_name(sname, 18),
                    "x": icon_x,
                    "y": icon_y,
                    "w": subnet_icon_w,
                    "h": subnet_icon_h,
                })

                subnet_vms = subnet_to_vms.get(sname.lower(), [])
                # Prioritize VMSS in the render order: each VMSS represents an
                # entire scale set (one logical entity), so it MUST appear inside
                # its subnet even when the subnet has many individual VMs above
                # the per-subnet cap. Without this, VMSS were getting pushed past
                # slot 4 and silently dropped — making them look like they
                # rendered "outside" the subnet box.
                vmss_items = [v for v in subnet_vms if v.get("is_vmss")]
                vm_items = [v for v in subnet_vms if not v.get("is_vmss")]
                remaining_slots = max(0, vm_max_per_subnet - len(vmss_items))
                show_vms = vmss_items + vm_items[:remaining_slots]
                # Centered 2x2 VM grid with explicit text labels under each icon.
                cell_w = 160
                cell_h = 68
                cell_gap_x = 26
                grid_w = (2 * cell_w) + cell_gap_x
                grid_left = int(sx + max(10, (sw - grid_w) / 2))
                grid_top = sy + 28
                vm_icon_w = 44
                vm_icon_h = 50
                subnet_vm_ids: List[str] = []   # collect IDs so LB can arrow to them
                for vm_idx, vm in enumerate(show_vms):
                    vm_name = str(vm.get("name") or f"vm-{vm_idx+1}")
                    vm_row = vm_idx // 2
                    vm_col = vm_idx % 2
                    cell_x = grid_left + (vm_col * (cell_w + cell_gap_x))
                    cell_y = grid_top + (vm_row * cell_h)
                    icon_vm_x = int(cell_x + (cell_w - vm_icon_w) / 2)
                    icon_vm_y = cell_y
                    # Use SQL VM icon when marked, VMSS icon when this row represents a VM Scale Set
                    if vm.get("is_sql_vm"):
                        use_icon = icon_sql_vm
                    else:
                        use_icon = icon_vmss if vm.get("is_vmss") or (isinstance(vm.get("type"), str) and "scaleset" in vm.get("type", "").lower()) else icon_vm
                    vm_nid = add_icon_vertex(short_name(vm_name, 14), icon_vm_x, icon_vm_y, vm_icon_w, vm_icon_h, use_icon)
                    subnet_vm_ids.append(vm_nid)
                    try:
                        compute_node_ids[vm_name.lower()] = vm_nid
                    except Exception:
                        pass

                # ── AKS clusters: render inside the subnet, or standalone LBs otherwise ──
                subnet_key = f"{vname.lower()}::{sname.lower()}"
                aks_items = subnet_to_aks.get(subnet_key, [])
                lb_items = subnet_to_lbs.get(subnet_key, [])

                if aks_items:
                    # Build AKS tile inside this subnet (below the VM grid)
                    # RG → LB lookup for pairing
                    rg_to_lbs_local: Dict[str, List[dict]] = defaultdict(list)
                    for lb in load_balancers:
                        lb_rg = str(lb.get("resource_group") or "").strip().lower()
                        if lb_rg:
                            rg_to_lbs_local[lb_rg].append(lb)

                    n_aks = len(aks_items)
                    cell_pw = min(130, max(90, (sw - 24) // n_aks - 12))  # fit inside subnet width
                    cell_ph = 90
                    cell_gap = 10
                    header_h = 26
                    tile_w = (n_aks * cell_pw) + ((n_aks - 1) * cell_gap) + 24
                    tile_h = cell_ph + header_h + 14
                    tile_x = sx + max(0, (sw - tile_w) // 2)  # horizontally centred
                    tile_y = sy + (subnet_h - 10)              # just below the normal VM area

                    # Outer AKS tile box
                    add_vertex(
                        "Azure Kubernetes Services",
                        tile_x, tile_y, tile_w, tile_h,
                        "rounded=1;whiteSpace=wrap;html=1;"
                        "strokeColor=#326ce5;fillColor=#e8f0fe;"
                        "strokeWidth=2;verticalAlign=top;fontSize=11;fontStyle=1;",
                    )

                    for a_idx, aks_item in enumerate(aks_items):
                        aks_name = str(aks_item.get("name") or f"aks-{a_idx+1}").strip()
                        aks_rg = str(aks_item.get("resource_group") or "").strip().lower()

                        cell_x = tile_x + 12 + a_idx * (cell_pw + cell_gap)
                        cell_y = tile_y + header_h + 4
                        # Per-cluster dashed cell
                        add_vertex(
                            "",
                            cell_x, cell_y, cell_pw, cell_ph,
                            "rounded=1;whiteSpace=wrap;html=1;"
                            "strokeColor=#5d8fd4;fillColor=#dce8fc;"
                            "strokeWidth=1;dashed=1;",
                        )
                        paired_lbs = rg_to_lbs_local.get(aks_rg, [])
                        lb_label = str(paired_lbs[0].get("name") or "Managed LB").strip() if paired_lbs else "Managed LB"
                        icon_lb_x = cell_x + 4
                        icon_aks_x = cell_x + cell_pw - 48
                        icon_y_inner = cell_y + 14
                        lb_id_k8s = add_icon_vertex(lb_label, icon_lb_x, icon_y_inner, 38, 44, icon_lb)
                        aks_id_k8s = add_icon_vertex(aks_name, icon_aks_x, icon_y_inner, 38, 44, icon_aks)
                        pending_aks_lb_edges.append({"lb": lb_id_k8s, "aks": aks_id_k8s})

                elif lb_items:
                    # Non-AKS subnet: LB at bottom-right with arrows pointing to VMs.
                    lb_edge_style = (
                        "endArrow=block;endFill=1;strokeWidth=1;"
                        "strokeColor=#5d6d7e;dashed=0;"
                    )
                    for lb_idx, lb_item in enumerate(lb_items[:3]):
                        lb_name = str(lb_item.get("name") or "load-balancer").strip()
                        lb_icon_x = int(sx + sw - 52 - lb_idx * 46)
                        lb_icon_y = int(sy + sh - 52)
                        lb_nid = add_icon_vertex(lb_name, lb_icon_x, lb_icon_y, 36, 42, icon_lb)
                        # Draw arrow from LB to each VM shown in this subnet
                        for vm_nid in subnet_vm_ids:
                            add_edge(lb_nid, vm_nid, style=lb_edge_style)

                # Application Gateways ─ rendered at bottom-left corner of the hosting subnet.
                appgw_items_in_subnet = subnet_to_appgw.get(subnet_key, [])
                for ag_idx, ag in enumerate(appgw_items_in_subnet[:2]):
                    ag_name = str(ag.get("name") or "app-gateway").strip()
                    ag_icon_x = int(sx + 10 + ag_idx * 52)
                    ag_icon_y = int(sy + sh - 52)
                    ag_nid = add_icon_vertex(ag_name, ag_icon_x, ag_icon_y, 40, 46, icon_appgw)
                    vnet_ag_nids.append((ag_nid, subnet_key))

                if subnet_vms or nsg_id:
                    # NSG icon only (no label), smaller and top-right.
                    add_icon_vertex("", sx + sw - 30, sy + 6, 24, 26, icon_nsg)

                shown = len(show_vms)
                total = len(subnet_vms)
                more = total - shown
                count_text = f"Subnet VMs: {total}"
                if more > 0:
                    count_text += f" (showing {shown}, +{more})"
                count_w = int(sw * 0.5)
                count_x = int(sx + ((sw - count_w) / 2))
                add_vertex(count_text, count_x, sy + 6, count_w, 20, style_count)

            vnet_total_vms = len(vnet_to_vms.get(vname.lower(), []))
            add_vertex(
                f"Total VMs in VNet: {vnet_total_vms}",
                vx + 16,
                vy + vnet_h - 26,
                vnet_w - 32,
                20,
                style_count
            )

            # Final pass for AppGW edges within this VNet (drawn after all subnet boxes exist)
            # Per USER FEEDBACK: One clean arrow to the target boxes instead of individual VMs.
            for ag_nid, source_skey in vnet_ag_nids:
                for target_skey, target_sid in vnet_subnet_boxes.items():
                    if target_skey != source_skey:
                        # Only point to subnets that actually have VMs? usually yes.
                        if subnet_to_vms.get(target_skey.split("::")[-1].lower()):
                             add_edge(
                                ag_nid, target_sid,
                                style="endArrow=block;endFill=1;strokeWidth=2.5;strokeColor=#e65100;rounded=1;edgeStyle=orthogonalEdgeStyle;"
                             )

        row_y += row_max_h + gap_y
        i += max_per_row

    # (AKS clusters are now drawn inside their subnet — no floating bottom section needed.)

    # Data/Storage items — track node IDs for PE/VNet-integration edge drawing.
    storage_node_ids: Dict[str, str] = {}   # resource_name.lower() → cell_id
    storage_rg_ids: Dict[str, List[str]] = defaultdict(list)  # rg.lower() → [cell_id,...]
    storage_node_positions: Dict[str, Dict[str, int]] = {}
    # Map compute VM names to their draw.io node ids so DB/PE edges can target VMs.
    compute_node_ids: Dict[str, str] = {}

    for i, s in enumerate((storage_accounts or [])[:16], start=1):
        name = str(s.get("name") or f"storage-{i}")
        rg = str(s.get("resource_group") or "").strip().lower()
        col = (i - 1) % 2
        row = (i - 1) // 2
        x = storage_x + 18 + col * 205
        y = storage_y + 26 + row * 92
        nid = add_icon_vertex(name, x + 58, y, 72, 76, icon_storage)
        storage_node_ids[name.lower()] = nid
        storage_node_positions[name.lower()] = {"x": x + 58, "y": y}
        if rg:
            storage_rg_ids[rg].append(nid)

    db_start_row = ((min(len(storage_accounts), 16) + 1) // 2) + 1
    for i, d in enumerate((databases or [])[:8], start=1):
        name = str(d.get("name") or d.get("engine") or f"db-{i}")
        rg = str(d.get("resource_group") or "").strip().lower()
        col = (i - 1) % 2
        row = db_start_row + ((i - 1) // 2)
        x = storage_x + 18 + col * 205
        y = storage_y + 26 + row * 92
        # Choose icon based on known engine/type hints
        engine_hint = str(d.get("engine") or d.get("type") or d.get("kind") or "").lower()
        if "mysql" in engine_hint or "mari" in engine_hint:
            db_icon = icon_mysql
        elif "post" in engine_hint or "postgres" in engine_hint:
            db_icon = icon_postgres
        elif "cosmos" in engine_hint:
            db_icon = icon_cosmos
        elif "sql" in engine_hint:
            db_icon = icon_database
        else:
            db_icon = icon_database
        nid = add_icon_vertex(name, x + 58, y, 72, 76, db_icon)
        storage_node_ids[name.lower()] = nid
        storage_node_positions[name.lower()] = {"x": x + 58, "y": y}
        if rg:
            storage_rg_ids[rg].append(nid)

    # Build mapping from storage/db node → compute VM node for SQL VMs (so arrows can point to the VM)
    storage_to_vm: Dict[str, str] = {}
    drawn_db_vm_edges: set = set()
    try:
        # First, attempt to map databases that explicitly reference a VM resource id in their raw_row
        for db in (summary.get("databases", []) or []):
            db_name = str(db.get("name") or "").strip().lower()
            raw = db.get("raw_row") or {}
            found = False
            # common column keys that may point to a VM resource id
            keys = [
                "virtualMachineResourceId", "virtual_machine_resource_id", "vmResourceId", "vm_resource_id",
                "serverResourceId", "server_resource_id", "resourceId", "resource_id", "vm",
            ]
            for k in keys:
                if not raw:
                    break
                val = raw.get(k) if isinstance(raw, dict) else None
                if not val:
                    # try case-insensitive key
                    for rk, rv in (raw.items() if isinstance(raw, dict) else []):
                        if rk and rk.lower() == k.lower():
                            val = rv
                            break
                if val:
                    try:
                        valstr = str(val).strip().lower()
                        # try to match against VM resource_id in compute summary
                        for vm in summary.get("compute", []):
                            rid = str(vm.get("resource_id") or "").strip().lower()
                            if not rid:
                                continue
                            if rid == valstr or rid.endswith(valstr.split("/")[-1]) or valstr.endswith(rid.split("/")[-1]):
                                vm_nid = compute_node_ids.get(str(vm.get("name") or "").strip().lower())
                                if vm_nid:
                                    storage_to_vm[db_name] = vm_nid
                                    found = True
                                    break
                        if found:
                            break
                    except Exception:
                        continue

        # Fallback fuzzy name matching between storage/db node names and VM names
        if not storage_to_vm:
            for storage_name in list(storage_node_ids.keys()):
                s_norm = re.sub(r"[^a-z0-9]", "", storage_name.lower())
                for vm in summary.get("compute", []):
                    if not vm.get("is_sql_vm"):
                        continue
                    vm_name = str(vm.get("name") or "").strip().lower()
                    v_norm = re.sub(r"[^a-z0-9]", "", vm_name)
                    if not vm_name:
                        continue
                    # fuzzy match: normalized name contained in other
                    if v_norm and (v_norm in s_norm or s_norm in v_norm or vm_name in storage_name or storage_name in vm_name):
                        vm_nid = compute_node_ids.get(vm_name)
                        if vm_nid:
                            storage_to_vm[storage_name] = vm_nid
                            break
    except Exception:
        storage_to_vm = {}

    # ── Private Endpoint / VNet Integration edges ──
    # Each PE is rendered as a small icon INSIDE its target subnet.
    # A dashed edge connects the matched storage/DB resource to the PE node.
    edge_style_pe = (
        "dashed=1;endArrow=open;endFill=0;strokeColor=#6c8ebf;strokeWidth=1.5;"
    )
    drawn_pe_edges: set = set()

    for pe in private_endpoints:
        pe_vnet_hint = str(pe.get("vnet") or "").strip()
        pe_subnet = str(pe.get("subnet") or "").strip()
        # If subnet contains a full ARM resource id, extract vnet and subnet names
        if pe_subnet and pe_subnet.lower().startswith("/subscriptions/") or ("/virtualnetworks/" in pe_subnet.lower() and "/subnets/" in pe_subnet.lower()):
            try:
                parts = pe_subnet.split("/")
                lw = [p.lower() for p in parts]
                if "virtualnetworks" in lw and "subnets" in lw:
                    vi = lw.index("virtualnetworks")
                    si = lw.index("subnets")
                    if vi + 1 < len(parts):
                        pe_vnet_hint = parts[vi + 1]
                    if si + 1 < len(parts):
                        pe_subnet = parts[si + 1]
            except Exception:
                pass
        pe_subnet = pe_subnet.strip().lower()
        pe_rg = str(pe.get("resource_group") or "").strip().lower()
        pe_name = str(pe.get("name") or "private-endpoint").strip()
        pe_name_lower = pe_name.lower()

        if not pe_vnet_hint and not pe_subnet:
            continue

        pe_vnet_name = _match_vnet_name(pe_vnet_hint)

        # Resolve target subnet box and its pixel position.
        target_subnet_cell = ""
        target_subnet_pos: Dict = {}
        if pe_vnet_name and pe_subnet:
            _k = f"{pe_vnet_name.lower()}::{pe_subnet}"
            target_subnet_cell = subnet_box_ids.get(_k, "")
            target_subnet_pos = subnet_box_positions.get(_k, {})
        if not target_subnet_cell and pe_vnet_name:
            for k, cid in subnet_box_ids.items():
                vpart, spart = (k.split("::", 1) + [""])[:2]
                if vpart == pe_vnet_name.lower() and (
                    not pe_subnet or pe_subnet in spart or spart in pe_subnet
                ):
                    target_subnet_cell = cid
                    target_subnet_pos = subnet_box_positions.get(k, {})
                    break

        if not target_subnet_cell:
            continue

        # Do not render a separate PE icon/name inside the subnet box (we render
        # a Private Link icon on the connecting edge instead). Keep pe_node_id
        # empty so edges target the subnet cell directly.
        pe_node_id = ""

        # ─ Match linked storage / DB resources (4-tier strategy) ─
        matched_nodes: List[str] = []
        
        # Tier 0: Linked Resource ARM ID match (Precise) — support multiple linked resources
        linked_candidates: List[str] = []
        if pe.get("linked_resource"):
            linked_candidates.append(str(pe.get("linked_resource")))
        for lr in (pe.get("linked_resources") or []):
            try:
                if isinstance(lr, str):
                    linked_candidates.append(lr)
                elif isinstance(lr, dict):
                    # Try common fields in dict
                    for k in ("id", "resourceId", "resource_id", "link"):
                        if lr.get(k):
                            linked_candidates.append(str(lr.get(k)))
                            break
            except Exception:
                continue

        for linked_res in linked_candidates:
            linked_res = str(linked_res or "").strip().lower()
            if not linked_res:
                continue
            target_res_name = linked_res.split("/")[-1]
            # Exact name match in storage/db nodes
            nid = storage_node_ids.get(target_res_name)
            if nid and nid not in matched_nodes:
                matched_nodes.append(nid)
                continue
            # Fuzzy: check if the extracted name is substring of any known resource name
            for res_name, nid2 in storage_node_ids.items():
                if target_res_name and (target_res_name in res_name or res_name in target_res_name):
                    if nid2 not in matched_nodes:
                        matched_nodes.append(nid2)
                        break

        # Tier 1: same RG as PE
        if not matched_nodes:
            matched_nodes = list(storage_rg_ids.get(pe_rg, []))
        
        # Tier 2: PE name contains resource name (or vice-versa)
        if not matched_nodes:
            clean_pe = re.sub(r"[^a-z0-9]", "", pe_name_lower)
            for res_name, nid in storage_node_ids.items():
                clean_res = re.sub(r"[^a-z0-9]", "", res_name)
                if clean_res and clean_pe and (
                    clean_res in clean_pe or clean_pe in clean_res
                ):
                    matched_nodes.append(nid)

        # Tier 3: SQL keyword match ─ PE name contains 'sql'; connect all SQL-type nodes
        if not matched_nodes and any(
            kw in pe_name_lower for kw in ("sql", "sqlserver", "database", "cosmos")
        ):
            for res_name, nid in storage_node_ids.items():
                if any(kw in res_name for kw in ("sql", "db", "database", "cosmos")):
                    matched_nodes.append(nid)

        target_cell = pe_node_id if pe_node_id else target_subnet_cell
        for node_id in matched_nodes:
            edge_key = (node_id, target_cell)
            if edge_key in drawn_pe_edges:
                continue
            drawn_pe_edges.add(edge_key)
            # Try to render a private-link icon at the midpoint of the edge.
            # Determine source coordinates from stored storage positions; determine
            # target coordinates from PE node position or subnet box center.
            src_pos = None
            # reverse lookup name for this node_id
            src_name = None
            for rn, nid in storage_node_ids.items():
                if nid == node_id:
                    src_name = rn
                    break
            # If this storage node maps to a VM (SQL VM), prefer the VM as the source target for the PE arrow.
            if src_name and storage_to_vm.get(src_name):
                vm_nid = storage_to_vm.get(src_name)
                # draw a small dashed connector from DB -> VM for clarity (once)
                try:
                    if (node_id, vm_nid) not in drawn_db_vm_edges:
                        add_edge(node_id, vm_nid, style="endArrow=block;endFill=1;strokeWidth=1;strokeColor=#888;dashed=1;")
                        drawn_db_vm_edges.add((node_id, vm_nid))
                except Exception:
                    pass
                # use VM node as the effective source for PE edges
                node_id = vm_nid
            if src_name and storage_node_positions.get(src_name):
                src_pos = storage_node_positions.get(src_name)

            tgt_x = None
            tgt_y = None
            if pe_node_id:
                # use the PE icon coordinates we computed earlier
                tgt_x = pe_icon_x
                tgt_y = pe_icon_y
            elif target_subnet_pos:
                sp = target_subnet_pos
                tgt_x = int(sp["x"] + sp["w"] / 2)
                tgt_y = int(sp["y"] + sp["h"] / 2)

            if src_pos and tgt_x is not None:
                mid_x = int((src_pos["x"] + tgt_x) / 2) - 12
                mid_y = int((src_pos["y"] + tgt_y) / 2) - 12
                pe_icon_mid = add_icon_vertex("Private Endpoint", mid_x, mid_y, 28, 28, icon_private_link)
                add_edge(node_id, pe_icon_mid, style=edge_style_pe)
                add_edge(pe_icon_mid, target_cell, style=edge_style_pe)
            else:
                # fallback: direct edge if positions unavailable
                add_edge(node_id, target_cell, style=edge_style_pe)

    # Peering links/icons between VNets if peering metadata exists.
    used_edges = set()
    any_peering_text = any(_is_present(v.get("peerings")) for v in networks)
    rendered_peering = False
    for v in networks:
        src_name = str(v.get("name") or "").strip().lower()
        if src_name not in vnet_pos:
            continue
        src = vnet_pos[src_name]
        targets = _parse_peering_targets(v.get("peerings"))
        for t in targets:
            target_key = str(t or "").strip().lower()
            if target_key not in vnet_pos:
                # tolerant contains matching
                for k in vnet_pos.keys():
                    if target_key and (target_key in k or k in target_key):
                        target_key = k
                        break
            if target_key not in vnet_pos or target_key == src_name:
                continue
            edge_key = tuple(sorted([src_name, target_key]))
            if edge_key in used_edges:
                continue
            used_edges.add(edge_key)
            dst = vnet_pos[target_key]
            mid_x = int((src["cx"] + dst["cx"]) / 2) - 28
            mid_y = int((src["cy"] + dst["cy"]) / 2) - 30
            # Peering icon only, rendered late to stay in front.
            p_id = add_icon_vertex("peering", mid_x, mid_y, 44, 50, icon_peering)
            add_edge(str(src["id"]), p_id, style="dashed=1;endArrow=none;strokeColor=#5d6d7e;")
            add_edge(p_id, str(dst["id"]), style="dashed=1;endArrow=none;strokeColor=#5d6d7e;")
            rendered_peering = True

    # Fallback peering icon when peering text exists but name matching failed.
    if any_peering_text and not rendered_peering and len(vnet_layout) >= 2:
        a = vnet_layout[0]["name"].lower()
        b = vnet_layout[1]["name"].lower()
        if a in vnet_pos and b in vnet_pos:
            src = vnet_pos[a]
            dst = vnet_pos[b]
            mid_x = int((src["cx"] + dst["cx"]) / 2) - 22
            mid_y = int((src["cy"] + dst["cy"]) / 2) - 25
            p_id = add_icon_vertex("peering", mid_x, mid_y, 44, 50, icon_peering)
            add_edge(str(src["id"]), p_id, style="dashed=1;endArrow=none;strokeColor=#5d6d7e;")
            add_edge(p_id, str(dst["id"]), style="dashed=1;endArrow=none;strokeColor=#5d6d7e;")

    # Final pass: subnet icons on top/front.
    for sn in pending_subnet_icons:
        add_icon_vertex(sn["name"], sn["x"], sn["y"], sn["w"], sn["h"], icon_subnet)
    # Final pass: vnet icons on very top/front.
    for vn in pending_vnet_icons:
        add_icon_vertex(vn["name"], vn["x"], vn["y"], vn["w"], vn["h"], icon_vnet)
    # Final pass: draw LB -> AKS edges after all nodes exist.
    for edge in pending_aks_lb_edges:
        add_edge(edge["lb"], edge["aks"], style="endArrow=block;endFill=1;strokeWidth=1;strokeColor=#5d6d7e;")

    # Draw a fallback VMSS panel only for VM Scale Sets that were NOT placed into any VNet/subnet.
    # IMPORTANT: previous positioning was `storage_x + storage_w + 30` which placed the panel
    # OUTSIDE the environment box (env_w only extends to storage_x + storage_w + 40, but the
    # panel itself is 260 wide → ~250px hanging off the right edge). We now stack the panel
    # BELOW the storage box, aligned with it, so it stays inside the visible diagram area.
    unplaced_vmss = [v for v in vm_scale_sets if not str(v.get("vnet") or "").strip()]
    if unplaced_vmss:
        panel_x = storage_x
        panel_w = storage_w
        panel_y = storage_y + storage_h + 20
        panel_h = max(120, min(600, 28 + len(unplaced_vmss) * 28))
        panel_id = add_vertex(f"VM Scale Sets ({len(unplaced_vmss)})", panel_x, panel_y, panel_w, panel_h, style_outer)
        # Add simple icon entries vertically for only the unplaced VMSS
        icon_x = panel_x + 12
        icon_y = panel_y + 10
        for idx, vmss in enumerate(unplaced_vmss[:16], start=1):
            name = str(vmss.get("name") or f"vmss-{idx}")
            # Icon (may fail to display if remote host blocks requests)
            add_icon_vertex(name, icon_x, icon_y + (idx - 1) * 28, 20, 20, icon_vmss, parent=panel_id)
            # Text fallback to ensure the VMSS name is always visible next to the icon
            label_x = icon_x + 28
            label_y = icon_y + (idx - 1) * 28 - 2
            label_w = panel_w - (label_x - panel_x) - 12
            label_h = 20
            style_fallback_text = "rounded=0;whiteSpace=wrap;html=1;align=left;verticalAlign=middle;fontSize=11;strokeColor=none;fillColor=none;"
            add_vertex(name, label_x, label_y, label_w, label_h, style_fallback_text, parent=panel_id)

    # Ensure deterministic, compact XML output.
    return ET.tostring(mxfile, encoding="unicode")


def _call_openai(prompt: str) -> str:
    """Call Azure OpenAI to generate the diagram."""
    try:
        from openai import AzureOpenAI

        client = AzureOpenAI(
            api_key=AzureOpenAIConfig.API_KEY,
            api_version=AzureOpenAIConfig.API_VERSION,
            azure_endpoint=AzureOpenAIConfig.ENDPOINT,
        )

        response = client.chat.completions.create(
            model=AzureOpenAIConfig.DEPLOYMENT_NAME,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert cloud architect who creates clear, accurate Mermaid diagrams. Output ONLY valid Mermaid code, no markdown fences, no explanations.",
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,  # Lower for more consistent diagrams
            max_completion_tokens=min(AzureOpenAIConfig.MAX_TOKENS_PER_REQUEST, 16000), # Use config value, cap at 16k for output safety
        )

        return response.choices[0].message.content.strip()

    except Exception as e:
        print(f"❌ OpenAI call failed: {e}")
        # Return a fallback diagram
        return _generate_fallback_diagram(prompt)


def _clean_mermaid_output(code: str) -> str:
    """Remove markdown fences and clean up GPT output."""
    # Remove ```mermaid ... ``` wrapping
    code = re.sub(r'^```(?:mermaid)?\s*\n?', '', code, flags=re.MULTILINE)
    code = re.sub(r'\n?```\s*$', '', code, flags=re.MULTILINE)

    # Ensure it starts with a valid graph directive
    lines = code.strip().split('\n')
    if lines and not any(lines[0].strip().startswith(d) for d in ['graph', 'flowchart', 'sequenceDiagram', 'classDiagram', 'stateDiagram', 'gantt', 'pie']):
        # Prepend graph TD if missing
        code = 'graph TD\n' + code

    code = _normalize_mermaid_edges(code)
    return code.strip()


def _normalize_mermaid_edges(code: str) -> str:
    """
    Normalize common GPT edge-label mistakes into valid Mermaid flowchart syntax.
    Example fix:
      A -- "Peering" -- B   ->   A ---|Peering| B
      A --> "HTTPS" --> B   ->   A -->|HTTPS| B
    """
    node_id = r"[A-Za-z][A-Za-z0-9_]*"

    undirected_labeled = re.compile(
        rf"\b({node_id})\b\s*--\s*\"([^\n\"]+)\"\s*--\s*\b({node_id})\b"
    )
    directed_labeled = re.compile(
        rf"\b({node_id})\b\s*-->\s*\"([^\n\"]+)\"\s*-->\s*\b({node_id})\b"
    )

    normalized_lines = []
    for raw_line in code.splitlines():
        line = raw_line
        line = undirected_labeled.sub(r"\1 ---|\2| \3", line)
        line = directed_labeled.sub(r"\1 -->|\2| \3", line)
        normalized_lines.append(line)

    return "\n".join(normalized_lines)


def _generate_fallback_diagram(prompt: str) -> str:
    """Generate a simple fallback diagram when GPT is unavailable."""
    return """graph TD
    A["⚠️ AI Unavailable"] --> B["Architecture diagram could not be generated"]
    B --> C["Please check Azure OpenAI configuration"]
    style A fill:#FFEB9C,stroke:#F57C00,color:#E65100
    style C fill:#E3F2FD,stroke:#1565C0,color:#0D47A1"""


# ---------------------------------------------------------------------------
# CLI Test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python architecture_generator.py <excel_file> [diagram_type] [cloud_type]")
        print("  diagram_type: network (default), resource_groups, migration")
        print("  cloud_type: azure (default), aws, gcp")
        sys.exit(1)

    file_path = sys.argv[1]
    diagram_type = sys.argv[2] if len(sys.argv) > 2 else "network"
    cloud_type = sys.argv[3] if len(sys.argv) > 3 else None

    result = generate_architecture_diagram(file_path, diagram_type, cloud_type)

    print(f"\n{'='*60}")
    print(f"Cloud: {result['cloud']}")
    print(f"Type: {result['diagram_type']}")
    print(f"{'='*60}")
    print(result['diagram'])
    print(f"{'='*60}")
