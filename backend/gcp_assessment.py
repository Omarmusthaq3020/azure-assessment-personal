"""
GCP to Azure Migration Assessment Tool
========================================
Extracts GCP resources and maps them to Azure equivalents with SKU recommendations.
"""

import json
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from google.cloud import asset_v1
from google.cloud import compute_v1
from google.cloud import storage
from google.oauth2 import service_account
import concurrent.futures

try:
    from sku_mapper import SkuMapper
except ImportError:
    import sys
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from sku_mapper import SkuMapper

GLOBAL_SKU_MAPPER = SkuMapper()

# ============================================================
# Column definitions — what columns are GCP vs Azure
# ============================================================
GCP_COLUMNS = [
    "Name", "GCP Service", "Resource Type", "Machine Type / Class",
    "vCPUs", "Memory (GB)", "Storage Size (GB)", "Storage Class",
    "Location", "Project", "Status", "Labels"
]

AZURE_COLUMNS = [
    "Azure Service", "Azure VM SKU", "Azure Region"
]

# ============================================================
# Azure region mapping (GCP region → nearest Azure region)
# ============================================================
GCP_TO_AZURE_REGION = {
    # Asia Pacific
    "asia-south1": "Central India",
    "asia-south2": "South India",
    "asia-southeast1": "Southeast Asia",
    "asia-southeast2": "Southeast Asia",
    "asia-east1": "East Asia",
    "asia-east2": "East Asia",
    "asia-northeast1": "Japan East",
    "asia-northeast2": "Japan West",
    "asia-northeast3": "Korea Central",
    # Americas
    "us-central1": "Central US",
    "us-east1": "East US",
    "us-east4": "East US 2",
    "us-west1": "West US",
    "us-west2": "West US 2",
    "us-west3": "West US 3",
    "us-west4": "West US 3",
    "northamerica-northeast1": "Canada East",
    "northamerica-northeast2": "Canada Central",
    "southamerica-east1": "Brazil South",
    # Europe
    "europe-west1": "West Europe",
    "europe-west2": "UK South",
    "europe-west3": "Germany West Central",
    "europe-west4": "Netherlands",
    "europe-west6": "Switzerland North",
    "europe-central2": "Poland Central",
    "europe-north1": "Norway East",
    # Middle East / Africa
    "me-west1": "Israel Central",
    "me-central1": "Qatar Central",
    "africa-south1": "South Africa North",
    # Global
    "global": "Global",
}

# ============================================================
# GCP service → Azure service mapping
# ============================================================
GCP_AZURE_SERVICE_MAP = {
    "compute/Instance": "Virtual Machines",
    "compute/Disk": "Azure Managed Disks",
    "compute/Snapshot": "Azure Managed Disk Snapshots",
    "compute/Image": "Azure Compute Gallery",
    "compute/Network": "Azure Virtual Network",
    "compute/Subnetwork": "Azure Subnet",
    "compute/Firewall": "Azure Network Security Group",
    "compute/Address": "Azure Public IP Address",
    "compute/ForwardingRule": "Azure Load Balancer",
    "compute/TargetHttpProxy": "Azure Application Gateway",
    "compute/TargetHttpsProxy": "Azure Application Gateway",
    "compute/UrlMap": "Azure Traffic Manager",
    "compute/InstanceGroup": "Azure Virtual Machine Scale Sets",
    "compute/BackendService": "Azure Load Balancer",
    "compute/HealthCheck": "Azure Load Balancer Health Probes",
    "compute/Network": "Azure Virtual Network (VNet)",
    "compute/Subnetwork": "Azure Subnet",
    "compute/Firewall": "Azure Network Security Group (NSG) Rule",
    "storage/Bucket": "Azure Blob Storage",
    "sqladmin/Instance": "Azure SQL Database / Azure DB for MySQL/PostgreSQL",
    "container/Cluster": "Azure Kubernetes Service (AKS)",
    "bigquery/Dataset": "Azure Synapse Analytics",
    "pubsub/Topic": "Azure Service Bus",
    "pubsub/Subscription": "Azure Service Bus",
    "functions/CloudFunction": "Azure Functions",
    "run/Service": "Azure Container Apps",
    "redis/Instance": "Azure Cache for Redis",
    "filestore/Instance": "Azure Files",
    "bigtable/Instance": "Azure Cosmos DB (Cassandra API)",
    "spanner/Instance": "Azure Cosmos DB",
    "dataproc/Cluster": "Azure HDInsight / Azure Databricks",
    "composer/Environment": "Azure Data Factory (Managed Airflow)",
    "dns/ManagedZone": "Azure DNS",
    "iam/ServiceAccount": "Azure Managed Identity",
    "cloudkms/CryptoKey": "Azure Key Vault",
    "cloudkms/KeyRing": "Azure Key Vault",
    "logging/LogSink": "Azure Monitor / Azure Log Analytics",
    "monitoring/AlertPolicy": "Azure Monitor Alerts",
    "artifactregistry/Repository": "Azure Container Registry",
    "secretmanager/Secret": "Azure Key Vault Secrets",
    "serviceusage/Service": "N/A (API Enablement)",
    "cloudresourcemanager/Project": "N/A (GCP Project)",
}


def get_azure_region(gcp_location):
    """Map GCP location to nearest Azure region."""
    if not gcp_location:
        return "—"
    loc = gcp_location.lower()
    # Try exact match first
    if loc in GCP_TO_AZURE_REGION:
        return GCP_TO_AZURE_REGION[loc]
    # Try prefix match (e.g., asia-south1-c → asia-south1)
    for prefix, azure in GCP_TO_AZURE_REGION.items():
        if loc.startswith(prefix):
            return azure
    return gcp_location  # Return as-is if unknown


def get_azure_service(gcp_normalized_type):
    """Map GCP normalized type to Azure service."""
    return GCP_AZURE_SERVICE_MAP.get(gcp_normalized_type, "Assess Manually")


class GcpResourceExtractor:
    """Extract resources from GCP using Cloud Asset Inventory and service-specific APIs"""

    def __init__(self, project_ids, credentials_json=None, credentials_path=None, credentials=None):
        """
        Initialize GCP clients.

        Args:
            project_ids: Single GCP Project ID (str) or list of project IDs
            credentials_json: Service account JSON as dict
            credentials_path: Path to service account JSON file
            credentials: google.auth credentials object (OAuth token-based)
        """
        # Normalize to list and handle accidental Service Account emails
        if isinstance(project_ids, str):
            pids = [project_ids]
        else:
            pids = list(project_ids)
            
        self.project_ids = []
        for pid in pids:
            if "@" in str(pid) and ".iam.gserviceaccount.com" in str(pid):
                parts = pid.split("@")
                if len(parts) > 1:
                    domain_parts = parts[1].split(".")
                    if domain_parts:
                        pid = domain_parts[0]
            self.project_ids.append(pid)

        # Load credentials
        if credentials is not None:
            self.credentials = credentials
        elif credentials_json:
            self.credentials = service_account.Credentials.from_service_account_info(credentials_json)
        elif credentials_path:
            self.credentials = service_account.Credentials.from_service_account_file(credentials_path)
        else:
            self.credentials = None

        # Initialize clients (project-agnostic)
        self.asset_client = asset_v1.AssetServiceClient(credentials=self.credentials)

    def _get_compute_client(self):
        return compute_v1.InstancesClient(credentials=self.credentials)

    def _get_storage_client(self, project_id):
        return storage.Client(project=project_id, credentials=self.credentials)

    def extract_resources(self):
        """Extract ALL resources from all configured GCP projects."""
        all_resources = []
        for project_id in self.project_ids:
            print(f"🔍 Extracting GCP resources from project: {project_id}")
            resources = self._extract_project_resources(project_id)
            all_resources.extend(resources)
        return all_resources

    def _extract_project_resources(self, project_id):
        """Extract resources from a single GCP project."""
        all_resources = []
        parent = f"projects/{project_id}"

        try:
            response = self.asset_client.list_assets(
                request={
                    "parent": parent,
                    "content_type": asset_v1.ContentType.RESOURCE,
                }
            )
            for asset in response:
                resource = self._parse_asset(asset, project_id)
                if resource:
                    all_resources.append(resource)
            print(f"   ✅ Found {len(all_resources)} resources via Cloud Asset Inventory in {project_id}")

        except Exception as e:
            print(f"   ⚠️ Cloud Asset Inventory failed for {project_id}: {e}")
            if "Cloud Asset API" in str(e) and ("has not been used" in str(e) or "is disabled" in str(e)):
                print("   💡 Enable Cloud Asset API: https://console.cloud.google.com/apis/api/cloudasset.googleapis.com")
            print("   Falling back to service-specific APIs...")
            try:
                all_resources = self._scan_native_services(project_id)
            except Exception as scan_err:
                print(f"   ⚠️ Service-specific scan also failed: {scan_err}")
                all_resources = []

        # Enrich
        if all_resources:
            self._enrich_resources(all_resources, project_id)

        return all_resources

    def _parse_asset(self, asset, project_id):
        """Parse a Cloud Asset Inventory asset into normalized format."""
        try:
            asset_type = asset.asset_type
            name = asset.name

            resource_data = {}
            if asset.resource:
                resource_data = json.loads(type(asset.resource).to_json(asset.resource))

            service, resource_type = self._parse_asset_type(asset_type)

            # Skip internal/system resource types
            if service in ("serviceusage", "cloudresourcemanager"):
                return None

            resource_name = resource_data.get("name", name.split("/")[-1])
            location = self._extract_location(resource_data)

            # Extract key specs
            machine_type = resource_data.get("machineType", "")
            if machine_type and "/" in machine_type:
                machine_type = machine_type.split("/")[-1]

            vcpus, memory_gb = self._get_vm_specs(machine_type)
            storage_size_gb, storage_class = self._get_storage_specs(resource_data, service, resource_type)

            return {
                "_raw_type": asset_type,
                "_normalized_type": f"{service}/{resource_type}",
                "_project": project_id,
                "_machine_type": machine_type,
                "_resource_data": resource_data,
                # Display columns
                "Name": resource_name,
                "GCP Service": service.replace("googleapis.com", ""),
                "Resource Type": resource_type,
                "Machine Type / Class": machine_type or self._get_class_label(service, resource_type, resource_data),
                "vCPUs": vcpus,
                "Memory (GB)": memory_gb,
                "Storage Size (GB)": storage_size_gb,
                "Storage Class": storage_class,
                "Location": location,
                "Project": project_id,
                "Status": self._extract_status(resource_data),
                "Labels": self._extract_labels(resource_data),
            }

        except Exception as e:
            print(f"   ⚠️ Failed to parse asset: {e}")
            return None

    def _parse_asset_type(self, asset_type):
        parts = asset_type.split("/")
        if len(parts) == 2:
            service = parts[0].replace(".googleapis.com", "")
            return service, parts[1]
        return "unknown", asset_type

    def _extract_location(self, resource_data):
        for field in ["zone", "region", "location"]:
            if field in resource_data:
                loc = resource_data[field]
                if isinstance(loc, str) and "/" in loc:
                    return loc.split("/")[-1]
                return loc or "global"
        return "global"

    def _extract_status(self, resource_data):
        for field in ["status", "state", "lifecycleState"]:
            if field in resource_data:
                return str(resource_data[field])
        return "Active"

    def _extract_labels(self, resource_data):
        labels = resource_data.get("labels", {})
        if labels:
            return json.dumps(labels)
        return ""

    def _get_class_label(self, service, resource_type, resource_data):
        """Get a human-readable 'machine type' for non-VM resources."""
        if service == "storage":
            return resource_data.get("storageClass", "STANDARD")
        if service == "sqladmin":
            return resource_data.get("settings", {}).get("tier", "")
        if service == "redis":
            return resource_data.get("tier", "")
        if service == "container":
            return "GKE Cluster"
        return ""

    def _get_vm_specs(self, machine_type):
        """Return (vcpus, memory_gb) for a GCP machine type."""
        if not machine_type:
            return "", ""
        spec = GLOBAL_SKU_MAPPER._parse_gcp_machine_type(machine_type.lower())
        if spec:
            return spec["vcpu"], spec["ram"]
        return "", ""

    def _get_storage_specs(self, resource_data, service, resource_type):
        """Return (size_gb, storage_class)."""
        size_gb = ""
        storage_class = ""

        if service == "storage" and resource_type == "Bucket":
            # Bucket size is not available from Asset Inventory metadata; mark for enrichment
            storage_class = resource_data.get("storageClass", "STANDARD")

        elif service == "compute" and resource_type == "Disk":
            size_gb = resource_data.get("sizeGb", "")
            storage_class = resource_data.get("type", "").split("/")[-1] if resource_data.get("type") else ""

        elif service == "compute" and resource_type == "Instance":
            # Sum all disk sizes
            disks = resource_data.get("disks", [])
            total = 0
            for d in disks:
                try:
                    total += int(d.get("diskSizeGb", 0))
                except (ValueError, TypeError):
                    pass
            if total:
                size_gb = total

        elif service == "sqladmin":
            settings = resource_data.get("settings", {})
            size_gb = settings.get("dataDiskSizeGb", "")
            storage_class = settings.get("dataDiskType", "")

        elif service == "filestore":
            capacity = resource_data.get("capacityGb", "")
            if capacity:
                size_gb = capacity

        return size_gb, storage_class

    def _enrich_resources(self, resources, project_id):
        """Enrich resources (bucket sizes, Azure mapping) with parallel processing."""

        import sys
        
        # 1. Fetch bucket sizes sequentially but with a timeout per bucket to be safe
        # (We are already inside a thread pool if called from extract_resources, 
        # so avoiding nested complexity)
        bucket_sizes = {}
        try:
            print(f"   📊 Measuring storage sizes for project {project_id}...", flush=True)
            sc = self._get_storage_client(project_id)
            buckets_list = list(sc.list_buckets())
            
            def process_bucket(bucket_obj):
                try:
                    # List all blobs to get accurate total size. 
                    # Note: For buckets with millions of objects, this will be slow but accurate.
                    blobs = sc.list_blobs(bucket_obj.name)
                    total = 0
                    count = 0
                    for b in blobs:
                        total += (b.size or 0)
                        count += 1
                        if count % 10000 == 0:
                            print(f"      ⏳ Still measuring {bucket_obj.name}... ({count} objects so far)", flush=True)
                    
                    size = round(total / (1024 ** 3), 4)
                    return bucket_obj.name, size
                except Exception as e:
                    print(f"      ⚠️ Could not size {bucket_obj.name}: {e}", flush=True)
                    return bucket_obj.name, "—"

            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as bucket_executor:
                results = list(bucket_executor.map(process_bucket, buckets_list))
                for name, size in results:
                    bucket_sizes[name] = size
                    if size != "—":
                         print(f"      ✅ {name}: {size} GB", flush=True)

        except Exception as e:
            print(f"   ⚠️ Could not list buckets for sizing: {e}", flush=True)
            sys.stdout.flush()


        def enrich_item(resource):
            try:
                normalized_type = resource.get("_normalized_type", "")
                machine_type = resource.get("_machine_type", "")

                # Azure mappings
                resource["Azure Service"] = get_azure_service(normalized_type)
                resource["Azure VM SKU"] = (
                    GLOBAL_SKU_MAPPER.map_gcp_to_azure(machine_type, "VM")
                    if machine_type and "compute/Instance" in normalized_type
                    else "—"
                )
                resource["Azure Region"] = get_azure_region(resource.get("Location", ""))

                # Enrich bucket size
                if "storage/Bucket" in normalized_type:
                    name = resource.get("Name", "")
                    resource["Storage Size (GB)"] = bucket_sizes.get(name, "—")

            except Exception as e:
                print(f"   ⚠️ Enrichment failed for {resource.get('Name')}: {e}")
            return resource

        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
            enriched = list(executor.map(enrich_item, resources))

        resources.clear()
        resources.extend(enriched)
        print(f"   ✅ Enriched {len(resources)} resources from {project_id}")

    def _scan_native_services(self, project_id):
        """Fallback: scan using native GCP clients."""
        all_resources = []
        print("   📦 Scanning Compute Engine instances...")
        all_resources.extend(self._scan_compute_instances(project_id))
        print("   📦 Scanning Compute Engine networks (VPCs/Subnets)...")
        all_resources.extend(self._scan_compute_networks(project_id))
        print("   📦 Scanning Cloud Storage buckets...")
        all_resources.extend(self._scan_storage_buckets(project_id))
        return all_resources

    def _scan_compute_networks(self, project_id):
        """Scan all Compute Engine networks and subnets (especially useful for Host Projects)."""
        resources = []
        try:
            print(f"      🌐 Listing VPC networks in {project_id}...", flush=True)
            # Networks
            net_client = compute_v1.NetworksClient(credentials=self.credentials)
            for net in net_client.list(project=project_id):
                resources.append({
                    "_raw_type": "compute.googleapis.com/Network",
                    "_normalized_type": "compute/Network",
                    "_project": project_id,
                    "_machine_type": "",
                    "_resource_data": {},
                    "Name": net.name,
                    "GCP Service": "compute",
                    "Resource Type": "Network",
                    "Machine Type / Class": "VPC Network",
                    "vCPUs": "",
                    "Memory (GB)": "",
                    "Storage Size (GB)": "",
                    "Storage Class": "",
                    "Location": "global",
                    "Project": project_id,
                    "Status": "Active",
                    "Labels": "",
                    "Azure Service": get_azure_service("compute/Network"),
                    "Azure VM SKU": "—",
                    "Azure Region": "Global",
                })
        except Exception as e:
            print(f"      ⚠️ Network list failed: {e}", flush=True)

        try:
            print(f"      🌐 Listing subnetworks in {project_id}...", flush=True)
            # Subnets
            subnet_client = compute_v1.SubnetworksClient(credentials=self.credentials)
            # Use aggregated list to get subnets across all regions
            for region_name, subnets_scoped_list in subnet_client.aggregated_list(project=project_id):
                if subnets_scoped_list.subnetworks:
                    for subnet in subnets_scoped_list.subnetworks:
                        region = subnet.region.split("/")[-1] if subnet.region else "global"
                        resources.append({
                            "_raw_type": "compute.googleapis.com/Subnetwork",
                            "_normalized_type": "compute/Subnetwork",
                            "_project": project_id,
                            "_machine_type": "",
                            "_resource_data": {},
                            "Name": subnet.name,
                            "GCP Service": "compute",
                            "Resource Type": "Subnetwork",
                            "Machine Type / Class": "VPC Subnet",
                            "vCPUs": "",
                            "Memory (GB)": "",
                            "Storage Size (GB)": "",
                            "Storage Class": "",
                            "Location": region,
                            "Project": project_id,
                            "Status": "Active",
                            "Labels": "",
                            "Azure Service": get_azure_service("compute/Subnetwork"),
                            "Azure VM SKU": "—",
                            "Azure Region": get_azure_region(region),
                        })
        except Exception as e:
            msg = str(e)
            if "PermissionDenied" in msg or "403" in msg:
                print(f"   ❌ Access Denied: Missing network permissions (roles/compute.networkViewer).", flush=True)
            else:
                print(f"   ⚠️ Network scan failed: {e}", flush=True)
        return resources


    def _scan_compute_instances(self, project_id):
        """Scan all Compute Engine instances using aggregated list (much faster than zone-by-zone)."""
        resources = []
        try:
            cc = self._get_compute_client()
            # Use aggregated list to get instances across all zones in one go
            for zone_path, instances_scoped_list in cc.aggregated_list(project=project_id):
                if instances_scoped_list.instances:
                    zone_name = zone_path.split("/")[-1]
                    for instance in instances_scoped_list.instances:
                        machine_type = instance.machine_type.split("/")[-1]
                        vcpus, memory_gb = self._get_vm_specs(machine_type)
                        
                        # Sum disk sizes safely, handling potential None values
                        disk_size = 0
                        for d in instance.disks:
                            try:
                                # disk_size_gb is usually present, but for some disk types it might be 0/None
                                sz = getattr(d, 'disk_size_gb', 0)
                                if sz:
                                    disk_size += int(sz)
                            except (ValueError, TypeError):
                                pass

                        resources.append({
                            "_raw_type": "compute.googleapis.com/Instance",
                            "_normalized_type": "compute/Instance",
                            "_project": project_id,
                            "_machine_type": machine_type,
                            "_resource_data": {},
                            "Name": instance.name,
                            "GCP Service": "compute",
                            "Resource Type": "Instance",
                            "Machine Type / Class": machine_type,
                            "vCPUs": vcpus,
                            "Memory (GB)": memory_gb,
                            "Storage Size (GB)": disk_size or "",
                            "Storage Class": "",
                            "Location": zone_name,
                            "Project": project_id,
                            "Status": instance.status,
                            "Labels": json.dumps(dict(instance.labels) if instance.labels else {}),
                            "Azure Service": get_azure_service("compute/Instance"),
                            "Azure VM SKU": GLOBAL_SKU_MAPPER.map_gcp_to_azure(machine_type, "VM"),
                            "Azure Region": get_azure_region(zone_name),
                        })
        except Exception as e:
            msg = str(e)
            if "PermissionDenied" in msg or "403" in msg:
                print(f"   ❌ Access Denied: Missing 'roles/compute.viewer' in {project_id}.")
            else:
                print(f"   ⚠️ Compute scan failed for {project_id}: {e}")
        return resources

    def _scan_storage_buckets(self, project_id):
        """Scan all Cloud Storage buckets with size."""
        resources = []
        try:
            sc = self._get_storage_client(project_id)
            for bucket in sc.list_buckets():
                size_gb = "—"
                try:
                    def _measure():
                        # Use same logic as enrich_item for consistency
                        blobs = sc.list_blobs(bucket.name)
                        total = sum(b.size or 0 for b in blobs)
                        return round(total / (1024 ** 3), 4)
                    
                    # We still keep a small executor here for timeout protection
                    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as _ex:
                        fut = _ex.submit(_measure)
                        # Increased timeout to 30s as listing can take time
                        size_gb = fut.result(timeout=30) 
                except (concurrent.futures.TimeoutError, Exception):
                    size_gb = "—" # Mark for enrichment or report as unknown if too large

                resources.append({
                    "_raw_type": "storage.googleapis.com/Bucket",
                    "_normalized_type": "storage/Bucket",
                    "_project": project_id,
                    "_machine_type": "",
                    "_resource_data": {},
                    "Name": bucket.name,
                    "GCP Service": "storage",
                    "Resource Type": "Bucket",
                    "Machine Type / Class": bucket.storage_class or "STANDARD",
                    "vCPUs": "",
                    "Memory (GB)": "",
                    "Storage Size (GB)": size_gb,
                    "Storage Class": bucket.storage_class,
                    "Location": bucket.location,
                    "Project": project_id,
                    "Status": "Active",
                    "Labels": json.dumps(dict(bucket.labels) if bucket.labels else {}),
                    "Azure Service": get_azure_service("storage/Bucket"),
                    "Azure VM SKU": "—",
                    "Azure Region": get_azure_region(bucket.location),
                })
        except Exception as e:
            msg = str(e)
            if "PermissionDenied" in msg or "403" in msg:
                print(f"   ❌ Access Denied: Missing 'roles/storage.objectViewer'.")
            else:
                print(f"   ⚠️ Storage scan failed: {e}")
        return resources


# ============================================================
# Report generation
# ============================================================

DISPLAY_COLUMNS = GCP_COLUMNS + AZURE_COLUMNS


def generate_gcp_report(resources, duration=None):
    """Generate a clean Excel report from extracted GCP resources."""
    if not resources:
        print("⚠️ No resources to export")
        return None

    print(f"📊 Generating GCP-to-Azure report for {len(resources)} resources...")

    # Drop internal columns
    clean_resources = []
    for r in resources:
        clean = {k: v for k, v in r.items() if not k.startswith("_")}
        clean_resources.append(clean)

    df = pd.DataFrame(clean_resources)
    
    # Drop columns where ALL values are effectively empty: [], "[]", None, or ""
    def is_empty_val(v):
        if v is None or (isinstance(v, float) and pd.isna(v)): return True
        if isinstance(v, list) and not v: return True
        sv = str(v).strip()
        if sv in ["", "[]"]: return True
        return False

    cols_to_drop = []
    for col in df.columns:
        if df[col].apply(is_empty_val).all():
            cols_to_drop.append(col)
    
    # Don't drop base columns to maintain structure
    protected_cols = DISPLAY_COLUMNS
    cols_to_drop = [c for c in cols_to_drop if c not in protected_cols]
    
    if cols_to_drop:
        print(f"Dropping {len(cols_to_drop)} empty columns: {cols_to_drop}")
        df.drop(columns=cols_to_drop, inplace=True)

    # Generate output file with UUID to prevent collisions
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    import uuid as _uuid
    unique_id = str(_uuid.uuid4())[:8]
    filename = f"gcp_to_azure_assessment_{timestamp}_{unique_id}.xlsx"
    filepath = os.path.join(os.path.dirname(__file__), filename)

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # --- Summary Sheet ---
        summary_df = _create_summary(df)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        # Add Evaluation Metadata
        ws_summary = writer.sheets["Summary"]
        last_item_row = len(summary_df) + 3
        
        ws_summary.cell(row=last_item_row, column=1, value="Assessment Date")
        ws_summary.cell(row=last_item_row, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        if duration:
            ws_summary.cell(row=last_item_row+1, column=1, value="Assessment Duration")
            ws_summary.cell(row=last_item_row+1, column=2, value=duration)
            
        # Style the metadata labels
        from openpyxl.styles import Font
        ws_summary.cell(row=last_item_row, column=1).font = Font(bold=True)
        if duration:
            ws_summary.cell(row=last_item_row+1, column=1).font = Font(bold=True)

        # --- All Resources Sheet ---
        all_cols = [c for c in DISPLAY_COLUMNS if c in df.columns]
        df[all_cols].to_excel(writer, sheet_name="All Resources", index=False)

        # --- Per Service-Type Sheets ---
        grouped = df.groupby("Resource Type")
        for rtype, group_df in grouped:
            sheet_name = str(rtype)[:31]
            # sub = group_df[all_cols].copy() - Redundant, use all columns present in this group
            sub = group_df.copy()

            # Aggressive drop for this specific resource type
            sub_cols_to_drop = []
            for col in sub.columns:
                if sub[col].apply(is_empty_val).all() and col not in protected_cols:
                    sub_cols_to_drop.append(col)
            
            if sub_cols_to_drop:
                sub.drop(columns=sub_cols_to_drop, inplace=True)

            sub.to_excel(writer, sheet_name=sheet_name, index=False)

    # Apply formatting
    _format_excel(filepath)

    print(f"✅ Report generated: {filepath}")
    return filepath


def _create_summary(df):
    """Create summary statistics sheet."""
    total = len(df)
    projects = df["Project"].nunique() if "Project" in df.columns else 0
    services = df["Resource Type"].nunique() if "Resource Type" in df.columns else 0
    locations = df["Location"].nunique() if "Location" in df.columns else 0

    service_counts = df["Resource Type"].value_counts().reset_index()
    service_counts.columns = ["Resource Type", "Count"]

    meta = pd.DataFrame({
        "Metric": ["Total Resources", "Total Projects", "Unique Resource Types", "Unique Locations"],
        "Value": [total, projects, services, locations]
    })

    # Create empty rows as separator
    sep = pd.DataFrame({"Metric": ["", "Resource Breakdown"], "Value": ["", ""]})

    # Combine
    summary_data = []
    for _, row in meta.iterrows():
        summary_data.append({"Metric": row["Metric"], "Value": row["Value"]})
    summary_data.append({"Metric": "", "Value": ""})
    summary_data.append({"Metric": "--- Resource Breakdown ---", "Value": "Count"})
    for _, row in service_counts.iterrows():
        summary_data.append({"Metric": row["Resource Type"], "Value": row["Count"]})

    return pd.DataFrame(summary_data)


def _format_excel(filepath):
    """Apply color-coded formatting: GCP columns = orange header, Azure columns = blue header."""
    wb = load_workbook(filepath)

    gcp_fill = PatternFill(start_color="E87722", end_color="E87722", fill_type="solid")    # Orange - GCP
    azure_fill = PatternFill(start_color="0078D4", end_color="0078D4", fill_type="solid")  # Azure Blue
    summary_fill = PatternFill(start_color="1A1A2E", end_color="1A1A2E", fill_type="solid")  # Dark header

    white_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style='thin', color="D0D0D0"),
        right=Side(style='thin', color="D0D0D0"),
        top=Side(style='thin', color="D0D0D0"),
        bottom=Side(style='thin', color="D0D0D0"),
    )

    # Alternating row fills
    gcp_row_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    azure_row_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F8F8F8", end_color="F8F8F8", fill_type="solid")

    for ws in wb.worksheets:
        if ws.title == "Summary":
            # Simple summarize formatting
            for cell in ws[1]:
                cell.fill = summary_fill
                cell.font = white_font
                cell.border = border
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.border = border
        else:
            # Colour-code headers based on GCP vs Azure column
            header_row = list(ws[1])
            azure_col_indices = set()

            for i, cell in enumerate(header_row, start=1):
                col_name = str(cell.value or "")
                if col_name in AZURE_COLUMNS:
                    cell.fill = azure_fill
                    azure_col_indices.add(i)
                else:
                    cell.fill = gcp_fill
                cell.font = white_font
                cell.border = border
                cell.alignment = Alignment(wrap_text=False, vertical="center")

            # Data rows: colour the Azure cells light blue, GCP cells light orange
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column), start=2):
                for col_idx, cell in enumerate(row, start=1):
                    cell.border = border
                    cell.alignment = Alignment(vertical="center")
                    if col_idx in azure_col_indices:
                        cell.fill = azure_row_fill
                    elif row_idx % 2 == 0:
                        cell.fill = alt_row_fill

        # Auto-fit columns
        for column in ws.columns:
            max_length = 0
            col_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_length + 3, 55)

        # Freeze header row
        ws.freeze_panes = "A2"

    wb.save(filepath)
    print("✅ Excel formatted with GCP (orange) / Azure (blue) colour coding")
