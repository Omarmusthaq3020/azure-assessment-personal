import boto3
import pandas as pd
import os
import re
from datetime import datetime, timedelta
import uuid
from botocore.config import Config
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


def _short_err(e):
    """Truncate boto3's verbose "Unknown service: 'x'. Valid service names
    are: <giant list>" exception messages so error logs stay readable.
    Also caps any single error message at 300 chars. Module-level so it's
    usable from any method, not just _scan_native_essentials."""
    s = str(e)
    if "Valid service names are" in s:
        return s.split(". Valid service names are")[0]
    return s if len(s) < 300 else s[:300] + "..."


# Azure managed-disk SKU tiers (Standard SSD / HDD and Premium SSD all use
# the same size ladder). When migrating an AWS disk to Azure, the closest
# provisionable size is the next tier ≥ the AWS size.
_AZURE_DISK_TIERS = [4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32767]


def _round_to_azure_disk_tier(size_gb):
    """Round an AWS disk size (GB) up to the next Azure managed-disk SKU
    tier. Returns 0 for zero/negative, None for missing/non-numeric."""
    if size_gb is None:
        return None
    try:
        size = float(size_gb)
    except (ValueError, TypeError):
        return None
    if pd.isna(size):
        return None
    if size <= 0:
        return 0
    for tier in _AZURE_DISK_TIERS:
        if size <= tier:
            return tier
    return _AZURE_DISK_TIERS[-1]


try:
    from sku_mapper import SkuMapper
except ImportError:
    import sys
    import os
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from sku_mapper import SkuMapper

GLOBAL_SKU_MAPPER = SkuMapper()

def get_azure_sku_recommendation(aws_sku, resource_type):
    """
    Maps AWS Instance Types to Azure VM/DB Series recommendations using SkuMapper.
    """
    try:
        if GLOBAL_SKU_MAPPER:
             return GLOBAL_SKU_MAPPER.map_sku(aws_sku, resource_type)
        return "Mapper Not Loaded"
    except Exception as e:
        return f"Mapping Error: {str(e)}"

class AwsResourceExtractor:
    def __init__(self, aws_access_key_id, aws_secret_access_key, region_name, aws_session_token=None):
        self.session = boto3.Session(
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key,
            aws_session_token=aws_session_token,
            region_name=region_name
        )
        self.region = region_name
        # Fail-fast configuration — applied to every boto3 client created via
        # self._client(...) so a single slow / unreachable service can't hang
        # the whole assessment for minutes. Default boto3 is 60s connect + 60s
        # read with 5 retries → up to 5 minutes per failed call.
        # MUST be set before any _client() call below.
        self._fast_config = Config(
            retries={'max_attempts': 2, 'mode': 'standard'},
            connect_timeout=5,
            read_timeout=10,
        )
        # Keep clients cached
        self.tagging_client = self._client('resourcegroupstaggingapi')
        self.cloud_control = self._client('cloudcontrol')

    def _client(self, service_name, **kwargs):
        """Get a boto3 client with fast-fail timeouts applied by default.

        Pass `config=...` to override. Otherwise uses self._fast_config
        (5s connect / 10s read / 2 retries).
        """
        kwargs.setdefault('config', self._fast_config)
        # NOTE: call session.client directly here — going through self._client
        # would infinite-recurse since the global replace previously affected
        # this method body too.
        return self.session.client(service_name, **kwargs)

    def _prepare_networking_metadata(self):
        """
        Builds a cache of VPC/Subnet relationships and identifies Public vs Private subnets.
        A subnet is considered 'Public' if its associated route table has a route to an Internet Gateway.
        """
        print(f"Pre-fetching networking metadata for {self.region}...", flush=True)
        self._subnet_cache = {}
        try:
            ec2 = self._client('ec2')
            
            # 1. Fetch all Subnets
            subnets = []
            paginator = ec2.get_paginator('describe_subnets')
            for page in paginator.paginate():
                for s in page['Subnets']:
                    self._subnet_cache[s['SubnetId']] = {
                        'VpcId': s['VpcId'],
                        'Cidr': s['CidrBlock'],
                        'IsPublic': False # Default
                    }
            
            # 2. Fetch all Route Tables to determine Public status
            paginator = ec2.get_paginator('describe_route_tables')
            for page in paginator.paginate():
                for rt in page['RouteTables']:
                    is_public_rt = False
                    for route in rt.get('Routes', []):
                        if str(route.get('GatewayId', '')).startswith('igw-'):
                            is_public_rt = True
                            break
                    
                    if is_public_rt:
                        # Check explicit associations
                        for assoc in rt.get('Associations', []):
                            sid = assoc.get('SubnetId')
                            if sid in self._subnet_cache:
                                self._subnet_cache[sid]['IsPublic'] = True
                        
                        # Main route table fallback
                        if any(assoc.get('Main') for assoc in rt.get('Associations', [])):
                            vpc_id = rt['VpcId']
                            for sid, sdata in self._subnet_cache.items():
                                if sdata['VpcId'] == vpc_id:
                                    sdata.setdefault('IsPublic', True)

        except Exception as e:
            print(f"   ⚠️ Networking metadata fetch failed: {_short_err(e)}")

    def _map_security_group_attachments(self, all_resources):
        """
        Map security groups back to the EC2 instances they protect,
        resolving instance Names where possible.
        """
        self._sg_attachment_cache = {}
        try:
            # 1. Map IDs to Names
            inst_names = {}
            for r in all_resources:
                if 'EC2::Instance' in r.get('NormalizedResourceType', ''):
                    inst_names[r['ResourceId']] = r['ResourceName']
            
            # 2. Get SG Associations via ENIs
            ec2 = self._client('ec2')
            eni_paginator = ec2.get_paginator('describe_network_interfaces')
            for eni_page in eni_paginator.paginate():
                for eni in eni_page['NetworkInterfaces']:
                    inst_id = eni.get('Attachment', {}).get('InstanceId')
                    if inst_id:
                         name = inst_names.get(inst_id, inst_id)
                         for sg in eni.get('Groups', []):
                             sg_id = sg['GroupId']
                             if sg_id not in self._sg_attachment_cache:
                                 self._sg_attachment_cache[sg_id] = set()
                             self._sg_attachment_cache[sg_id].add(name)
        except: pass

    def _get_connectivity_info(self, props, r_type):
        """
        Compiles a human-readable connectivity string for the report.
        """
        if not props: return ""
        info = []
        r_type = str(r_type)
        
        # 1. Attachment Info (Critical for Volumes, EIPs, ENIs)
        inst_id = props.get('InstanceId') or props.get('Attachment', {}).get('InstanceId')
        if not inst_id and 'Attachments' in props and props['Attachments']:
            inst_id = props['Attachments'][0].get('InstanceId')

        if inst_id:
            info.append(f"Attached to: {inst_id}")
            
        # 2. VPC/Subnet Info
        vpc_id = props.get('VpcId')
        if vpc_id:
            info.append(f"VPC: {vpc_id}")
            
        subnet_id = props.get('SubnetId')
        if not subnet_id and 'SubnetIds' in props:
            sids = props['SubnetIds']
            if isinstance(sids, list):
                subnet_id = sids[0] if sids else None
                if len(sids) > 1: info.append(f"Subnets: [{len(sids)} total]")
            else:
                subnet_id = sids
                
        if subnet_id:
            scope = "Private" # Default
            if hasattr(self, '_subnet_cache') and subnet_id in self._subnet_cache:
                scope = "Public" if self._subnet_cache[subnet_id].get('IsPublic') else "Private"
            info.append(f"Subnet: {subnet_id} ({scope})")
            
        # 3. IP / DNS Info
        public_ip = props.get('PublicIp') or props.get('PublicIpAddress') or props.get('IpAddress')
        if public_ip:
            info.append(f"Public IP: {public_ip}")
            
        private_ip = props.get('PrivateIp') or props.get('PrivateIpAddress')
        if private_ip:
            info.append(f"Private IP: {private_ip}")
            
        dns = props.get('DNSName') or props.get('PublicDnsName') or props.get('EndpointAddress')
        if dns:
            info.append(f"DNS: {dns}")

        # 4. Port Info
        port = props.get('Port') or props.get('FromPort')
        if port and port != -1:
             info.append(f"Port: {port}")

        # 5. Security Rule Context (Inherited mapping from SG-to-EC2 cache)
        if 'SecurityGroupIngress' in r_type or 'SecurityGroupEgress' in r_type:
            direction = props.get('Direction', 'Rule')
            parent = props.get('ParentGroup', 'Group')
            
            # Map SG rules back to the instances they protect
            sg_id = props.get('GroupId')
            attachments = list(self._sg_attachment_cache.get(sg_id, [])) if hasattr(self, '_sg_attachment_cache') else []
            att_str = f" | Protects: {', '.join(attachments)}" if attachments else ""
            
            # Fix protocol formatting (Handle -1 as All Traffic)
            proto = str(props.get('IpProtocol', '')).upper()
            if proto == '-1' or proto == 'ALL': 
                proto = 'ALL TRAFFIC'
                port_str = 'All Ports'
            else:
                fp = props.get('FromPort')
                tp = props.get('ToPort')
                port_str = f"Port: {fp}" if fp == tp else f"Port: {fp}-{tp}"
            
            return f"{direction} for {parent} | {proto}:{port_str}{att_str}"

        return " | ".join(info) if info else "N/A"

    def extract_resources(self, progress_callback=None):
        """
        Extract ALL resources from AWS using Resource Groups Tagging API
        and detailed clients for enrichment.
        """
        self.progress_callback = progress_callback
        def log_progress(step, pct, msg):
            if self.progress_callback:
                self.progress_callback(step, pct, msg)

        log_progress("Initializing", 5, f"Starting AWS extraction in {self.region}...")
        print(f"Fetching ALL resources in region {self.region}...", flush=True)
        self._ensure_cfn_type_cache() # Ensure metadata is loaded first
        self._prepare_networking_metadata() # Fetch VPC/Subnet status up-front

        all_resources = []
        known_arns = set()

        # Production-grade discovery: NATIVE APIs ONLY.
        # We deliberately do NOT use:
        #   - Resource Groups Tagging API (structurally undercounts: untagged
        #     resources are invisible, and the permission is frequently denied).
        #   - Cloud Control API list_resources (no count contract: filters
        #     internally, returns partial lists on throttle, varies by service).
        # Native describe_*/list_* calls are AWS's canonical enumeration source
        # (same APIs the Console and CLI use) and are the only way to guarantee
        # deterministic counts that exactly match the AWS Console.
        # If a service shows 0 in the report, it is either truly empty or a
        # native scan needs to be added in _scan_native_essentials.
        print(f"🔒 Discovery mode: NATIVE APIs only (deterministic, console-accurate).", flush=True)
        log_progress("Extracting", 15, "Running native API discovery (no tagging/Cloud Control)...")

        # 1.1 Direct Native Scanning for Critical Active Services
        self._scan_native_active_services(all_resources, known_arns)

        # 1.5 Native Essentials — comprehensive per-service describe_*/list_*
        # calls (~62 service types covered). This is the sole authoritative
        # source for resource counts.
        try:
            self._scan_native_essentials(all_resources, known_arns)
        except Exception as e:
            print(f"   ⚠️ Native Essentials fail: {e}")

        # 1.7 Map SG Rules to named instances (must run after core discovery)
        self._map_security_group_attachments(all_resources)
        
        log_progress("Enriching", 40, f"Discovered {len(all_resources)} total resources. Starting enrichment...")
        print(f"Total resources after discovery: {len(all_resources)}. Starting enrichment...")
        
        # 2. Enrich high-value resources
        # We process the list and update 'Status' and 'Details' for specific types
        try:
            self._enrich_resources(all_resources)
        except Exception as e:
            print(f"⚠️ Enrichment partially failed: {_short_err(e)}")
        
        # 3. Clean up internal keys
        for r in all_resources:
            r.pop('_arn', None)
            r.pop('_service', None)
            r.pop('_type_id', None)
            r.pop('_id_value', None)

        # Debug: Print found RDS instances
        rds_found = [r['ResourceName'] for r in all_resources if 'RDS' in r.get('NormalizedResourceType', '') or 'rds' in r.get('ResourceType', '')]
        print(f"DEBUG: Final List of RDS Instances in Memory: {rds_found}")

        return all_resources

    def _scan_native_essentials(self, all_resources, known_arns):
        """
        Scan specific resource types using native Boto3 clients to ensure 100% coverage
        and detailed attributes (Enrichment) that Cloud Control often misses.
        """
        print("Scanning Native Services for deep enrichment (S3, IAM, RDS, ELB, Lambda)...")
        found_count = 0

        # Shared helper: pull the "Name" tag from a Tags list and fall back
        # to the resource ID when no Name tag is set. Used by VPC, Subnet,
        # NetworkInterface, EBS Volume — and any other scan whose resource
        # shape is AWS's standard {Tags: [{Key, Value}, ...]} pattern. This
        # is what makes the report show "prod-vpc" instead of "vpc-0e0e...".
        def _name_or_id(resource, id_field_or_value, tags_key='Tags'):
            if isinstance(id_field_or_value, str):
                fallback = id_field_or_value
            else:
                fallback = ''
            for t in (resource.get(tags_key) or []):
                if isinstance(t, dict) and t.get('Key') == 'Name':
                    v = t.get('Value')
                    if v:
                        return v
            return fallback

        # Region availability for AI/ML services that aren't GA in every region.
        # When the assessment region isn't in this list, we skip the scan
        # silently instead of letting boto3 raise a "Could not connect" error.
        # Source: AWS regional services list (verified late 2025).
        _AI_SERVICE_REGIONS = {
            'lexv2-models': {
                'us-east-1', 'us-west-2', 'eu-west-1', 'eu-west-2', 'eu-central-1',
                'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1', 'ap-northeast-2',
                'ca-central-1', 'us-east-2',
            },
            'codeguruprofiler': {
                'us-east-1', 'us-east-2', 'us-west-2', 'eu-west-1', 'eu-central-1',
                'eu-north-1', 'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1',
            },
            'qbusiness': {
                'us-east-1', 'us-west-2', 'eu-west-1',
            },
        }

        def _service_available_here(svc):
            """True if `svc` is GA in self.region (or unrestricted)."""
            allowed = _AI_SERVICE_REGIONS.get(svc)
            return allowed is None or self.region in allowed

        # Helper maps for Upsert/Enrichment
        # We index by ARN and ID to match existing Tagging results
        res_map_arn = {r.get('_arn'): r for r in all_resources if r.get('_arn')}
        res_map_id = {r.get('_id_value'): r for r in all_resources if r.get('_id_value')}

        # Matches AWS resource-ID patterns like i-0abc..., sg-0abc...,
        # vol-0abc..., subnet-0abc..., eipalloc-0abc..., etc. Used so the
        # upsert prefers a human-friendly name (Name tag) over a record
        # whose ResourceName is just the raw ID.
        _id_like_re = re.compile(r'^[a-z][a-z0-9-]*-[a-f0-9]{8,}$')

        def _is_id_like(s):
            return bool(_id_like_re.match(str(s or '').strip()))

        def upsert_resource(arn, name, type_norm, group, details_dict, status='Active', tags_list=[]):
            nonlocal found_count

            # 2. Check if exists
            # Try ARN match first (definitively the same resource).
            existing = res_map_arn.get(arn)
            if not existing and name:
                # Fall back to name-based match — but ONLY when the existing
                # record has the same NormalizedResourceType. Otherwise it's
                # a cross-type name collision (e.g., S3 bucket "my-app-prod"
                # and ECR repo "my-app-prod" share a name) and merging them
                # would silently overwrite the first resource's type, making
                # it disappear from its sheet.
                candidate = res_map_id.get(name)
                if candidate and candidate.get('NormalizedResourceType') == type_norm:
                    existing = candidate
            # ARN-tail fallback: _scan_native_active_services adds resources
            # with `_id_value=instance_id` ('i-xxxx'), while essentials upserts
            # with the FULL ARN ('arn:aws:ec2:...:instance/i-xxxx') and a
            # human-friendly `name` (Name tag). Without this fallback the two
            # records never merge — the active-services row keeps its
            # ID-as-ResourceName and the essentials row gets dropped during
            # global dedup, hiding the real Name tag from the report.
            if not existing and arn and arn.startswith('arn:'):
                tail = arn.replace(':', '/').rsplit('/', 1)[-1]
                if tail:
                    candidate = res_map_id.get(tail)
                    if candidate and candidate.get('NormalizedResourceType') == type_norm:
                        existing = candidate

            # 1. Prepare standardized record
            # Flatten details for report columns
            flat_details = self._flatten_dict(details_dict)
            
            # Stringify for the 'Details' column
            pretty_details = ", ".join([f"{k}: {v}" for k,v in details_dict.items() if isinstance(v, (str, int, float))])
            
            connectivity = self._get_connectivity_info(details_dict, type_norm)

            # Standardize and Priority Status
            final_status = str(status) if status else 'Active'
            # Map common healthy states to 'Active'
            if final_status.lower() in ['available', 'running', 'in-use', 'creating', 'stable', 'active']:
                final_status = 'Active'

            common_props = {
                'Status': final_status,
                'Details': pretty_details,
                'Connectivity': connectivity,
                'Tags': self._format_tags(tags_list)
            }
            
            # If existing status is 'Unknown', always override.
            if existing: 
                current_status = str(existing.get('Status', 'Unknown'))
                if final_status == 'Unknown' and current_status != 'Unknown':
                    # Don't downgrade status from a real value to 'Unknown'
                    common_props.pop('Status', None)

            # Merge flat details into the main dict (High-level columns like 'Engine', 'Memory')
            common_props.update(flat_details)
                 
            if existing:
                # ENRICH existing
                existing.update(common_props)
                # Enforce Type consistency
                existing['ResourceType'] = type_norm
                existing['NormalizedResourceType'] = type_norm
                # Upgrade ResourceName to a human-friendly value when the
                # current one is just an AWS resource ID. `_scan_native_active_services`
                # adds resources with `ResourceName=InstanceId` ('i-xxxx');
                # essentials passes the Name tag here. Promote the Name tag
                # so the report shows readable names instead of opaque IDs.
                cur_name = str(existing.get('ResourceName') or '').strip()
                new_name = str(name or '').strip()
                if new_name and new_name != cur_name:
                    if not cur_name or (_is_id_like(cur_name) and not _is_id_like(new_name)):
                        existing['ResourceName'] = new_name
                        # Also bump _id_value so future lookups by friendly
                        # name still find this record.
                        existing['_id_value'] = new_name
            else:
                # ADD new
                if arn not in known_arns:
                    known_arns.add(arn)
                    if name: known_arns.add(name)
                    
                    new_r = {
                        'ResourceId': arn,
                        'ResourceName': name,
                        'ResourceType': type_norm,
                        'NormalizedResourceType': type_norm,
                        'ResourceTypeSource': 'Native',
                        'ResourceGroup': group,
                        'Location': self.region,
                        '_arn': arn,
                        '_service': type_norm.split('::')[1].lower() if '::' in type_norm else 'unknown',
                        '_type_id': type_norm.split('::')[2].lower() if '::' in type_norm else 'unknown',
                        '_id_value': name
                    }
                    new_r.update(common_props)
                    all_resources.append(new_r)
                    found_count += 1

        # A. S3 Buckets
        try:
            s3 = self._client('s3')
            resp = s3.list_buckets()
            if 'Buckets' in resp:
                for b in resp['Buckets']:
                    name = b['Name']
                    props = {'CreationDate': str(b['CreationDate'])}
                    
                    # Fetch detailed configuration for migration assessment
                    try:
                        loc = s3.get_bucket_location(Bucket=name)
                        # LocationConstraint is None for us-east-1
                        props['Location'] = loc.get('LocationConstraint') or 'us-east-1'
                    except: props['Location'] = 'Unknown'
                    
                    try:
                        ver = s3.get_bucket_versioning(Bucket=name)
                        props['Versioning'] = ver.get('Status', 'Disabled')
                    except: props['Versioning'] = 'Unknown'

                    try:
                         enc = s3.get_bucket_encryption(Bucket=name)
                         rules = enc.get('ServerSideEncryptionConfiguration', {}).get('Rules', [])
                         if rules:
                             props['Encryption'] = rules[0].get('ApplyServerSideEncryptionByDefault', {}).get('SSEAlgorithm', 'Unknown')
                    except: props['Encryption'] = 'None'

                    # infer Storage Class and Size
                    # 1. Lifecycle
                    try:
                        lc = s3.get_bucket_lifecycle_configuration(Bucket=name)
                        rules = lc.get('Rules', [])
                        # Summarize transitions
                        transitions = []
                        for r in rules:
                            for t in r.get('Transitions', []):
                                if 'StorageClass' in t: transitions.append(t['StorageClass'])
                        if transitions:
                            props['StorageClass'] = f"Standard -> {', '.join(set(transitions))}"
                        else:
                            props['StorageClass'] = "Standard"
                    except: 
                        props['StorageClass'] = "Standard" # Default
                        
                    # 2. Size (CloudWatch)
                    try:
                        cw = self._client('cloudwatch')
                        end_time = datetime.utcnow()
                        start_time = end_time - timedelta(days=2) # Check last 48h to be safe
                        
                        stats = cw.get_metric_statistics(
                            Namespace='AWS/S3',
                            MetricName='BucketSizeBytes',
                            Dimensions=[
                                {'Name': 'BucketName', 'Value': name},
                                {'Name': 'StorageType', 'Value': 'StandardStorage'}
                            ],
                            StartTime=start_time,
                            EndTime=end_time,
                            Period=86400,
                            Statistics=['Average']
                        )
                        if stats['Datapoints']:
                            # Sort by time to get latest
                            latest = sorted(stats['Datapoints'], key=lambda x: x['Timestamp'])[-1]
                            bytes_size = latest['Average']
                            gb_size = bytes_size / (1024**3)
                            props['Size (GB)'] = f"{gb_size:.2f}"
                        else:
                             props['Size (GB)'] = "0 (No Metrics)"
                    except Exception as e: 
                        # print(f"S3 Size Error: {e}")
                        props['Size (GB)'] = "Unknown"

                    upsert_resource(
                        arn=f"arn:aws:s3:::{name}",
                        name=name,
                        type_norm='AWS::S3::Bucket',
                        group='Storage',
                        details_dict=props,
                        status='Active'
                    )
        except Exception: pass

        # B. IAM Roles
        try:
            iam = self._client('iam')
            paginator = iam.get_paginator('list_roles')
            for page in paginator.paginate():
                for r in page['Roles']:
                    upsert_resource(
                        arn=r['Arn'],
                        name=r['RoleName'],
                        type_norm='AWS::IAM::Role',
                        group='Identity',
                        details_dict={'CreateDate': str(r['CreateDate']), 'Path': r.get('Path')},
                        status='Active'
                    )
        except Exception: pass

        # C. EC2 KeyPairs
        try:
            ec2 = self._client('ec2')
            resp = ec2.describe_key_pairs()
            if 'KeyPairs' in resp:
                for k in resp['KeyPairs']:
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:key-pair/{k['KeyName']}",
                        name=k['KeyName'],
                        type_norm='AWS::EC2::KeyPair',
                        group='Compute',
                        details_dict={'Fingerprint': k.get('KeyFingerprint')},
                        status='Active',
                        tags_list=k.get('Tags', [])
                    )
        except Exception: pass
        
        # D. ACM Certificates
        try:
            acm = self._client('acm')
            paginator = acm.get_paginator('list_certificates')
            for page in paginator.paginate():
                for c in page['CertificateSummaryList']:
                    upsert_resource(
                        arn=c['CertificateArn'],
                        name=c['DomainName'],
                        type_norm='AWS::CertificateManager::Certificate',
                        group='Security',
                        details_dict={'Domain': c['DomainName'], 'Status': c.get('Status')},
                        status=c.get('Status', 'Active')
                    )
        except Exception: pass

        # E. RDS Databases (Detailed)
        try:
            rds = self._client('rds')
            paginator = rds.get_paginator('describe_db_instances')
            for page in paginator.paginate():
                for db in page['DBInstances']:
                    name = db['DBInstanceIdentifier']
                    print(f"RDS Native found: {name} ({db['DBInstanceStatus']})") 
                    
                    # Capture ALL useful columns safely
                    props = {
                        'Engine': db.get('Engine'),
                        'EngineVersion': db.get('EngineVersion'),
                        'DBInstanceClass': db.get('DBInstanceClass'),
                        'AllocatedStorage': db.get('AllocatedStorage'),
                        'StorageType': db.get('StorageType', 'standard'),
                        'MultiAZ': db.get('MultiAZ'),
                        'PubliclyAccessible': db.get('PubliclyAccessible'),
                        'VpcId': db.get('DBSubnetGroup', {}).get('VpcId'),
                        'AvailabilityZone': db.get('AvailabilityZone'),
                        'SubnetIds': [s['SubnetIdentifier'] for s in db.get('DBSubnetGroup', {}).get('Subnets', [])],
                        'EndpointAddress': db.get('Endpoint', {}).get('Address'),
                        'PublicIp': db.get('Endpoint', {}).get('Address'), # For Connectivity helper
                        'Port': db.get('Endpoint', {}).get('Port')
                    }
                    upsert_resource(
                        arn=db['DBInstanceArn'],
                        name=name, # Force Name to be Identifier
                        type_norm='AWS::RDS::DBInstance',
                        group='Database',
                        details_dict=props,
                        status=db['DBInstanceStatus'],
                        tags_list=db.get('TagList', [])
                    )
                    # Explicitly update ResourceName in all_resources to overwrite Tag-derived name "suite" if it exists
                    # This ensures "p-mysql" shows as "p-mysql" not "suite" if it has the same tag
                    for r in all_resources:
                        if r.get('ResourceId') == db['DBInstanceArn']:
                            r['ResourceName'] = name
        except Exception: pass

        # F. Load Balancers (ELBv2 - ALB/NLB)
        try:
            elbv2 = self._client('elbv2')
            paginator = elbv2.get_paginator('describe_load_balancers')
            for page in paginator.paginate():
                for lb in page['LoadBalancers']:
                    props = {
                        'Type': lb['Type'],
                        'Scheme': lb['Scheme'],
                        'VpcId': lb.get('VpcId'),
                        'SubnetIds': [az['SubnetId'] for az in lb.get('AvailabilityZones', [])],
                        'DNSName': lb.get('DNSName'),
                        'IpAddressType': lb.get('IpAddressType'),
                        'State': lb.get('State', {}).get('Code')
                    }
                    upsert_resource(
                        arn=lb['LoadBalancerArn'],
                        name=lb['LoadBalancerName'],
                        type_norm='AWS::ElasticLoadBalancingV2::LoadBalancer',
                        group='Networking',
                        details_dict=props,
                        status=lb.get('State', {}).get('Code', 'Unknown'),
                        tags_list=[] # ELBv2 requires separate call for tags, skipping for speed/rate-limit
                    )
        except Exception: pass

        # G. Lambda Functions
        try:
            lam = self._client('lambda')
            paginator = lam.get_paginator('list_functions')
            for page in paginator.paginate():
                for f in page['Functions']:
                    props = {
                        'Runtime': f.get('Runtime'),
                        'Handler': f.get('Handler'),
                        'CodeSize': f.get('CodeSize'),
                        'LastModified': f.get('LastModified'),
                        'MemorySize': f.get('MemorySize')
                    }
                    upsert_resource(
                        arn=f['FunctionArn'],
                        name=f['FunctionName'],
                        type_norm='AWS::Lambda::Function',
                        group='Compute',
                        details_dict=props,
                        status='Active' # Lambda doesn't have a status in list output usually
                    )
        except Exception: pass

        # H. EBS Volumes
        try:
            ec2 = self._client('ec2')
            paginator = ec2.get_paginator('describe_volumes')
            for page in paginator.paginate():
                for v in page['Volumes']:
                    props = {
                        'Size': v['Size'],
                        'VolumeType': v['VolumeType'],
                        'Iops': v.get('Iops'),
                        'Throughput': v.get('Throughput'),
                        'Encrypted': v.get('Encrypted'),
                        'InstanceId': v['Attachments'][0].get('InstanceId') if v.get('Attachments') else None,
                        'Device': v['Attachments'][0].get('Device') if v.get('Attachments') else None,
                        'State': v['State']
                    }
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{v['AvailabilityZone'][:-1]}:volume/{v['VolumeId']}", # Best effort ARN
                        name=_name_or_id(v, v['VolumeId']),
                        type_norm='AWS::EC2::Volume',
                        group='Storage',
                        details_dict=props,
                        status=v['State'],
                        tags_list=v.get('Tags', [])
                    )
        except Exception: pass

        # H2. EC2 Instances (Native for OS/Storage details)
        try:
            ec2 = self._client('ec2')

            # Resolve account_id ONCE (was being fetched per-instance via STS
            # before — N instances = N pointless STS calls).
            try:
                _ec2_account_id = self._client('sts').get_caller_identity()['Account']
            except Exception:
                _ec2_account_id = '000000000000'

            # Build a rich volume cache: size + type + IOPS + encrypted +
            # snapshot id. describe_instances only returns VolumeId references,
            # so a single pass over describe_volumes here lets us enrich every
            # data disk per instance without N×instance API calls.
            vol_cache = {}
            try:
                vol_paginator = ec2.get_paginator('describe_volumes')
                for vol_page in vol_paginator.paginate():
                    for vol in vol_page['Volumes']:
                        vol_cache[vol['VolumeId']] = {
                            'size': vol.get('Size', 0),
                            'type': vol.get('VolumeType'),
                            'iops': vol.get('Iops'),
                            'throughput': vol.get('Throughput'),
                            'encrypted': vol.get('Encrypted'),
                            'snapshot_id': vol.get('SnapshotId'),
                            'kms_key_id': vol.get('KmsKeyId'),
                        }
            except Exception as e:
                print(f"   Volume cache build failed: {_short_err(e)}")

            def _name_from_tags(tags):
                """Return the Name tag value, or None if absent."""
                for t in tags or []:
                    if t.get('Key') == 'Name':
                        return t.get('Value')
                return None

            paginator = ec2.get_paginator('describe_instances')
            _ec2_count = 0
            for page in paginator.paginate():
                for r in page['Reservations']:
                    for i in r['Instances']:
                        _ec2_count += 1
                        instance_id = i['InstanceId']
                        root_device = i.get('RootDeviceName')

                        # Walk block device mappings; separate OS disk from
                        # data disks so we can report per-disk detail.
                        os_disk = {}
                        data_disks = []
                        for bdm in i.get('BlockDeviceMappings', []) or []:
                            if 'Ebs' not in bdm:
                                continue
                            vol_id = bdm['Ebs'].get('VolumeId', '')
                            vol_info = vol_cache.get(vol_id, {})
                            disk = {
                                'device': bdm.get('DeviceName'),
                                'volume_id': vol_id,
                                'size_gb': vol_info.get('size'),
                                'type': vol_info.get('type'),
                                'iops': vol_info.get('iops'),
                                'throughput': vol_info.get('throughput'),
                                'encrypted': vol_info.get('encrypted'),
                                'delete_on_termination': bdm['Ebs'].get('DeleteOnTermination'),
                            }
                            if bdm.get('DeviceName') == root_device:
                                os_disk = disk
                            else:
                                data_disks.append(disk)

                        os_disk_size = os_disk.get('size_gb') or 30
                        data_disk_total = sum((d.get('size_gb') or 0) for d in data_disks)
                        total_storage = os_disk_size + data_disk_total

                        # Render a compact per-data-disk summary for the
                        # Excel report (one line per disk, joined with ;)
                        data_disk_summary = "; ".join(
                            f"{d.get('device','?')}={d.get('volume_id','?')} "
                            f"({d.get('size_gb','?')}GB, {d.get('type','?')}"
                            f"{', encrypted' if d.get('encrypted') else ''})"
                            for d in data_disks
                        ) or "None"

                        # Per-disk Azure managed-disk tier mapping. AWS 25 GB →
                        # Azure 32 GB (P4); 50 GB → 64 GB (P6); 100 GB → 128 GB
                        # (P10). Comma-separated string so each data disk shows
                        # its own rounded target tier in the Excel column.
                        data_disk_azure_sizes = ", ".join(
                            str(_round_to_azure_disk_tier(d.get('size_gb') or 0) or 0)
                            for d in data_disks
                        ) if data_disks else "0"

                        # IAM instance profile (just the name, not full ARN).
                        iam_profile = ''
                        iam = i.get('IamInstanceProfile') or {}
                        if isinstance(iam, dict):
                            iam_profile = (iam.get('Arn') or '').split('/')[-1]

                        # Security group ID list + name list.
                        sg_ids = [sg.get('GroupId', '') for sg in i.get('SecurityGroups', []) if sg.get('GroupId')]
                        sg_names = [sg.get('GroupName', '') for sg in i.get('SecurityGroups', []) if sg.get('GroupName')]

                        name_tag = _name_from_tags(i.get('Tags'))
                        # Use Name tag → fall back to InstanceId.
                        display_name = name_tag or instance_id

                        props = {
                            # Identity
                            'InstanceId': instance_id,
                            'NameTag': name_tag or '',
                            'KeyName': i.get('KeyName', ''),
                            # Compute
                            'InstanceType': i['InstanceType'],
                            'Architecture': i.get('Architecture'),
                            'Platform': i.get('PlatformDetails') or i.get('Platform') or 'Linux/UNIX',
                            'State': i['State']['Name'],
                            'StateReason': (i.get('StateTransitionReason') or '').strip(),
                            'LaunchTime': str(i.get('LaunchTime', '')),
                            'AvailabilityZone': i.get('Placement', {}).get('AvailabilityZone'),
                            'Tenancy': i.get('Placement', {}).get('Tenancy'),
                            # Network
                            'PublicIp': i.get('PublicIpAddress'),
                            'PrivateIp': i.get('PrivateIpAddress'),
                            'PublicDnsName': i.get('PublicDnsName'),
                            'PrivateDnsName': i.get('PrivateDnsName'),
                            'VpcId': i.get('VpcId'),
                            'SubnetId': i.get('SubnetId'),
                            'SecurityGroupIds': ", ".join(sg_ids),
                            'SecurityGroupNames': ", ".join(sg_names),
                            # AMI / virtualization
                            'ImageId': i.get('ImageId'),
                            'Hypervisor': i.get('Hypervisor'),
                            'VirtualizationType': i.get('VirtualizationType'),
                            'RootDeviceType': i.get('RootDeviceType'),
                            # Identity / IAM
                            'IamInstanceProfile': iam_profile,
                            # Performance flags
                            'EbsOptimized': i.get('EbsOptimized'),
                            'EnaSupport': i.get('EnaSupport'),
                            'SriovNetSupport': i.get('SriovNetSupport'),
                            'Monitoring': i.get('Monitoring', {}).get('State'),
                            # Storage — OS + data
                            'OSDiskDevice': os_disk.get('device') or root_device,
                            'OSDiskVolumeId': os_disk.get('volume_id'),
                            'OSDiskSize(GB)': os_disk_size,
                            'OSDiskType': os_disk.get('type'),
                            'OSDiskEncrypted': os_disk.get('encrypted'),
                            'DataDiskCount': len(data_disks),
                            'DataDiskTotalSize(GB)': data_disk_total,
                            'DataDisks': data_disk_summary,
                            # Per-disk Azure-tier rounded sizes for migration
                            # planning. Consumed by the report's "Azure Data
                            # Disk Size" column so each AWS data disk maps to
                            # its own Azure managed-disk SKU.
                            'AzureDataDiskSizes': data_disk_azure_sizes,
                            'TotalStorage(GB)': total_storage,
                        }
                        upsert_resource(
                            arn=f"arn:aws:ec2:{self.region}:{_ec2_account_id}:instance/{instance_id}",
                            name=display_name,
                            type_norm='AWS::EC2::Instance',
                            group='Compute',
                            details_dict=props,
                            status=i['State']['Name'],
                            tags_list=i.get('Tags', [])
                        )
            print(f"   Native EC2 Instances (enriched: Name tag, OS+data disks, AMI, IAM, SG, AZ): {_ec2_count}")
        except Exception as e: print(f"Native EC2 Instances failed: {_short_err(e)}")

        # H3. EC2 Elastic IPs (native — Cloud Control's AWS::EC2::EIP enumeration
        # is unreliable and frequently returns a partial list. describe_addresses
        # is a single, authoritative call that returns every EIP in the region.)
        try:
            ec2 = self._client('ec2')
            account_id = self._client('sts').get_caller_identity()['Account']
            resp = ec2.describe_addresses()
            eip_count = 0
            for addr in resp.get('Addresses', []):
                alloc_id = addr.get('AllocationId') or addr.get('PublicIp')
                public_ip = addr.get('PublicIp')
                associated = bool(addr.get('AssociationId') or addr.get('InstanceId') or addr.get('NetworkInterfaceId'))

                attached_to = ''
                if addr.get('InstanceId'):
                    attached_to = f"EC2 Instance: {addr['InstanceId']}"
                elif addr.get('NetworkInterfaceId'):
                    nic_owner = addr.get('NetworkInterfaceOwnerId', account_id)
                    attached_to = f"Network Interface: {addr['NetworkInterfaceId']} (owner {nic_owner})"
                elif addr.get('NatGatewayId'):
                    attached_to = f"NAT Gateway: {addr['NatGatewayId']}"

                props = {
                    'PublicIp': public_ip,
                    'AllocationId': addr.get('AllocationId'),
                    'Domain': addr.get('Domain'),
                    'AssociationId': addr.get('AssociationId'),
                    'InstanceId': addr.get('InstanceId'),
                    'NetworkInterfaceId': addr.get('NetworkInterfaceId'),
                    'PrivateIpAddress': addr.get('PrivateIpAddress'),
                    'NetworkBorderGroup': addr.get('NetworkBorderGroup'),
                    'PublicIpv4Pool': addr.get('PublicIpv4Pool'),
                    'Attached': 'Yes' if associated else 'No',
                    'Attached To': attached_to or 'None',
                }
                # AWS canonical EIP ARN form
                eip_arn = f"arn:aws:ec2:{self.region}:{account_id}:elastic-ip/{alloc_id}"
                upsert_resource(
                    arn=eip_arn,
                    name=public_ip or alloc_id,
                    type_norm='AWS::EC2::EIP',
                    group='Networking',
                    details_dict=props,
                    status='Active' if associated else 'Unassociated',
                    tags_list=addr.get('Tags', [])
                )
                eip_count += 1
            print(f"   Native EIPs: discovered {eip_count} elastic IPs in {self.region}.")
        except Exception as e:
            print(f"Native EIPs failed: {_short_err(e)}")

        try:
            asg_client = self._client('autoscaling')
            paginator = asg_client.get_paginator('describe_auto_scaling_groups')
            for page in paginator.paginate():
                for g in page['AutoScalingGroups']:
                    props = {
                        'MinSize': g['MinSize'],
                        'MaxSize': g['MaxSize'],
                        'DesiredCapacity': g['DesiredCapacity'],
                        'LaunchConfigurationName': g.get('LaunchConfigurationName'),
                    }
                    if 'LaunchTemplate' in g:
                        props['LaunchTemplate'] = g['LaunchTemplate'].get('LaunchTemplateName')
                        
                    upsert_resource(
                        arn=g['AutoScalingGroupARN'],
                        name=g['AutoScalingGroupName'],
                        type_norm='AWS::AutoScaling::AutoScalingGroup',
                        group='Compute',
                        details_dict=props,
                        status='Active',
                        tags_list=g.get('Tags', [])
                    )
        except Exception as e: print(f"Native ASG failed: {_short_err(e)}")

        # J. DynamoDB
        try:
            ddb = self._client('dynamodb')
            # List tables then describe each (no paginator for describe, but list handles names)
            paginator = ddb.get_paginator('list_tables')
            for page in paginator.paginate():
                for t_name in page['TableNames']:
                    try:
                        desc = ddb.describe_table(TableName=t_name)['Table']
                        props = {
                            'TableSizeBytes': desc.get('TableSizeBytes'),
                            'ItemCount': desc.get('ItemCount'),
                            'BillingMode': desc.get('BillingModeSummary', {}).get('BillingMode', 'PROVISIONED'),
                        }
                        if 'ProvisionedThroughput' in desc:
                             props['RCU'] = desc['ProvisionedThroughput'].get('ReadCapacityUnits')
                             props['WCU'] = desc['ProvisionedThroughput'].get('WriteCapacityUnits')
                        
                        upsert_resource(
                            arn=desc['TableArn'],
                            name=t_name,
                            type_norm='AWS::DynamoDB::Table',
                            group='Database',
                            details_dict=props,
                            status=desc['TableStatus'],
                            # Tags need describe_tags call, skipping for speed unless needed
                        )
                    except Exception as e: print(f"Native DynamoDB Describe failed for {t_name}: {e}")
        except Exception as e: print(f"Native DynamoDB failed: {_short_err(e)}")

        # K. ElastiCache
        try:
            elasticache = self._client('elasticache')
            paginator = elasticache.get_paginator('describe_cache_clusters')
            for page in paginator.paginate():
                for c in page['CacheClusters']:
                    props = {
                        'Engine': c['Engine'],
                        'EngineVersion': c.get('EngineVersion'),
                        'CacheNodeType': c.get('CacheNodeType'),
                        'NumCacheNodes': c.get('NumCacheNodes'),
                        'Status': c['CacheClusterStatus']
                    }
                    upsert_resource(
                        arn=c['ARN'],
                        name=c['CacheClusterId'],
                        type_norm='AWS::ElastiCache::CacheCluster',
                        group='Database',
                        details_dict=props,
                        status=c['CacheClusterStatus']
                    )
        except Exception as e: print(f"Native ElastiCache failed: {_short_err(e)}")

        # L. Redshift Clusters
        try:
            redshift = self._client('redshift')
            paginator = redshift.get_paginator('describe_clusters')
            for page in paginator.paginate():
                for c in page['Clusters']:
                    props = {
                        'NodeType': c['NodeType'],
                        'NumberOfNodes': c['NumberOfNodes'],
                        'ClusterStatus': c['ClusterStatus'],
                        'VpcId': c.get('VpcId')
                    }
                    upsert_resource(
                        arn=f"arn:aws:redshift:{self.region}:{self._client('sts').get_caller_identity()['Account']}:cluster:{c['ClusterIdentifier']}", 
                        name=c['ClusterIdentifier'],
                        type_norm='AWS::Redshift::Cluster',
                        group='Database',
                        details_dict=props,
                        status=c['ClusterStatus']
                    )
        except Exception as e: print(f"Native Redshift failed: {_short_err(e)}")

        # M. OpenSearch Domains
        try:
            opensearch = self._client('opensearch')
            # list_domain_names -> describe_domain (Batch?)
            # describe_domains accepts list.
            resp = opensearch.list_domain_names()
            d_names = [d['DomainName'] for d in resp.get('DomainNames', [])]
            
            # Batch describe (max 5) - simple loop is fine for now
            if d_names:
                 desc_resp = opensearch.describe_domains(DomainNames=d_names)
                 for d in desc_resp['DomainStatusList']:
                     cluster_conf = d.get('ClusterConfig', {})
                     props = {
                         'InstanceType': cluster_conf.get('InstanceType'),
                         'InstanceCount': cluster_conf.get('InstanceCount'),
                         'DedicatedMasterEnabled': cluster_conf.get('DedicatedMasterEnabled'),
                         'EngineVersion': d.get('EngineVersion')
                     }
                     upsert_resource(
                         arn=d['ARN'],
                         name=d['DomainName'],
                         type_norm='AWS::OpenSearchService::Domain',
                         group='Analytics',
                         details_dict=props,
                         status=str(d.get('Processing', False))
                     )
        except Exception as e: print(f"Native OpenSearch failed: {_short_err(e)}")

        # N. SSM Documents (Custom ones only)
        try:
            ssm = self._client('ssm')
            # Fetch Account ID once for ARNs
            account_id = self._client('sts').get_caller_identity()['Account']
            
            paginator = ssm.get_paginator('list_documents')
            # 'Self' filter ensures we only get custom documents
            for page in paginator.paginate(DocumentFilterList=[{'key': 'Owner', 'value': 'Self'}]):
                for doc in page['DocumentIdentifiers']:
                    doc_name = doc['Name']
                    upsert_resource(
                        arn=f"arn:aws:ssm:{self.region}:{account_id}:document/{doc_name}",
                        name=doc_name,
                        type_norm='AWS::SSM::Document',
                        group='Management',
                        details_dict={'DocumentType': doc.get('DocumentType'), 'Owner': doc.get('Owner')},
                        status='Active'
                    )
        except Exception as e: print(f"Native SSM failed: {_short_err(e)}")

        # O. VPCs and Subnets (Ensure CIDRs are ALWAYS captured)
        try:
            ec2 = self._client('ec2')
            account_id = self._client('sts').get_caller_identity()['Account']
            
            # 1. VPCs
            vpcs = ec2.describe_vpcs()['Vpcs']
            for vpc in vpcs:
                vpc_id = vpc['VpcId']
                props = {
                    'CidrBlock': vpc.get('CidrBlock'),
                    'InstanceTenancy': vpc.get('InstanceTenancy'),
                    'State': vpc.get('State'),
                    'IsDefault': vpc.get('IsDefault'),
                    'CidrBlockAssociationSet': [c.get('CidrBlock') for c in vpc.get('CidrBlockAssociationSet', [])]
                }
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:vpc/{vpc_id}",
                    name=_name_or_id(vpc, vpc_id),
                    type_norm='AWS::EC2::VPC',
                    group='Networking',
                    details_dict=props,
                    status=vpc.get('State', 'Active'),
                    tags_list=vpc.get('Tags', [])
                )

            # 2. Subnets
            subnets = ec2.describe_subnets()['Subnets']
            for sub in subnets:
                sub_id = sub['SubnetId']
                props = {
                    'VpcId': sub.get('VpcId'),
                    'CidrBlock': sub.get('CidrBlock'),
                    'AvailabilityZone': sub.get('AvailabilityZone'),
                    'AvailableIpAddressCount': sub.get('AvailableIpAddressCount'),
                    'State': sub.get('State'),
                    'MapPublicIpOnLaunch': sub.get('MapPublicIpOnLaunch')
                }
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:subnet/{sub_id}",
                    name=_name_or_id(sub, sub_id),
                    type_norm='AWS::EC2::Subnet',
                    group='Networking',
                    details_dict=props,
                    status=sub.get('State', 'Active'),
                    tags_list=sub.get('Tags', [])
                )

            # 3. Security Groups
            def _explode_sg_rule(rule):
                """Yield one formatted line per (proto+port, source) pair.

                AWS Console shows EACH source (CIDR, prefix list, peer SG)
                as its OWN row even when they share the same protocol+port.
                We mirror that: one IpPermissions entry with N IpRanges
                produces N output lines, so the SG sheet row count matches
                the Console exactly. Earlier this collapsed to 1 line, which
                undercounted rules (e.g., 20 Console rows → 12 sheet rows).
                """
                proto = str(rule.get('IpProtocol', '-1'))
                if proto == '-1':
                    proto_label, port_label = 'ALL', 'ALL'
                else:
                    proto_label = proto.upper()
                    fp, tp = rule.get('FromPort'), rule.get('ToPort')
                    if fp is None and tp is None:
                        port_label = 'ALL'
                    elif fp == tp:
                        port_label = str(fp)
                    else:
                        port_label = f"{fp}-{tp}"

                lines = []
                for ipr in rule.get('IpRanges', []) or []:
                    src = ipr.get('CidrIp', '')
                    desc = ipr.get('Description', '')
                    lines.append(f"{proto_label}:{port_label} from {src} | {desc}".rstrip(' |'))
                for ipr6 in rule.get('Ipv6Ranges', []) or []:
                    src = ipr6.get('CidrIpv6', '')
                    desc = ipr6.get('Description', '')
                    lines.append(f"{proto_label}:{port_label} from {src} | {desc}".rstrip(' |'))
                for pl in rule.get('PrefixListIds', []) or []:
                    src = pl.get('PrefixListId', '')
                    desc = pl.get('Description', '')
                    lines.append(f"{proto_label}:{port_label} from {src} | {desc}".rstrip(' |'))
                for ug in rule.get('UserIdGroupPairs', []) or []:
                    src = ug.get('GroupId') or ug.get('GroupName') or ''
                    desc = ug.get('Description', '')
                    lines.append(f"{proto_label}:{port_label} from {src} | {desc}".rstrip(' |'))
                # No sources of any kind — emit one row so the rule still
                # appears (rare, but possible for malformed rules).
                if not lines:
                    lines.append(f"{proto_label}:{port_label} from any | ")
                return lines

            sgs = ec2.describe_security_groups()['SecurityGroups']
            for sg in sgs:
                sg_id = sg['GroupId']
                sg_name = sg.get('GroupName', sg_id)

                ingress_lines = []
                for r in sg.get('IpPermissions', []) or []:
                    ingress_lines.extend(_explode_sg_rule(r))
                egress_lines = []
                for r in sg.get('IpPermissionsEgress', []) or []:
                    egress_lines.extend(_explode_sg_rule(r))

                # A. The Group itself
                props = {
                    'GroupName': sg_name,
                    'Description': sg.get('Description'),
                    'VpcId': sg.get('VpcId'),
                    'IngressRulesCount': len(ingress_lines),
                    'EgressRulesCount': len(egress_lines),
                    # Populate text columns the rule-explosion code reads.
                    # Empty string (NOT NaN) when no rules — prevents the
                    # bug where pandas NaN gets stringified to 'nan' and
                    # treated as a fake rule by the explosion code.
                    'Ingress Rules': '\n'.join(ingress_lines),
                    'Egress Rules': '\n'.join(egress_lines),
                }
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:security-group/{sg_id}",
                    name=sg_name,
                    type_norm='AWS::EC2::SecurityGroup',
                    group='Network Security',
                    details_dict=props,
                    status='Active',
                    tags_list=sg.get('Tags', [])
                )
                
                # B. Ingress Rules (Individual Rows)
                for i, rule in enumerate(sg.get('IpPermissions', [])):
                    rule_id = f"{sg_id}-ingress-{i}"
                    rule_props = {
                        'ParentGroup': sg_name,
                        'GroupId': sg_id,
                        'VpcId': sg.get('VpcId'),
                        'Direction': 'Ingress',
                    }
                    # Flatten the rule structure for individual columns
                    rule_props.update(rule)
                    
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:security-group-rule/{rule_id}",
                        name=f"{sg_name} In-Rule {i}",
                        type_norm='AWS::EC2::SecurityGroupIngress',
                        group='Network Security',
                        details_dict=rule_props,
                        status='Active'
                    )
                    
                # C. Egress Rules (Individual Rows)
                for i, rule in enumerate(sg.get('IpPermissionsEgress', [])):
                    rule_id = f"{sg_id}-egress-{i}"
                    rule_props = {
                        'ParentGroup': sg_name,
                        'GroupId': sg_id,
                        'VpcId': sg.get('VpcId'),
                        'Direction': 'Egress',
                    }
                    rule_props.update(rule)
                    
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:security-group-rule/{rule_id}",
                        name=f"{sg_name} Out-Rule {i}",
                        type_norm='AWS::EC2::SecurityGroupEgress',
                        group='Network Security',
                        details_dict=rule_props,
                        status='Active'
                    )

            # 4. Network Interfaces
            enis = ec2.describe_network_interfaces()['NetworkInterfaces']
            for eni in enis:
                eni_id = eni['NetworkInterfaceId']
                props = {
                    'SubnetId': eni.get('SubnetId'),
                    'VpcId': eni.get('VpcId'),
                    'PrivateIpAddress': eni.get('PrivateIpAddress'),
                    'InterfaceType': eni.get('InterfaceType'),
                    'Status': eni.get('Status'),
                    'InstanceId': eni.get('Attachment', {}).get('InstanceId') if eni.get('Attachment') else None
                }
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:network-interface/{eni_id}",
                    name=_name_or_id(eni, eni_id),
                    type_norm='AWS::EC2::NetworkInterface',
                    group='Networking',
                    details_dict=props,
                    status=eni.get('Status', 'Active'),
                    tags_list=eni.get('Tags', [])
                )

        except Exception as e: print(f"Native Networking (VPC/Subnet/SG/ENI) failed: {_short_err(e)}")

        # =====================================================================
        # COMPREHENSIVE NATIVE COVERAGE — authoritative describe_* / list_* calls
        # for every commonly-undercounted type. Each block is independent: a
        # single permission deny or service-not-available failure only skips
        # that one block. Counts are printed for visibility.
        # =====================================================================

        def _name_from_tags(tags, fallback):
            if not tags:
                return fallback
            for t in tags:
                if isinstance(t, dict) and t.get('Key', '').lower() == 'name':
                    return t.get('Value') or fallback
            return fallback

        try:
            account_id = self._client('sts').get_caller_identity()['Account']
        except Exception:
            account_id = '000000000000'

        # ---- EC2: NAT Gateways ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_nat_gateways').paginate():
                for ng in page['NatGateways']:
                    ng_id = ng['NatGatewayId']
                    addrs = ng.get('NatGatewayAddresses', [])
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:natgateway/{ng_id}",
                        name=_name_from_tags(ng.get('Tags'), ng_id),
                        type_norm='AWS::EC2::NatGateway',
                        group='Networking',
                        details_dict={
                            'VpcId': ng.get('VpcId'),
                            'SubnetId': ng.get('SubnetId'),
                            'State': ng.get('State'),
                            'ConnectivityType': ng.get('ConnectivityType'),
                            'PublicIp': ", ".join(a.get('PublicIp', '') for a in addrs if a.get('PublicIp')),
                            'AllocationIds': ", ".join(a.get('AllocationId', '') for a in addrs if a.get('AllocationId')),
                        },
                        status=ng.get('State', 'Active'),
                        tags_list=ng.get('Tags', []),
                    )
                    count += 1
            print(f"   Native NAT Gateways: {count}")
        except Exception as e:
            print(f"Native NAT Gateways failed: {_short_err(e)}")

        # ---- EC2: Internet Gateways ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_internet_gateways').paginate():
                for igw in page['InternetGateways']:
                    igw_id = igw['InternetGatewayId']
                    attachments = igw.get('Attachments', [])
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:internet-gateway/{igw_id}",
                        name=_name_from_tags(igw.get('Tags'), igw_id),
                        type_norm='AWS::EC2::InternetGateway',
                        group='Networking',
                        details_dict={
                            'AttachedVpcs': ", ".join(a.get('VpcId', '') for a in attachments),
                            'State': attachments[0].get('State') if attachments else 'detached',
                        },
                        status='Active',
                        tags_list=igw.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Internet Gateways: {count}")
        except Exception as e:
            print(f"Native Internet Gateways failed: {_short_err(e)}")

        # ---- EC2: Egress-Only Internet Gateways ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_egress_only_internet_gateways').paginate():
                for igw in page['EgressOnlyInternetGateways']:
                    eigw_id = igw['EgressOnlyInternetGatewayId']
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:egress-only-internet-gateway/{eigw_id}",
                        name=_name_from_tags(igw.get('Tags'), eigw_id),
                        type_norm='AWS::EC2::EgressOnlyInternetGateway',
                        group='Networking',
                        details_dict={'Attachments': str(igw.get('Attachments', []))},
                        status='Active',
                        tags_list=igw.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Egress-Only IGWs: {count}")
        except Exception as e:
            print(f"Native Egress-Only IGWs failed: {_short_err(e)}")

        # ---- EC2: Route Tables ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_route_tables').paginate():
                for rt in page['RouteTables']:
                    rt_id = rt['RouteTableId']
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:route-table/{rt_id}",
                        name=_name_from_tags(rt.get('Tags'), rt_id),
                        type_norm='AWS::EC2::RouteTable',
                        group='Networking',
                        details_dict={
                            'VpcId': rt.get('VpcId'),
                            'RouteCount': len(rt.get('Routes', [])),
                            'AssociationCount': len(rt.get('Associations', [])),
                            'Main': any(a.get('Main') for a in rt.get('Associations', [])),
                        },
                        status='Active',
                        tags_list=rt.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Route Tables: {count}")
        except Exception as e:
            print(f"Native Route Tables failed: {_short_err(e)}")

        # ---- EC2: Network ACLs ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_network_acls').paginate():
                for acl in page['NetworkAcls']:
                    acl_id = acl['NetworkAclId']
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:network-acl/{acl_id}",
                        name=_name_from_tags(acl.get('Tags'), acl_id),
                        type_norm='AWS::EC2::NetworkAcl',
                        group='Networking',
                        details_dict={
                            'VpcId': acl.get('VpcId'),
                            'IsDefault': acl.get('IsDefault'),
                            'EntryCount': len(acl.get('Entries', [])),
                            'AssociationCount': len(acl.get('Associations', [])),
                        },
                        status='Active',
                        tags_list=acl.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Network ACLs: {count}")
        except Exception as e:
            print(f"Native Network ACLs failed: {_short_err(e)}")

        # ---- EC2: VPC Endpoints ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_vpc_endpoints').paginate():
                for ep in page['VpcEndpoints']:
                    ep_id = ep['VpcEndpointId']
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:vpc-endpoint/{ep_id}",
                        name=_name_from_tags(ep.get('Tags'), ep_id),
                        type_norm='AWS::EC2::VPCEndpoint',
                        group='Networking',
                        details_dict={
                            'VpcId': ep.get('VpcId'),
                            'ServiceName': ep.get('ServiceName'),
                            'VpcEndpointType': ep.get('VpcEndpointType'),
                            'State': ep.get('State'),
                            'PrivateDnsEnabled': ep.get('PrivateDnsEnabled'),
                        },
                        status=ep.get('State', 'Active'),
                        tags_list=ep.get('Tags', []),
                    )
                    count += 1
            print(f"   Native VPC Endpoints: {count}")
        except Exception as e:
            print(f"Native VPC Endpoints failed: {_short_err(e)}")

        # ---- EC2: Transit Gateways + Attachments ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_transit_gateways').paginate():
                for tgw in page['TransitGateways']:
                    tgw_id = tgw['TransitGatewayId']
                    upsert_resource(
                        arn=tgw.get('TransitGatewayArn') or f"arn:aws:ec2:{self.region}:{account_id}:transit-gateway/{tgw_id}",
                        name=_name_from_tags(tgw.get('Tags'), tgw_id),
                        type_norm='AWS::EC2::TransitGateway',
                        group='Networking',
                        details_dict={
                            'State': tgw.get('State'),
                            'OwnerId': tgw.get('OwnerId'),
                            'Description': tgw.get('Description'),
                            'CreationTime': str(tgw.get('CreationTime', '')),
                        },
                        status=tgw.get('State', 'Active'),
                        tags_list=tgw.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Transit Gateways: {count}")
        except Exception as e:
            print(f"Native Transit Gateways failed: {_short_err(e)}")

        # ---- EC2: VPC Peering Connections ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_vpc_peering_connections').paginate():
                for pc in page['VpcPeeringConnections']:
                    pc_id = pc['VpcPeeringConnectionId']
                    a, b = pc.get('AccepterVpcInfo', {}), pc.get('RequesterVpcInfo', {})
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}:{account_id}:vpc-peering-connection/{pc_id}",
                        name=_name_from_tags(pc.get('Tags'), pc_id),
                        type_norm='AWS::EC2::VPCPeeringConnection',
                        group='Networking',
                        details_dict={
                            'Status': pc.get('Status', {}).get('Code'),
                            'AccepterVpc': a.get('VpcId'),
                            'AccepterAccount': a.get('OwnerId'),
                            'AccepterRegion': a.get('Region'),
                            'RequesterVpc': b.get('VpcId'),
                            'RequesterAccount': b.get('OwnerId'),
                            'RequesterRegion': b.get('Region'),
                        },
                        status=pc.get('Status', {}).get('Code', 'active'),
                        tags_list=pc.get('Tags', []),
                    )
                    count += 1
            print(f"   Native VPC Peering Connections: {count}")
        except Exception as e:
            print(f"Native VPC Peering Connections failed: {_short_err(e)}")

        # ---- EC2: EBS Snapshots (owned by this account) ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for page in ec2.get_paginator('describe_snapshots').paginate(OwnerIds=['self']):
                for snap in page['Snapshots']:
                    sid = snap['SnapshotId']
                    upsert_resource(
                        arn=f"arn:aws:ec2:{self.region}::snapshot/{sid}",
                        name=_name_from_tags(snap.get('Tags'), sid),
                        type_norm='AWS::EC2::Snapshot',
                        group='Storage',
                        details_dict={
                            'VolumeId': snap.get('VolumeId'),
                            'VolumeSize': snap.get('VolumeSize'),
                            'State': snap.get('State'),
                            'Encrypted': snap.get('Encrypted'),
                            'Progress': snap.get('Progress'),
                            'StartTime': str(snap.get('StartTime', '')),
                        },
                        status=snap.get('State', 'completed'),
                        tags_list=snap.get('Tags', []),
                    )
                    count += 1
            print(f"   Native EBS Snapshots (self-owned): {count}")
        except Exception as e:
            print(f"Native EBS Snapshots failed: {_short_err(e)}")

        # ---- EC2: AMIs (owned by this account) ----
        try:
            ec2 = self._client('ec2')
            count = 0
            resp = ec2.describe_images(Owners=['self'])
            for img in resp.get('Images', []):
                img_id = img['ImageId']
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}::image/{img_id}",
                    name=img.get('Name') or _name_from_tags(img.get('Tags'), img_id),
                    type_norm='AWS::EC2::Image',
                    group='Compute',
                    details_dict={
                        'Architecture': img.get('Architecture'),
                        'Platform': img.get('PlatformDetails'),
                        'RootDeviceType': img.get('RootDeviceType'),
                        'VirtualizationType': img.get('VirtualizationType'),
                        'State': img.get('State'),
                        'CreationDate': img.get('CreationDate'),
                        'Public': img.get('Public'),
                    },
                    status=img.get('State', 'available'),
                    tags_list=img.get('Tags', []),
                )
                count += 1
            print(f"   Native AMIs (self-owned): {count}")
        except Exception as e:
            print(f"Native AMIs failed: {_short_err(e)}")

        # ---- EC2: VPN Gateways / Customer Gateways / VPN Connections ----
        try:
            ec2 = self._client('ec2')
            count = 0
            for vgw in ec2.describe_vpn_gateways().get('VpnGateways', []):
                vid = vgw['VpnGatewayId']
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:vpn-gateway/{vid}",
                    name=_name_from_tags(vgw.get('Tags'), vid),
                    type_norm='AWS::EC2::VPNGateway',
                    group='Networking',
                    details_dict={
                        'State': vgw.get('State'),
                        'Type': vgw.get('Type'),
                        'VpcAttachments': str(vgw.get('VpcAttachments', [])),
                    },
                    status=vgw.get('State', 'available'),
                    tags_list=vgw.get('Tags', []),
                )
                count += 1
            print(f"   Native VPN Gateways: {count}")
        except Exception as e:
            print(f"Native VPN Gateways failed: {_short_err(e)}")

        try:
            ec2 = self._client('ec2')
            count = 0
            for cgw in ec2.describe_customer_gateways().get('CustomerGateways', []):
                cid = cgw['CustomerGatewayId']
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:customer-gateway/{cid}",
                    name=_name_from_tags(cgw.get('Tags'), cid),
                    type_norm='AWS::EC2::CustomerGateway',
                    group='Networking',
                    details_dict={
                        'State': cgw.get('State'),
                        'Type': cgw.get('Type'),
                        'IpAddress': cgw.get('IpAddress'),
                        'BgpAsn': cgw.get('BgpAsn'),
                    },
                    status=cgw.get('State', 'available'),
                    tags_list=cgw.get('Tags', []),
                )
                count += 1
            print(f"   Native Customer Gateways: {count}")
        except Exception as e:
            print(f"Native Customer Gateways failed: {_short_err(e)}")

        try:
            ec2 = self._client('ec2')
            count = 0
            for vc in ec2.describe_vpn_connections().get('VpnConnections', []):
                vid = vc['VpnConnectionId']
                upsert_resource(
                    arn=f"arn:aws:ec2:{self.region}:{account_id}:vpn-connection/{vid}",
                    name=_name_from_tags(vc.get('Tags'), vid),
                    type_norm='AWS::EC2::VPNConnection',
                    group='Networking',
                    details_dict={
                        'State': vc.get('State'),
                        'Type': vc.get('Type'),
                        'CustomerGatewayId': vc.get('CustomerGatewayId'),
                        'VpnGatewayId': vc.get('VpnGatewayId'),
                        'TransitGatewayId': vc.get('TransitGatewayId'),
                    },
                    status=vc.get('State', 'available'),
                    tags_list=vc.get('Tags', []),
                )
                count += 1
            print(f"   Native VPN Connections: {count}")
        except Exception as e:
            print(f"Native VPN Connections failed: {_short_err(e)}")

        # ---- Classic Load Balancers (ELBv1) ----
        try:
            elb = self._client('elb')
            count = 0
            for page in elb.get_paginator('describe_load_balancers').paginate():
                for lb in page['LoadBalancerDescriptions']:
                    name = lb['LoadBalancerName']
                    upsert_resource(
                        arn=f"arn:aws:elasticloadbalancing:{self.region}:{account_id}:loadbalancer/{name}",
                        name=name,
                        type_norm='AWS::ElasticLoadBalancing::LoadBalancer',
                        group='Networking',
                        details_dict={
                            'Scheme': lb.get('Scheme'),
                            'DNSName': lb.get('DNSName'),
                            'VPCId': lb.get('VPCId'),
                            'AvailabilityZones': ", ".join(lb.get('AvailabilityZones', [])),
                            'InstanceCount': len(lb.get('Instances', [])),
                            'CreatedTime': str(lb.get('CreatedTime', '')),
                        },
                        status='Active',
                    )
                    count += 1
            print(f"   Native Classic ELBs: {count}")
        except Exception as e:
            print(f"Native Classic ELBs failed: {_short_err(e)}")

        # ---- EFS File Systems ----
        try:
            efs = self._client('efs')
            count = 0
            for page in efs.get_paginator('describe_file_systems').paginate():
                for fs in page['FileSystems']:
                    fs_id = fs['FileSystemId']
                    upsert_resource(
                        arn=fs.get('FileSystemArn') or f"arn:aws:elasticfilesystem:{self.region}:{account_id}:file-system/{fs_id}",
                        name=fs.get('Name') or fs_id,
                        type_norm='AWS::EFS::FileSystem',
                        group='Storage',
                        details_dict={
                            'LifeCycleState': fs.get('LifeCycleState'),
                            'PerformanceMode': fs.get('PerformanceMode'),
                            'ThroughputMode': fs.get('ThroughputMode'),
                            'Encrypted': fs.get('Encrypted'),
                            'SizeInBytes': fs.get('SizeInBytes', {}).get('Value'),
                            'NumberOfMountTargets': fs.get('NumberOfMountTargets'),
                        },
                        status=fs.get('LifeCycleState', 'available'),
                        tags_list=fs.get('Tags', []),
                    )
                    count += 1
            print(f"   Native EFS File Systems: {count}")
        except Exception as e:
            print(f"Native EFS failed: {_short_err(e)}")

        # ---- Route 53 Hosted Zones (global service) ----
        try:
            r53 = self._client('route53')
            count = 0
            for page in r53.get_paginator('list_hosted_zones').paginate():
                for hz in page['HostedZones']:
                    zid = hz['Id'].split('/')[-1]
                    upsert_resource(
                        arn=f"arn:aws:route53:::hostedzone/{zid}",
                        name=hz['Name'].rstrip('.'),
                        type_norm='AWS::Route53::HostedZone',
                        group='Networking',
                        details_dict={
                            'PrivateZone': hz.get('Config', {}).get('PrivateZone'),
                            'RecordSetCount': hz.get('ResourceRecordSetCount'),
                            'Comment': hz.get('Config', {}).get('Comment'),
                        },
                        status='Active',
                    )
                    count += 1
            print(f"   Native Route 53 Hosted Zones: {count}")
        except Exception as e:
            print(f"Native Route 53 failed: {_short_err(e)}")

        # ---- CloudFront Distributions (global service) ----
        try:
            cf = self._client('cloudfront')
            count = 0
            for page in cf.get_paginator('list_distributions').paginate():
                dl = page.get('DistributionList', {})
                for d in dl.get('Items', []) or []:
                    did = d['Id']
                    upsert_resource(
                        arn=d.get('ARN') or f"arn:aws:cloudfront::{account_id}:distribution/{did}",
                        name=d.get('DomainName') or did,
                        type_norm='AWS::CloudFront::Distribution',
                        group='Networking',
                        details_dict={
                            'Enabled': d.get('Enabled'),
                            'Status': d.get('Status'),
                            'DomainName': d.get('DomainName'),
                            'Comment': d.get('Comment'),
                            'PriceClass': d.get('PriceClass'),
                            'OriginCount': d.get('Origins', {}).get('Quantity', 0),
                        },
                        status=d.get('Status', 'Deployed'),
                    )
                    count += 1
            print(f"   Native CloudFront Distributions: {count}")
        except Exception as e:
            print(f"Native CloudFront failed: {_short_err(e)}")

        # ---- SNS Topics ----
        try:
            sns = self._client('sns')
            count = 0
            for page in sns.get_paginator('list_topics').paginate():
                for t in page['Topics']:
                    arn = t['TopicArn']
                    name = arn.split(':')[-1]
                    upsert_resource(
                        arn=arn,
                        name=name,
                        type_norm='AWS::SNS::Topic',
                        group='Messaging',
                        details_dict={'TopicArn': arn},
                        status='Active',
                    )
                    count += 1
            print(f"   Native SNS Topics: {count}")
        except Exception as e:
            print(f"Native SNS failed: {_short_err(e)}")

        # ---- SQS Queues ----
        try:
            sqs = self._client('sqs')
            count = 0
            urls = []
            paginator = sqs.get_paginator('list_queues')
            for page in paginator.paginate():
                urls.extend(page.get('QueueUrls', []) or [])
            for url in urls:
                name = url.rsplit('/', 1)[-1]
                arn = f"arn:aws:sqs:{self.region}:{account_id}:{name}"
                try:
                    attrs = sqs.get_queue_attributes(QueueUrl=url, AttributeNames=['QueueArn', 'CreatedTimestamp', 'ApproximateNumberOfMessages']).get('Attributes', {})
                    arn = attrs.get('QueueArn', arn)
                    details = {
                        'QueueArn': arn,
                        'Messages': attrs.get('ApproximateNumberOfMessages'),
                        'Url': url,
                        'IsFIFO': name.endswith('.fifo'),
                    }
                except Exception:
                    details = {'Url': url, 'IsFIFO': name.endswith('.fifo')}
                upsert_resource(arn=arn, name=name, type_norm='AWS::SQS::Queue',
                                group='Messaging', details_dict=details, status='Active')
                count += 1
            print(f"   Native SQS Queues: {count}")
        except Exception as e:
            print(f"Native SQS failed: {_short_err(e)}")

        # ---- EventBridge Rules ----
        try:
            events = self._client('events')
            count = 0
            buses = []
            try:
                for page in events.get_paginator('list_event_buses').paginate():
                    buses.extend([b['Name'] for b in page.get('EventBuses', [])])
            except Exception:
                buses = ['default']
            for bus in buses or ['default']:
                try:
                    resp = events.list_rules(EventBusName=bus, Limit=100)
                    for r in resp.get('Rules', []):
                        upsert_resource(
                            arn=r.get('Arn'),
                            name=r.get('Name'),
                            type_norm='AWS::Events::Rule',
                            group='Integration',
                            details_dict={
                                'EventBus': bus,
                                'State': r.get('State'),
                                'ScheduleExpression': r.get('ScheduleExpression'),
                                'EventPattern': (r.get('EventPattern') or '')[:200],
                            },
                            status=r.get('State', 'ENABLED'),
                        )
                        count += 1
                except Exception:
                    continue
            print(f"   Native EventBridge Rules: {count}")
        except Exception as e:
            print(f"Native EventBridge failed: {_short_err(e)}")

        # ---- Step Functions State Machines ----
        try:
            sfn = self._client('stepfunctions')
            count = 0
            for page in sfn.get_paginator('list_state_machines').paginate():
                for sm in page['stateMachines']:
                    upsert_resource(
                        arn=sm['stateMachineArn'],
                        name=sm['name'],
                        type_norm='AWS::StepFunctions::StateMachine',
                        group='Integration',
                        details_dict={
                            'Type': sm.get('type'),
                            'CreationDate': str(sm.get('creationDate', '')),
                        },
                        status='Active',
                    )
                    count += 1
            print(f"   Native Step Functions: {count}")
        except Exception as e:
            print(f"Native Step Functions failed: {_short_err(e)}")

        # ---- KMS Customer-Managed Keys ----
        try:
            kms = self._client('kms')
            count = 0
            for page in kms.get_paginator('list_keys').paginate():
                for k in page['Keys']:
                    try:
                        meta = kms.describe_key(KeyId=k['KeyId'])['KeyMetadata']
                        if meta.get('KeyManager') != 'CUSTOMER':
                            continue
                        upsert_resource(
                            arn=meta['Arn'],
                            name=meta.get('Description') or meta['KeyId'],
                            type_norm='AWS::KMS::Key',
                            group='Security',
                            details_dict={
                                'KeyId': meta['KeyId'],
                                'KeyState': meta.get('KeyState'),
                                'KeyUsage': meta.get('KeyUsage'),
                                'KeySpec': meta.get('KeySpec'),
                                'Origin': meta.get('Origin'),
                                'MultiRegion': meta.get('MultiRegion'),
                            },
                            status=meta.get('KeyState', 'Enabled'),
                        )
                        count += 1
                    except Exception:
                        continue
            print(f"   Native KMS Keys (customer-managed): {count}")
        except Exception as e:
            print(f"Native KMS failed: {_short_err(e)}")

        # ---- Secrets Manager ----
        try:
            sm = self._client('secretsmanager')
            count = 0
            for page in sm.get_paginator('list_secrets').paginate():
                for s in page['SecretList']:
                    upsert_resource(
                        arn=s['ARN'],
                        name=s['Name'],
                        type_norm='AWS::SecretsManager::Secret',
                        group='Security',
                        details_dict={
                            'Description': s.get('Description'),
                            'KmsKeyId': s.get('KmsKeyId'),
                            'LastChangedDate': str(s.get('LastChangedDate', '')),
                            'RotationEnabled': s.get('RotationEnabled'),
                        },
                        status='Active',
                        tags_list=s.get('Tags', []),
                    )
                    count += 1
            print(f"   Native Secrets Manager Secrets: {count}")
        except Exception as e:
            print(f"Native Secrets Manager failed: {_short_err(e)}")

        # ---- WAFv2 Web ACLs (Regional + CloudFront/global) ----
        try:
            wafv2 = self._client('wafv2')
            count = 0
            for scope in ('REGIONAL', 'CLOUDFRONT'):
                try:
                    # CLOUDFRONT scope is only valid in us-east-1
                    if scope == 'CLOUDFRONT' and self.region != 'us-east-1':
                        continue
                    resp = wafv2.list_web_acls(Scope=scope, Limit=100)
                    for acl in resp.get('WebACLs', []):
                        upsert_resource(
                            arn=acl['ARN'],
                            name=acl['Name'],
                            type_norm='AWS::WAFv2::WebACL',
                            group='Security',
                            details_dict={'Scope': scope, 'Description': acl.get('Description')},
                            status='Active',
                        )
                        count += 1
                except Exception:
                    continue
            print(f"   Native WAFv2 Web ACLs: {count}")
        except Exception as e:
            print(f"Native WAFv2 failed: {_short_err(e)}")

        # ---- RDS DB Clusters (Aurora / DocumentDB / Neptune) ----
        try:
            rds = self._client('rds')
            count = 0
            for page in rds.get_paginator('describe_db_clusters').paginate():
                for c in page['DBClusters']:
                    upsert_resource(
                        arn=c['DBClusterArn'],
                        name=c['DBClusterIdentifier'],
                        type_norm='AWS::RDS::DBCluster',
                        group='Database',
                        details_dict={
                            'Engine': c.get('Engine'),
                            'EngineVersion': c.get('EngineVersion'),
                            'Status': c.get('Status'),
                            'MultiAZ': c.get('MultiAZ'),
                            'StorageEncrypted': c.get('StorageEncrypted'),
                            'MemberCount': len(c.get('DBClusterMembers', [])),
                            'AllocatedStorage': c.get('AllocatedStorage'),
                        },
                        status=c.get('Status', 'available'),
                        tags_list=c.get('TagList', []),
                    )
                    count += 1
            print(f"   Native RDS DB Clusters: {count}")
        except Exception as e:
            print(f"Native RDS DB Clusters failed: {_short_err(e)}")

        # ---- RDS Manual Snapshots ----
        try:
            rds = self._client('rds')
            count = 0
            for page in rds.get_paginator('describe_db_snapshots').paginate(SnapshotType='manual'):
                for s in page['DBSnapshots']:
                    upsert_resource(
                        arn=s['DBSnapshotArn'],
                        name=s['DBSnapshotIdentifier'],
                        type_norm='AWS::RDS::DBSnapshot',
                        group='Database',
                        details_dict={
                            'Engine': s.get('Engine'),
                            'Status': s.get('Status'),
                            'DBInstanceIdentifier': s.get('DBInstanceIdentifier'),
                            'AllocatedStorage': s.get('AllocatedStorage'),
                            'SnapshotCreateTime': str(s.get('SnapshotCreateTime', '')),
                        },
                        status=s.get('Status', 'available'),
                        tags_list=s.get('TagList', []),
                    )
                    count += 1
            print(f"   Native RDS Manual Snapshots: {count}")
        except Exception as e:
            print(f"Native RDS Snapshots failed: {_short_err(e)}")

        # ---- ECS Clusters + Services ----
        try:
            ecs = self._client('ecs')
            count_c = count_s = 0
            for page in ecs.get_paginator('list_clusters').paginate():
                arns = page.get('clusterArns', [])
                if not arns:
                    continue
                for batch_start in range(0, len(arns), 10):
                    batch = arns[batch_start:batch_start + 10]
                    for c in ecs.describe_clusters(clusters=batch).get('clusters', []):
                        upsert_resource(
                            arn=c['clusterArn'],
                            name=c['clusterName'],
                            type_norm='AWS::ECS::Cluster',
                            group='Container',
                            details_dict={
                                'Status': c.get('status'),
                                'RegisteredContainerInstancesCount': c.get('registeredContainerInstancesCount'),
                                'RunningTasksCount': c.get('runningTasksCount'),
                                'PendingTasksCount': c.get('pendingTasksCount'),
                                'ActiveServicesCount': c.get('activeServicesCount'),
                            },
                            status=c.get('status', 'ACTIVE'),
                            tags_list=c.get('tags', []),
                        )
                        count_c += 1
                        try:
                            svc_arns = []
                            for sp in ecs.get_paginator('list_services').paginate(cluster=c['clusterArn']):
                                svc_arns.extend(sp.get('serviceArns', []))
                            for sbs in range(0, len(svc_arns), 10):
                                sb = svc_arns[sbs:sbs + 10]
                                if not sb:
                                    continue
                                for s in ecs.describe_services(cluster=c['clusterArn'], services=sb).get('services', []):
                                    upsert_resource(
                                        arn=s['serviceArn'],
                                        name=s['serviceName'],
                                        type_norm='AWS::ECS::Service',
                                        group='Container',
                                        details_dict={
                                            'ClusterArn': c['clusterArn'],
                                            'LaunchType': s.get('launchType'),
                                            'TaskDefinition': s.get('taskDefinition'),
                                            'DesiredCount': s.get('desiredCount'),
                                            'RunningCount': s.get('runningCount'),
                                            'Status': s.get('status'),
                                        },
                                        status=s.get('status', 'ACTIVE'),
                                        tags_list=s.get('tags', []),
                                    )
                                    count_s += 1
                        except Exception:
                            pass
            print(f"   Native ECS Clusters: {count_c}, Services: {count_s}")
        except Exception as e:
            print(f"Native ECS failed: {_short_err(e)}")

        # ---- ECR Repositories ----
        try:
            ecr = self._client('ecr')
            count = 0
            for page in ecr.get_paginator('describe_repositories').paginate():
                for r in page['repositories']:
                    upsert_resource(
                        arn=r['repositoryArn'],
                        name=r['repositoryName'],
                        type_norm='AWS::ECR::Repository',
                        group='Container',
                        details_dict={
                            'RepositoryUri': r.get('repositoryUri'),
                            'ImageTagMutability': r.get('imageTagMutability'),
                            'CreatedAt': str(r.get('createdAt', '')),
                        },
                        status='Active',
                    )
                    count += 1
            print(f"   Native ECR Repositories: {count}")
        except Exception as e:
            print(f"Native ECR failed: {_short_err(e)}")

        # ---- EKS Clusters ----
        try:
            eks = self._client('eks')
            count = 0
            for page in eks.get_paginator('list_clusters').paginate():
                for name in page.get('clusters', []):
                    try:
                        c = eks.describe_cluster(name=name)['cluster']
                        upsert_resource(
                            arn=c['arn'],
                            name=c['name'],
                            type_norm='AWS::EKS::Cluster',
                            group='Container',
                            details_dict={
                                'Version': c.get('version'),
                                'Status': c.get('status'),
                                'PlatformVersion': c.get('platformVersion'),
                                'Endpoint': c.get('endpoint'),
                                'RoleArn': c.get('roleArn'),
                            },
                            status=c.get('status', 'ACTIVE'),
                            tags_list=[{'Key': k, 'Value': v} for k, v in (c.get('tags') or {}).items()],
                        )
                        count += 1
                    except Exception:
                        continue
            print(f"   Native EKS Clusters: {count}")
        except Exception as e:
            print(f"Native EKS failed: {_short_err(e)}")

        # ---- API Gateway REST (v1) ----
        try:
            apigw = self._client('apigateway')
            count = 0
            for page in apigw.get_paginator('get_rest_apis').paginate():
                for api in page.get('items', []):
                    apid = api['id']
                    upsert_resource(
                        arn=f"arn:aws:apigateway:{self.region}::/restapis/{apid}",
                        name=api.get('name'),
                        type_norm='AWS::ApiGateway::RestApi',
                        group='Integration',
                        details_dict={
                            'Id': apid,
                            'Description': api.get('description'),
                            'CreatedDate': str(api.get('createdDate', '')),
                            'EndpointTypes': ", ".join(api.get('endpointConfiguration', {}).get('types', [])),
                        },
                        status='Active',
                        tags_list=[{'Key': k, 'Value': v} for k, v in (api.get('tags') or {}).items()],
                    )
                    count += 1
            print(f"   Native API Gateway (REST v1): {count}")
        except Exception as e:
            print(f"Native API Gateway v1 failed: {_short_err(e)}")

        # ---- API Gateway v2 (HTTP / WebSocket) ----
        try:
            apigw2 = self._client('apigatewayv2')
            count = 0
            nxt = None
            while True:
                resp = apigw2.get_apis(NextToken=nxt) if nxt else apigw2.get_apis()
                for api in resp.get('Items', []):
                    apid = api['ApiId']
                    upsert_resource(
                        arn=f"arn:aws:apigateway:{self.region}::/apis/{apid}",
                        name=api.get('Name'),
                        type_norm='AWS::ApiGatewayV2::Api',
                        group='Integration',
                        details_dict={
                            'ApiId': apid,
                            'ProtocolType': api.get('ProtocolType'),
                            'ApiEndpoint': api.get('ApiEndpoint'),
                            'CreatedDate': str(api.get('CreatedDate', '')),
                        },
                        status='Active',
                        tags_list=[{'Key': k, 'Value': v} for k, v in (api.get('Tags') or {}).items()],
                    )
                    count += 1
                nxt = resp.get('NextToken')
                if not nxt:
                    break
            print(f"   Native API Gateway v2: {count}")
        except Exception as e:
            print(f"Native API Gateway v2 failed: {_short_err(e)}")

        # ---- Glue Databases + Crawlers + Jobs ----
        try:
            glue = self._client('glue')
            c_db = c_cr = c_job = 0
            try:
                for page in glue.get_paginator('get_databases').paginate():
                    for db in page.get('DatabaseList', []):
                        upsert_resource(
                            arn=f"arn:aws:glue:{self.region}:{account_id}:database/{db['Name']}",
                            name=db['Name'],
                            type_norm='AWS::Glue::Database',
                            group='Analytics',
                            details_dict={'Description': db.get('Description'), 'CreateTime': str(db.get('CreateTime', ''))},
                            status='Active',
                        )
                        c_db += 1
            except Exception:
                pass
            try:
                for page in glue.get_paginator('get_crawlers').paginate():
                    for cr in page.get('Crawlers', []):
                        upsert_resource(
                            arn=f"arn:aws:glue:{self.region}:{account_id}:crawler/{cr['Name']}",
                            name=cr['Name'],
                            type_norm='AWS::Glue::Crawler',
                            group='Analytics',
                            details_dict={'Role': cr.get('Role'), 'State': cr.get('State'), 'TablePrefix': cr.get('TablePrefix')},
                            status=cr.get('State', 'READY'),
                        )
                        c_cr += 1
            except Exception:
                pass
            try:
                for page in glue.get_paginator('get_jobs').paginate():
                    for j in page.get('Jobs', []):
                        upsert_resource(
                            arn=f"arn:aws:glue:{self.region}:{account_id}:job/{j['Name']}",
                            name=j['Name'],
                            type_norm='AWS::Glue::Job',
                            group='Analytics',
                            details_dict={'Role': j.get('Role'), 'GlueVersion': j.get('GlueVersion')},
                            status='Active',
                        )
                        c_job += 1
            except Exception:
                pass
            print(f"   Native Glue: Databases={c_db}, Crawlers={c_cr}, Jobs={c_job}")
        except Exception as e:
            print(f"Native Glue failed: {_short_err(e)}")

        # ---- Athena WorkGroups (manual pagination — no boto3 paginator) ----
        try:
            ath = self._client('athena')
            count = 0
            nxt = None
            while True:
                resp = ath.list_work_groups(NextToken=nxt) if nxt else ath.list_work_groups()
                for wg in resp.get('WorkGroups', []):
                    name = wg['Name']
                    upsert_resource(
                        arn=f"arn:aws:athena:{self.region}:{account_id}:workgroup/{name}",
                        name=name,
                        type_norm='AWS::Athena::WorkGroup',
                        group='Analytics',
                        details_dict={'State': wg.get('State'), 'Description': wg.get('Description')},
                        status=wg.get('State', 'ENABLED'),
                    )
                    count += 1
                nxt = resp.get('NextToken')
                if not nxt:
                    break
            print(f"   Native Athena WorkGroups: {count}")
        except Exception as e:
            print(f"Native Athena failed: {_short_err(e)}")

        # ---- EMR Clusters ----
        try:
            emr = self._client('emr')
            count = 0
            for page in emr.get_paginator('list_clusters').paginate():
                for c in page.get('Clusters', []):
                    upsert_resource(
                        arn=c['ClusterArn'],
                        name=c['Name'],
                        type_norm='AWS::EMR::Cluster',
                        group='Analytics',
                        details_dict={
                            'Status': c.get('Status', {}).get('State'),
                            'NormalizedInstanceHours': c.get('NormalizedInstanceHours'),
                        },
                        status=c.get('Status', {}).get('State', 'RUNNING'),
                    )
                    count += 1
            print(f"   Native EMR Clusters: {count}")
        except Exception as e:
            print(f"Native EMR failed: {_short_err(e)}")

        # ---- Kinesis Data Streams ----
        try:
            kn = self._client('kinesis')
            count = 0
            paginator = kn.get_paginator('list_streams')
            for page in paginator.paginate():
                for name in page.get('StreamNames', []):
                    try:
                        s = kn.describe_stream_summary(StreamName=name)['StreamDescriptionSummary']
                        upsert_resource(
                            arn=s['StreamARN'],
                            name=name,
                            type_norm='AWS::Kinesis::Stream',
                            group='Streaming',
                            details_dict={
                                'StreamStatus': s.get('StreamStatus'),
                                'OpenShardCount': s.get('OpenShardCount'),
                                'RetentionPeriodHours': s.get('RetentionPeriodHours'),
                                'EncryptionType': s.get('EncryptionType'),
                            },
                            status=s.get('StreamStatus', 'ACTIVE'),
                        )
                        count += 1
                    except Exception:
                        continue
            print(f"   Native Kinesis Streams: {count}")
        except Exception as e:
            print(f"Native Kinesis failed: {_short_err(e)}")

        # ---- Kinesis Firehose Delivery Streams ----
        try:
            fh = self._client('firehose')
            count = 0
            nxt = None
            while True:
                resp = fh.list_delivery_streams(ExclusiveStartDeliveryStreamName=nxt) if nxt else fh.list_delivery_streams()
                names = resp.get('DeliveryStreamNames', [])
                for name in names:
                    try:
                        d = fh.describe_delivery_stream(DeliveryStreamName=name)['DeliveryStreamDescription']
                        upsert_resource(
                            arn=d['DeliveryStreamARN'],
                            name=name,
                            type_norm='AWS::KinesisFirehose::DeliveryStream',
                            group='Streaming',
                            details_dict={
                                'DeliveryStreamStatus': d.get('DeliveryStreamStatus'),
                                'DeliveryStreamType': d.get('DeliveryStreamType'),
                            },
                            status=d.get('DeliveryStreamStatus', 'ACTIVE'),
                        )
                        count += 1
                    except Exception:
                        continue
                if not resp.get('HasMoreDeliveryStreams') or not names:
                    break
                nxt = names[-1]
            print(f"   Native Firehose Streams: {count}")
        except Exception as e:
            print(f"Native Firehose failed: {_short_err(e)}")

        # ---- IAM Users + Groups + Policies (customer-managed) ----
        try:
            iam = self._client('iam')
            c_u = c_g = c_p = 0
            for page in iam.get_paginator('list_users').paginate():
                for u in page['Users']:
                    upsert_resource(
                        arn=u['Arn'],
                        name=u['UserName'],
                        type_norm='AWS::IAM::User',
                        group='Identity',
                        details_dict={'CreateDate': str(u.get('CreateDate', '')), 'Path': u.get('Path')},
                        status='Active',
                    )
                    c_u += 1
            for page in iam.get_paginator('list_groups').paginate():
                for g in page['Groups']:
                    upsert_resource(
                        arn=g['Arn'],
                        name=g['GroupName'],
                        type_norm='AWS::IAM::Group',
                        group='Identity',
                        details_dict={'CreateDate': str(g.get('CreateDate', '')), 'Path': g.get('Path')},
                        status='Active',
                    )
                    c_g += 1
            for page in iam.get_paginator('list_policies').paginate(Scope='Local'):
                for p in page['Policies']:
                    upsert_resource(
                        arn=p['Arn'],
                        name=p['PolicyName'],
                        type_norm='AWS::IAM::ManagedPolicy',
                        group='Identity',
                        details_dict={'AttachmentCount': p.get('AttachmentCount'), 'DefaultVersionId': p.get('DefaultVersionId')},
                        status='Active',
                    )
                    c_p += 1
            print(f"   Native IAM: Users={c_u}, Groups={c_g}, ManagedPolicies={c_p}")
        except Exception as e:
            print(f"Native IAM Users/Groups/Policies failed: {_short_err(e)}")

        # ---- CloudWatch Alarms ----
        try:
            cw = self._client('cloudwatch')
            count = 0
            for page in cw.get_paginator('describe_alarms').paginate():
                for a in page.get('MetricAlarms', []) + page.get('CompositeAlarms', []):
                    upsert_resource(
                        arn=a['AlarmArn'],
                        name=a.get('AlarmName'),
                        type_norm='AWS::CloudWatch::Alarm',
                        group='Monitoring',
                        details_dict={
                            'StateValue': a.get('StateValue'),
                            'MetricName': a.get('MetricName'),
                            'Namespace': a.get('Namespace'),
                            'ActionsEnabled': a.get('ActionsEnabled'),
                        },
                        status=a.get('StateValue', 'OK'),
                    )
                    count += 1
            print(f"   Native CloudWatch Alarms: {count}")
        except Exception as e:
            print(f"Native CloudWatch Alarms failed: {_short_err(e)}")

        # ---- CloudWatch Log Groups ----
        try:
            logs = self._client('logs')
            count = 0
            for page in logs.get_paginator('describe_log_groups').paginate():
                for lg in page.get('logGroups', []):
                    upsert_resource(
                        arn=lg['arn'].rstrip(':*'),
                        name=lg['logGroupName'],
                        type_norm='AWS::Logs::LogGroup',
                        group='Monitoring',
                        details_dict={
                            'StoredBytes': lg.get('storedBytes'),
                            'RetentionInDays': lg.get('retentionInDays'),
                            'KmsKeyId': lg.get('kmsKeyId'),
                        },
                        status='Active',
                    )
                    count += 1
            print(f"   Native CloudWatch Log Groups: {count}")
        except Exception as e:
            print(f"Native CloudWatch Logs failed: {_short_err(e)}")

        # ---- AWS Backup Vaults + Plans ----
        try:
            bk = self._client('backup')
            c_v = c_p = 0
            for page in bk.get_paginator('list_backup_vaults').paginate():
                for v in page.get('BackupVaultList', []):
                    upsert_resource(
                        arn=v['BackupVaultArn'],
                        name=v['BackupVaultName'],
                        type_norm='AWS::Backup::BackupVault',
                        group='Backup',
                        details_dict={
                            'EncryptionKeyArn': v.get('EncryptionKeyArn'),
                            'NumberOfRecoveryPoints': v.get('NumberOfRecoveryPoints'),
                            'CreationDate': str(v.get('CreationDate', '')),
                        },
                        status='Active',
                    )
                    c_v += 1
            for page in bk.get_paginator('list_backup_plans').paginate():
                for p in page.get('BackupPlansList', []):
                    upsert_resource(
                        arn=p['BackupPlanArn'],
                        name=p.get('BackupPlanName'),
                        type_norm='AWS::Backup::BackupPlan',
                        group='Backup',
                        details_dict={
                            'VersionId': p.get('VersionId'),
                            'CreationDate': str(p.get('CreationDate', '')),
                            'LastExecutionDate': str(p.get('LastExecutionDate', '')),
                        },
                        status='Active',
                    )
                    c_p += 1
            print(f"   Native AWS Backup: Vaults={c_v}, Plans={c_p}")
        except Exception as e:
            print(f"Native AWS Backup failed: {_short_err(e)}")

        # ---- Direct Connect Connections ----
        try:
            dx = self._client('directconnect')
            count = 0
            for c in dx.describe_connections().get('connections', []):
                cid = c['connectionId']
                upsert_resource(
                    arn=f"arn:aws:directconnect:{self.region}:{account_id}:dxcon/{cid}",
                    name=c.get('connectionName') or cid,
                    type_norm='AWS::DirectConnect::Connection',
                    group='Networking',
                    details_dict={
                        'ConnectionState': c.get('connectionState'),
                        'Bandwidth': c.get('bandwidth'),
                        'Location': c.get('location'),
                        'PartnerName': c.get('partnerName'),
                    },
                    status=c.get('connectionState', 'available'),
                    tags_list=c.get('tags', []),
                )
                count += 1
            print(f"   Native Direct Connect: {count}")
        except Exception as e:
            print(f"Native Direct Connect failed: {_short_err(e)}")

        # =====================================================================
        # AI / ML services — every block independent, wrapped in try/except so
        # an unsupported service or denied permission only skips that scan.
        # Older boto3 versions may not have some of the newer clients; the
        # outer try catches that via _short_err and prints a clean skip note.
        # =====================================================================
        print(f"Starting AI/ML native scans ({sum(1 for _ in [1]*19)} services)...", flush=True)
        import time as _time
        _ai_start = _time.monotonic()

        # ---- SageMaker (ML platform: notebooks, endpoints, models, etc.) ----
        try:
            sm = self._client('sagemaker')
            c_nb = c_ep = c_mdl = c_dom = c_pl = c_fg = c_hu = c_fd = 0
            try:
                for page in sm.get_paginator('list_notebook_instances').paginate():
                    for n in page.get('NotebookInstances', []):
                        upsert_resource(
                            arn=n['NotebookInstanceArn'], name=n['NotebookInstanceName'],
                            type_norm='AWS::SageMaker::NotebookInstance', group='AI/ML',
                            details_dict={
                                'InstanceType': n.get('InstanceType'),
                                'NotebookInstanceStatus': n.get('NotebookInstanceStatus'),
                                'CreationTime': str(n.get('CreationTime', '')),
                            },
                            status=n.get('NotebookInstanceStatus', 'InService'),
                        )
                        c_nb += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_endpoints').paginate():
                    for e in page.get('Endpoints', []):
                        upsert_resource(
                            arn=e['EndpointArn'], name=e['EndpointName'],
                            type_norm='AWS::SageMaker::Endpoint', group='AI/ML',
                            details_dict={'EndpointStatus': e.get('EndpointStatus'),
                                          'CreationTime': str(e.get('CreationTime', ''))},
                            status=e.get('EndpointStatus', 'InService'),
                        )
                        c_ep += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_models').paginate():
                    for m in page.get('Models', []):
                        upsert_resource(
                            arn=m['ModelArn'], name=m['ModelName'],
                            type_norm='AWS::SageMaker::Model', group='AI/ML',
                            details_dict={'CreationTime': str(m.get('CreationTime', ''))},
                            status='Active',
                        )
                        c_mdl += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_domains').paginate():
                    for d in page.get('Domains', []):
                        upsert_resource(
                            arn=d['DomainArn'], name=d['DomainName'],
                            type_norm='AWS::SageMaker::Domain', group='AI/ML',
                            details_dict={'Status': d.get('Status'), 'Url': d.get('Url'),
                                          'CreationTime': str(d.get('CreationTime', ''))},
                            status=d.get('Status', 'InService'),
                        )
                        c_dom += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_pipelines').paginate():
                    for p in page.get('PipelineSummaries', []):
                        upsert_resource(
                            arn=p['PipelineArn'], name=p.get('PipelineName'),
                            type_norm='AWS::SageMaker::Pipeline', group='AI/ML',
                            details_dict={'PipelineDisplayName': p.get('PipelineDisplayName'),
                                          'CreationTime': str(p.get('CreationTime', ''))},
                            status='Active',
                        )
                        c_pl += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_feature_groups').paginate():
                    for fg in page.get('FeatureGroupSummaries', []):
                        upsert_resource(
                            arn=fg['FeatureGroupArn'], name=fg.get('FeatureGroupName'),
                            type_norm='AWS::SageMaker::FeatureGroup', group='AI/ML',
                            details_dict={'FeatureGroupStatus': fg.get('FeatureGroupStatus'),
                                          'CreationTime': str(fg.get('CreationTime', ''))},
                            status=fg.get('FeatureGroupStatus', 'Created'),
                        )
                        c_fg += 1
            except Exception: pass
            try:
                # Augmented AI (Human-in-the-loop): listed via SageMaker
                for page in sm.get_paginator('list_human_task_uis').paginate():
                    for h in page.get('HumanTaskUiSummaries', []):
                        upsert_resource(
                            arn=h['HumanTaskUiArn'], name=h.get('HumanTaskUiName'),
                            type_norm='AWS::SageMaker::HumanTaskUi', group='AI/ML',
                            details_dict={'CreationTime': str(h.get('CreationTime', ''))},
                            status='Active',
                        )
                        c_hu += 1
            except Exception: pass
            try:
                for page in sm.get_paginator('list_flow_definitions').paginate():
                    for f in page.get('FlowDefinitionSummaries', []):
                        upsert_resource(
                            arn=f['FlowDefinitionArn'], name=f.get('FlowDefinitionName'),
                            type_norm='AWS::SageMaker::FlowDefinition', group='AI/ML',
                            details_dict={'FlowDefinitionStatus': f.get('FlowDefinitionStatus'),
                                          'CreationTime': str(f.get('CreationTime', ''))},
                            status=f.get('FlowDefinitionStatus', 'Active'),
                        )
                        c_fd += 1
            except Exception: pass
            print(f"   Native SageMaker: Notebooks={c_nb}, Endpoints={c_ep}, "
                  f"Models={c_mdl}, Domains={c_dom}, Pipelines={c_pl}, "
                  f"FeatureGroups={c_fg}, HumanTaskUIs={c_hu}, FlowDefs={c_fd}")
        except Exception as e:
            print(f"Native SageMaker failed: {_short_err(e)}")

        # ---- Bedrock (foundation models, custom models, agents, KBs) ----
        try:
            br = self._client('bedrock')
            c_cm = c_gr = c_pt = 0
            try:
                for m in br.list_custom_models().get('modelSummaries', []):
                    upsert_resource(
                        arn=m['modelArn'], name=m.get('modelName'),
                        type_norm='AWS::Bedrock::CustomModel', group='AI/ML',
                        details_dict={'BaseModelArn': m.get('baseModelArn'),
                                      'CreationTime': str(m.get('creationTime', ''))},
                        status='Active',
                    )
                    c_cm += 1
            except Exception: pass
            try:
                for g in br.list_guardrails().get('guardrails', []):
                    upsert_resource(
                        arn=g['arn'], name=g.get('name'),
                        type_norm='AWS::Bedrock::Guardrail', group='AI/ML',
                        details_dict={'Status': g.get('status'), 'Version': g.get('version')},
                        status=g.get('status', 'READY'),
                    )
                    c_gr += 1
            except Exception: pass
            try:
                for p in br.list_provisioned_model_throughputs().get('provisionedModelSummaries', []):
                    upsert_resource(
                        arn=p['provisionedModelArn'], name=p.get('provisionedModelName'),
                        type_norm='AWS::Bedrock::ProvisionedModelThroughput', group='AI/ML',
                        details_dict={'ModelArn': p.get('modelArn'),
                                      'ModelUnits': p.get('modelUnits'),
                                      'Status': p.get('status')},
                        status=p.get('status', 'InService'),
                    )
                    c_pt += 1
            except Exception: pass
            print(f"   Native Bedrock: CustomModels={c_cm}, Guardrails={c_gr}, ProvisionedThroughputs={c_pt}")
        except Exception as e:
            print(f"Native Bedrock failed: {_short_err(e)}")

        # ---- Bedrock Agents (knowledge bases, agents, prompts) ----
        try:
            ba = self._client('bedrock-agent')
            c_kb = c_ag = c_pr = 0
            try:
                for page in ba.get_paginator('list_knowledge_bases').paginate():
                    for kb in page.get('knowledgeBaseSummaries', []):
                        upsert_resource(
                            arn=f"arn:aws:bedrock:{self.region}:{account_id}:knowledge-base/{kb['knowledgeBaseId']}",
                            name=kb.get('name'),
                            type_norm='AWS::Bedrock::KnowledgeBase', group='AI/ML',
                            details_dict={'KnowledgeBaseId': kb['knowledgeBaseId'],
                                          'Status': kb.get('status'),
                                          'UpdatedAt': str(kb.get('updatedAt', ''))},
                            status=kb.get('status', 'ACTIVE'),
                        )
                        c_kb += 1
            except Exception: pass
            try:
                for page in ba.get_paginator('list_agents').paginate():
                    for ag in page.get('agentSummaries', []):
                        upsert_resource(
                            arn=f"arn:aws:bedrock:{self.region}:{account_id}:agent/{ag['agentId']}",
                            name=ag.get('agentName'),
                            type_norm='AWS::Bedrock::Agent', group='AI/ML',
                            details_dict={'AgentId': ag['agentId'],
                                          'AgentStatus': ag.get('agentStatus'),
                                          'UpdatedAt': str(ag.get('updatedAt', ''))},
                            status=ag.get('agentStatus', 'PREPARED'),
                        )
                        c_ag += 1
            except Exception: pass
            try:
                for page in ba.get_paginator('list_prompts').paginate():
                    for pr in page.get('promptSummaries', []):
                        upsert_resource(
                            arn=pr.get('arn') or f"arn:aws:bedrock:{self.region}:{account_id}:prompt/{pr['id']}",
                            name=pr.get('name'),
                            type_norm='AWS::Bedrock::Prompt', group='AI/ML',
                            details_dict={'Id': pr.get('id'), 'Version': pr.get('version')},
                            status='Active',
                        )
                        c_pr += 1
            except Exception: pass
            print(f"   Native Bedrock Agents: KnowledgeBases={c_kb}, Agents={c_ag}, Prompts={c_pr}")
        except Exception as e:
            print(f"Native Bedrock Agents failed: {_short_err(e)}")

        # ---- Comprehend (NLP) ----
        try:
            cp = self._client('comprehend')
            c_dc = c_er = c_ep = 0
            try:
                for page in cp.get_paginator('list_document_classifiers').paginate():
                    for d in page.get('DocumentClassifierPropertiesList', []):
                        upsert_resource(
                            arn=d['DocumentClassifierArn'], name=d['DocumentClassifierArn'].split('/')[-1],
                            type_norm='AWS::Comprehend::DocumentClassifier', group='AI/ML',
                            details_dict={'Status': d.get('Status'),
                                          'LanguageCode': d.get('LanguageCode')},
                            status=d.get('Status', 'TRAINED'),
                        )
                        c_dc += 1
            except Exception: pass
            try:
                for page in cp.get_paginator('list_entity_recognizers').paginate():
                    for e in page.get('EntityRecognizerPropertiesList', []):
                        upsert_resource(
                            arn=e['EntityRecognizerArn'], name=e['EntityRecognizerArn'].split('/')[-1],
                            type_norm='AWS::Comprehend::EntityRecognizer', group='AI/ML',
                            details_dict={'Status': e.get('Status'),
                                          'LanguageCode': e.get('LanguageCode')},
                            status=e.get('Status', 'TRAINED'),
                        )
                        c_er += 1
            except Exception: pass
            try:
                for page in cp.get_paginator('list_endpoints').paginate():
                    for ep in page.get('EndpointPropertiesList', []):
                        upsert_resource(
                            arn=ep['EndpointArn'], name=ep['EndpointArn'].split('/')[-1],
                            type_norm='AWS::Comprehend::Endpoint', group='AI/ML',
                            details_dict={'Status': ep.get('Status'),
                                          'ModelArn': ep.get('ModelArn')},
                            status=ep.get('Status', 'IN_SERVICE'),
                        )
                        c_ep += 1
            except Exception: pass
            print(f"   Native Comprehend: Classifiers={c_dc}, EntityRecognizers={c_er}, Endpoints={c_ep}")
        except Exception as e:
            print(f"Native Comprehend failed: {_short_err(e)}")

        # ---- Rekognition (image/video analysis) ----
        try:
            rek = self._client('rekognition')
            c_col = c_sp = c_prj = 0
            try:
                for page in rek.get_paginator('list_collections').paginate():
                    for col in page.get('CollectionIds', []):
                        upsert_resource(
                            arn=f"arn:aws:rekognition:{self.region}:{account_id}:collection/{col}",
                            name=col, type_norm='AWS::Rekognition::Collection', group='AI/ML',
                            details_dict={'CollectionId': col}, status='Active',
                        )
                        c_col += 1
            except Exception: pass
            try:
                for sp in rek.list_stream_processors().get('StreamProcessors', []):
                    upsert_resource(
                        arn=f"arn:aws:rekognition:{self.region}:{account_id}:streamprocessor/{sp['Name']}",
                        name=sp['Name'], type_norm='AWS::Rekognition::StreamProcessor', group='AI/ML',
                        details_dict={'Status': sp.get('Status')},
                        status=sp.get('Status', 'STOPPED'),
                    )
                    c_sp += 1
            except Exception: pass
            try:
                for prj in rek.describe_projects().get('ProjectDescriptions', []):
                    upsert_resource(
                        arn=prj['ProjectArn'], name=prj['ProjectArn'].split('/')[-2],
                        type_norm='AWS::Rekognition::Project', group='AI/ML',
                        details_dict={'Status': prj.get('Status'),
                                      'CreationTimestamp': str(prj.get('CreationTimestamp', ''))},
                        status=prj.get('Status', 'CREATED'),
                    )
                    c_prj += 1
            except Exception: pass
            print(f"   Native Rekognition: Collections={c_col}, StreamProcessors={c_sp}, Projects={c_prj}")
        except Exception as e:
            print(f"Native Rekognition failed: {_short_err(e)}")

        # ---- Polly (text-to-speech) ----
        try:
            polly = self._client('polly')
            count = 0
            for page in polly.get_paginator('list_lexicons').paginate():
                for lx in page.get('Lexicons', []):
                    upsert_resource(
                        arn=f"arn:aws:polly:{self.region}:{account_id}:lexicon/{lx['Name']}",
                        name=lx['Name'], type_norm='AWS::Polly::Lexicon', group='AI/ML',
                        details_dict={'Attributes': str(lx.get('Attributes', {}))},
                        status='Active',
                    )
                    count += 1
            print(f"   Native Polly Lexicons: {count}")
        except Exception as e:
            print(f"Native Polly failed: {_short_err(e)}")

        # ---- Transcribe (speech-to-text) ----
        try:
            ts = self._client('transcribe')
            c_v = c_tj = c_lm = c_vf = 0
            try:
                resp = ts.list_vocabularies(MaxResults=100)
                for v in resp.get('Vocabularies', []):
                    name = v.get('VocabularyName')
                    upsert_resource(
                        arn=f"arn:aws:transcribe:{self.region}:{account_id}:vocabulary/{name}",
                        name=name, type_norm='AWS::Transcribe::Vocabulary', group='AI/ML',
                        details_dict={'LanguageCode': v.get('LanguageCode'),
                                      'VocabularyState': v.get('VocabularyState')},
                        status=v.get('VocabularyState', 'READY'),
                    )
                    c_v += 1
            except Exception: pass
            try:
                resp = ts.list_transcription_jobs(MaxResults=100)
                for j in resp.get('TranscriptionJobSummaries', []):
                    name = j.get('TranscriptionJobName')
                    upsert_resource(
                        arn=f"arn:aws:transcribe:{self.region}:{account_id}:transcription-job/{name}",
                        name=name, type_norm='AWS::Transcribe::TranscriptionJob', group='AI/ML',
                        details_dict={'LanguageCode': j.get('LanguageCode'),
                                      'TranscriptionJobStatus': j.get('TranscriptionJobStatus')},
                        status=j.get('TranscriptionJobStatus', 'COMPLETED'),
                    )
                    c_tj += 1
            except Exception: pass
            try:
                resp = ts.list_language_models()
                for lm in resp.get('Models', []):
                    name = lm.get('ModelName')
                    upsert_resource(
                        arn=f"arn:aws:transcribe:{self.region}:{account_id}:language-model/{name}",
                        name=name, type_norm='AWS::Transcribe::LanguageModel', group='AI/ML',
                        details_dict={'ModelStatus': lm.get('ModelStatus'),
                                      'LanguageCode': lm.get('LanguageCode')},
                        status=lm.get('ModelStatus', 'COMPLETED'),
                    )
                    c_lm += 1
            except Exception: pass
            try:
                resp = ts.list_vocabulary_filters(MaxResults=100)
                for vf in resp.get('VocabularyFilters', []):
                    name = vf.get('VocabularyFilterName')
                    upsert_resource(
                        arn=f"arn:aws:transcribe:{self.region}:{account_id}:vocabulary-filter/{name}",
                        name=name, type_norm='AWS::Transcribe::VocabularyFilter', group='AI/ML',
                        details_dict={'LanguageCode': vf.get('LanguageCode')},
                        status='Active',
                    )
                    c_vf += 1
            except Exception: pass
            print(f"   Native Transcribe: Vocabs={c_v}, Jobs={c_tj}, LangModels={c_lm}, VocabFilters={c_vf}")
        except Exception as e:
            print(f"Native Transcribe failed: {_short_err(e)}")

        # ---- Translate ----
        try:
            tr = self._client('translate')
            c_t = c_pd = 0
            try:
                for page in tr.get_paginator('list_terminologies').paginate():
                    for t in page.get('TerminologyPropertiesList', []):
                        upsert_resource(
                            arn=t.get('Arn') or f"arn:aws:translate:{self.region}:{account_id}:terminology/{t['Name']}",
                            name=t['Name'], type_norm='AWS::Translate::Terminology', group='AI/ML',
                            details_dict={'SourceLanguageCode': t.get('SourceLanguageCode'),
                                          'TermCount': t.get('TermCount')},
                            status='Active',
                        )
                        c_t += 1
            except Exception: pass
            try:
                for page in tr.get_paginator('list_parallel_data').paginate():
                    for pd in page.get('ParallelDataPropertiesList', []):
                        upsert_resource(
                            arn=pd.get('Arn') or f"arn:aws:translate:{self.region}:{account_id}:parallel-data/{pd['Name']}",
                            name=pd['Name'], type_norm='AWS::Translate::ParallelData', group='AI/ML',
                            details_dict={'Status': pd.get('Status'),
                                          'SourceLanguageCode': pd.get('SourceLanguageCode')},
                            status=pd.get('Status', 'ACTIVE'),
                        )
                        c_pd += 1
            except Exception: pass
            print(f"   Native Translate: Terminologies={c_t}, ParallelData={c_pd}")
        except Exception as e:
            print(f"Native Translate failed: {_short_err(e)}")

        # ---- Lex V2 (chatbots) — manual NextToken pagination ----
        if _service_available_here('lexv2-models'):
            try:
                lex = self._client('lexv2-models')
                count = 0
                nxt = None
                while True:
                    resp = lex.list_bots(nextToken=nxt) if nxt else lex.list_bots()
                    for b in resp.get('botSummaries', []):
                        upsert_resource(
                            arn=f"arn:aws:lex:{self.region}:{account_id}:bot/{b['botId']}",
                            name=b.get('botName'),
                            type_norm='AWS::Lex::Bot', group='AI/ML',
                            details_dict={'BotId': b['botId'], 'BotStatus': b.get('botStatus'),
                                          'LatestBotVersion': b.get('latestBotVersion')},
                            status=b.get('botStatus', 'Available'),
                        )
                        count += 1
                    nxt = resp.get('nextToken')
                    if not nxt: break
                print(f"   Native Lex V2 Bots: {count}")
            except Exception as e:
                print(f"Native Lex V2 failed: {_short_err(e)}")

        # ---- Kendra (enterprise search) — manual NextToken pagination ----
        try:
            kn = self._client('kendra')
            count = 0
            nxt = None
            while True:
                resp = kn.list_indices(NextToken=nxt) if nxt else kn.list_indices()
                for idx in resp.get('IndexConfigurationSummaryItems', []):
                    upsert_resource(
                        arn=f"arn:aws:kendra:{self.region}:{account_id}:index/{idx['Id']}",
                        name=idx.get('Name'),
                        type_norm='AWS::Kendra::Index', group='AI/ML',
                        details_dict={'Id': idx['Id'], 'Edition': idx.get('Edition'),
                                      'Status': idx.get('Status')},
                        status=idx.get('Status', 'ACTIVE'),
                    )
                    count += 1
                nxt = resp.get('NextToken')
                if not nxt: break
            print(f"   Native Kendra Indexes: {count}")
        except Exception as e:
            print(f"Native Kendra failed: {_short_err(e)}")

        # ---- Textract (Adapters) ----
        try:
            tx = self._client('textract')
            count = 0
            nxt = None
            while True:
                resp = tx.list_adapters(MaxResults=100, NextToken=nxt) if nxt else tx.list_adapters(MaxResults=100)
                for a in resp.get('Adapters', []):
                    upsert_resource(
                        arn=f"arn:aws:textract:{self.region}:{account_id}:adapter/{a['AdapterId']}",
                        name=a.get('AdapterName'),
                        type_norm='AWS::Textract::Adapter', group='AI/ML',
                        details_dict={'AdapterId': a['AdapterId'],
                                      'FeatureTypes': ", ".join(a.get('FeatureTypes') or [])},
                        status='Active',
                    )
                    count += 1
                nxt = resp.get('NextToken')
                if not nxt: break
            print(f"   Native Textract Adapters: {count}")
        except Exception as e:
            print(f"Native Textract failed: {_short_err(e)}")

        # ---- Personalize ----
        try:
            pz = self._client('personalize')
            c_dg = c_sol = c_cmp = 0
            try:
                for page in pz.get_paginator('list_dataset_groups').paginate():
                    for dg in page.get('datasetGroups', []):
                        upsert_resource(
                            arn=dg['datasetGroupArn'], name=dg.get('name'),
                            type_norm='AWS::Personalize::DatasetGroup', group='AI/ML',
                            details_dict={'Status': dg.get('status')},
                            status=dg.get('status', 'ACTIVE'),
                        )
                        c_dg += 1
            except Exception: pass
            try:
                for page in pz.get_paginator('list_solutions').paginate():
                    for s in page.get('solutions', []):
                        upsert_resource(
                            arn=s['solutionArn'], name=s.get('name'),
                            type_norm='AWS::Personalize::Solution', group='AI/ML',
                            details_dict={'Status': s.get('status')},
                            status=s.get('status', 'ACTIVE'),
                        )
                        c_sol += 1
            except Exception: pass
            try:
                for page in pz.get_paginator('list_campaigns').paginate():
                    for c in page.get('campaigns', []):
                        upsert_resource(
                            arn=c['campaignArn'], name=c.get('name'),
                            type_norm='AWS::Personalize::Campaign', group='AI/ML',
                            details_dict={'Status': c.get('status')},
                            status=c.get('status', 'ACTIVE'),
                        )
                        c_cmp += 1
            except Exception: pass
            print(f"   Native Personalize: DatasetGroups={c_dg}, Solutions={c_sol}, Campaigns={c_cmp}")
        except Exception as e:
            print(f"Native Personalize failed: {_short_err(e)}")

        # ---- Forecast ----
        try:
            fc = self._client('forecast')
            c_ds = c_pr = c_fcst = 0
            try:
                for page in fc.get_paginator('list_datasets').paginate():
                    for d in page.get('Datasets', []):
                        upsert_resource(
                            arn=d['DatasetArn'], name=d.get('DatasetName'),
                            type_norm='AWS::Forecast::Dataset', group='AI/ML',
                            details_dict={'DatasetType': d.get('DatasetType'),
                                          'Domain': d.get('Domain')},
                            status='Active',
                        )
                        c_ds += 1
            except Exception: pass
            try:
                for page in fc.get_paginator('list_predictors').paginate():
                    for p in page.get('Predictors', []):
                        upsert_resource(
                            arn=p['PredictorArn'], name=p.get('PredictorName'),
                            type_norm='AWS::Forecast::Predictor', group='AI/ML',
                            details_dict={'Status': p.get('Status')},
                            status=p.get('Status', 'ACTIVE'),
                        )
                        c_pr += 1
            except Exception: pass
            try:
                for page in fc.get_paginator('list_forecasts').paginate():
                    for f in page.get('Forecasts', []):
                        upsert_resource(
                            arn=f['ForecastArn'], name=f.get('ForecastName'),
                            type_norm='AWS::Forecast::Forecast', group='AI/ML',
                            details_dict={'Status': f.get('Status')},
                            status=f.get('Status', 'ACTIVE'),
                        )
                        c_fcst += 1
            except Exception: pass
            print(f"   Native Forecast: Datasets={c_ds}, Predictors={c_pr}, Forecasts={c_fcst}")
        except Exception as e:
            print(f"Native Forecast failed: {_short_err(e)}")

        # ---- Lookout for Metrics / Vision ----
        # AWS retired both services in October 2025. boto3 no longer ships
        # `lookoutmetrics` or `lookoutvision` clients, and the endpoints are
        # gone. We removed the previous scans to avoid noise. The Azure
        # mapping entries are retained for historical assessment files.

        # ---- Lookout for Equipment ----
        try:
            le = self._client('lookoutequipment')
            c_d = c_m = 0
            try:
                for page in le.get_paginator('list_datasets').paginate():
                    for d in page.get('DatasetSummaries', []):
                        upsert_resource(
                            arn=d['DatasetArn'], name=d.get('DatasetName'),
                            type_norm='AWS::LookoutEquipment::Dataset', group='AI/ML',
                            details_dict={'Status': d.get('Status')},
                            status=d.get('Status', 'ACTIVE'),
                        )
                        c_d += 1
            except Exception: pass
            try:
                for page in le.get_paginator('list_models').paginate():
                    for m in page.get('ModelSummaries', []):
                        upsert_resource(
                            arn=m['ModelArn'], name=m.get('ModelName'),
                            type_norm='AWS::LookoutEquipment::Model', group='AI/ML',
                            details_dict={'Status': m.get('Status')},
                            status=m.get('Status', 'SUCCESS'),
                        )
                        c_m += 1
            except Exception: pass
            print(f"   Native Lookout for Equipment: Datasets={c_d}, Models={c_m}")
        except Exception as e:
            print(f"Native Lookout for Equipment failed: {_short_err(e)}")

        # Lookout for Metrics / Vision scans removed — services retired by AWS
        # in October 2025; boto3 no longer provides the clients.

        # ---- DevOps Guru ----
        try:
            dg = self._client('devops-guru')
            # DevOps Guru's "resource collection" is the set of resources it monitors.
            try:
                resp = dg.get_resource_collection(ResourceCollectionType='AWS_CLOUD_FORMATION')
                rc = resp.get('ResourceCollection', {})
                stacks = rc.get('CloudFormation', {}).get('StackNames') or []
                if stacks:
                    upsert_resource(
                        arn=f"arn:aws:devops-guru:{self.region}:{account_id}:resource-collection/cloudformation",
                        name='CloudFormation collection',
                        type_norm='AWS::DevOpsGuru::ResourceCollection', group='AI/ML',
                        details_dict={'StackCount': len(stacks), 'StackNames': ", ".join(stacks[:10])},
                        status='Active',
                    )
                    print(f"   Native DevOps Guru: monitoring {len(stacks)} CFN stack(s)")
                else:
                    print("   Native DevOps Guru: no resource collection configured")
            except Exception: pass
        except Exception as e:
            print(f"Native DevOps Guru failed: {_short_err(e)}")

        # CodeGuru Reviewer scan removed — AWS sunsetted the service in
        # June 2024. (CodeGuru Profiler is still alive — kept below.)

        # ---- CodeGuru Profiler — manual NextToken pagination ----
        if _service_available_here('codeguruprofiler'):
            try:
                cgp = self._client('codeguruprofiler')
                count = 0
                nxt = None
                while True:
                    resp = cgp.list_profiling_groups(nextToken=nxt) if nxt else cgp.list_profiling_groups()
                    for pg in resp.get('profilingGroupNames', []):
                        upsert_resource(
                            arn=f"arn:aws:codeguru-profiler:{self.region}:{account_id}:profilingGroup/{pg}",
                            name=pg,
                            type_norm='AWS::CodeGuruProfiler::ProfilingGroup', group='AI/ML',
                            details_dict={'Name': pg},
                            status='Active',
                        )
                        count += 1
                    nxt = resp.get('nextToken')
                    if not nxt: break
                print(f"   Native CodeGuru Profiler: ProfilingGroups={count}")
            except Exception as e:
                print(f"Native CodeGuru Profiler failed: {_short_err(e)}")

        # ---- Fraud Detector ----
        try:
            fd = self._client('frauddetector')
            c_d = c_m = 0
            try:
                resp = fd.get_detectors()
                for d in resp.get('detectors', []):
                    upsert_resource(
                        arn=d.get('arn') or f"arn:aws:frauddetector:{self.region}:{account_id}:detector/{d['detectorId']}",
                        name=d['detectorId'],
                        type_norm='AWS::FraudDetector::Detector', group='AI/ML',
                        details_dict={'Description': d.get('description'),
                                      'LastUpdatedTime': str(d.get('lastUpdatedTime', ''))},
                        status='Active',
                    )
                    c_d += 1
            except Exception: pass
            try:
                resp = fd.get_models()
                for m in resp.get('models', []):
                    upsert_resource(
                        arn=m.get('arn') or f"arn:aws:frauddetector:{self.region}:{account_id}:model/{m['modelId']}",
                        name=m['modelId'],
                        type_norm='AWS::FraudDetector::Model', group='AI/ML',
                        details_dict={'ModelType': m.get('modelType')},
                        status='Active',
                    )
                    c_m += 1
            except Exception: pass
            print(f"   Native Fraud Detector: Detectors={c_d}, Models={c_m}")
        except Exception as e:
            print(f"Native Fraud Detector failed: {_short_err(e)}")

        # ---- HealthLake (FHIR datastores) — manual NextToken pagination ----
        try:
            hl = self._client('healthlake')
            count = 0
            nxt = None
            while True:
                resp = hl.list_fhir_datastores(NextToken=nxt) if nxt else hl.list_fhir_datastores()
                for d in resp.get('DatastorePropertiesList', []):
                    upsert_resource(
                        arn=d['DatastoreArn'], name=d.get('DatastoreName'),
                        type_norm='AWS::HealthLake::FHIRDatastore', group='AI/ML',
                        details_dict={'DatastoreStatus': d.get('DatastoreStatus'),
                                      'DatastoreTypeVersion': d.get('DatastoreTypeVersion')},
                        status=d.get('DatastoreStatus', 'ACTIVE'),
                    )
                    count += 1
                nxt = resp.get('NextToken')
                if not nxt: break
            print(f"   Native HealthLake FHIR Datastores: {count}")
        except Exception as e:
            print(f"Native HealthLake failed: {_short_err(e)}")

        # ---- HealthOmics (Genomics) ----
        try:
            om = self._client('omics')
            c_ss = c_rs = c_wf = 0
            try:
                for page in om.get_paginator('list_sequence_stores').paginate():
                    for s in page.get('sequenceStores', []):
                        upsert_resource(
                            arn=s['arn'], name=s.get('name') or s['id'],
                            type_norm='AWS::Omics::SequenceStore', group='AI/ML',
                            details_dict={'Status': s.get('status'),
                                          'CreationTime': str(s.get('creationTime', ''))},
                            status=s.get('status', 'ACTIVE'),
                        )
                        c_ss += 1
            except Exception: pass
            try:
                for page in om.get_paginator('list_reference_stores').paginate():
                    for r in page.get('referenceStores', []):
                        upsert_resource(
                            arn=r['arn'], name=r.get('name') or r['id'],
                            type_norm='AWS::Omics::ReferenceStore', group='AI/ML',
                            details_dict={'CreationTime': str(r.get('creationTime', ''))},
                            status='Active',
                        )
                        c_rs += 1
            except Exception: pass
            try:
                for page in om.get_paginator('list_workflows').paginate():
                    for w in page.get('items', []):
                        upsert_resource(
                            arn=w.get('arn') or f"arn:aws:omics:{self.region}:{account_id}:workflow/{w['id']}",
                            name=w.get('name') or w['id'],
                            type_norm='AWS::Omics::Workflow', group='AI/ML',
                            details_dict={'Status': w.get('status'),
                                          'Type': w.get('type')},
                            status=w.get('status', 'ACTIVE'),
                        )
                        c_wf += 1
            except Exception: pass
            print(f"   Native HealthOmics: SequenceStores={c_ss}, ReferenceStores={c_rs}, Workflows={c_wf}")
        except Exception as e:
            print(f"Native HealthOmics failed: {_short_err(e)}")

        # ---- Amazon Q Business ----
        if _service_available_here('qbusiness'):
            try:
                qb = self._client('qbusiness')
                count = 0
                for page in qb.get_paginator('list_applications').paginate():
                    for app in page.get('applications', []):
                        upsert_resource(
                            arn=app.get('applicationArn') or f"arn:aws:qbusiness:{self.region}:{account_id}:application/{app['applicationId']}",
                            name=app.get('displayName'),
                            type_norm='AWS::QBusiness::Application', group='AI/ML',
                            details_dict={'Id': app['applicationId'],
                                          'Status': app.get('status')},
                            status=app.get('status', 'ACTIVE'),
                        )
                        count += 1
                print(f"   Native Amazon Q Business Applications: {count}")
            except Exception as e:
                print(f"Native Amazon Q Business failed: {_short_err(e)}")

        _ai_dur = _time.monotonic() - _ai_start
        print(f"   AI/ML scans finished in {_ai_dur:.1f}s.", flush=True)
        print(f"Native Detail Scan complete. Upserted/Added resources.")


    def _discover_untagged_resources(self, all_resources, known_arns):
        """
        Dynamically discover resources using AWS Cloud Control API for ALL available types (Parallelized).
        """
        import concurrent.futures
        
        # 1. Ensure map
        if not hasattr(self, 'cfn_type_map'):
            self._ensure_cfn_type_cache()
        # 2. Build Targeted Scan List
        # Instead of scanning ALL 1000+ types, we target:
        # A) Core infrastructure services (Networking, Compute, Storage)
        # B) Services where we already found tagged resources
        
        valid_cfn_types = set(self.cfn_type_map.values())
        found_in_tagging = set(r['ResourceType'] for r in all_resources)
        
        # 2.1 Identify services of interest
        tagged_svcs = set()
        for t in found_in_tagging:
            parts = t.split('::')
            if len(parts) > 1: tagged_svcs.add(parts[1].lower())
            
        critical_svcs = {
            'ec2', 's3', 'rds', 'lambda', 'elasticloadbalancingv2', 'autoscaling', 
            'iam', 'sns', 'sqs', 'cloudfront', 'route53', 'apigateway', 'apigatewayv2',
            'kms', 'secretsmanager', 'dynamodb', 'redshift', 'elasticache', 'efs', 'eks'
        }
        
        # 2.2 Global service & type blacklist (Slow, requires subscription, or noisy child resources)
        blacklist_svcs = {
             'securityhub', 'shield', 'guardduty', 'inspector', 'inspector2', 
             'ssm', 'logs', 'config', 'ce', 'connect', 'amplify', 'appconfig', 'appintegrations'
        }
        
        blacklist_types = {
            'AWS::IAM::GroupPolicy', 'AWS::IAM::RolePolicy', 'AWS::IAM::UserPolicy', 'AWS::IAM::ServiceLinkedRole',
            'AWS::ApiGateway::Resource', 'AWS::ApiGateway::Method', 'AWS::ApiGateway::Stage', 'AWS::ApiGateway::Deployment',
            'AWS::ApiGatewayV2::Route', 'AWS::ApiGatewayV2::Integration', 'AWS::ApiGatewayV2::Deployment',
            'AWS::S3::MultiRegionAccessPointPolicy'
        }

        # 2.3 Construct the final list
        final_scan_list = set()
        for cfn_type in valid_cfn_types:
            if cfn_type in blacklist_types: continue
            
            parts = cfn_type.split('::')
            if len(parts) < 2: continue
            svc = parts[1].lower()
            
            if svc in blacklist_svcs: continue
            
            # Include if it's a critical service OR we found something tagged there
            if svc in critical_svcs or svc in tagged_svcs:
                final_scan_list.add(cfn_type)

        # 2.4 Ensure specific tagged types are included if they match valid CFN types
        for tagged_type in found_in_tagging:
            if tagged_type in valid_cfn_types and tagged_type not in final_scan_list:
                # Check blacklist again
                svc = tagged_type.split('::')[1].lower() if '::' in tagged_type else ''
                if svc not in blacklist_svcs and tagged_type not in blacklist_types:
                    final_scan_list.add(tagged_type)

        scan_list = sorted(list(final_scan_list))
        print(f"Targeted Scan: Identified {len(scan_list)} relevant resource types to check for untagged resources (Worker Pool=10)...")
        
        found_new = 0


        def scan_type(p_type):
            results = []
            max_retries = 3
            backoff = 1
            
            try:
                paginator = self.cloud_control.get_paginator('list_resources')
                import time
                
                # Retry loop for throttling
                for attempt in range(max_retries):
                    try:
                        for page in paginator.paginate(TypeName=p_type):
                            for r_summary in page['ResourceDescriptions']:
                                identifier = r_summary['Identifier']
                                
                                # We defer dedup check to main thread merging or do it here with lock?
                                # Deferring is better for speed, but increases memory.
                                # We return all found, dedup later.
                                
                                props = {}
                                if 'Properties' in r_summary:
                                    import json
                                    props = json.loads(r_summary['Properties'])
                                    
                                name = identifier
                                # Heuristics
                                if 'Tags' in props and isinstance(props['Tags'], list):
                                     for t in props['Tags']:
                                         if t.get('Key') == 'Name': name = t.get('Value')
                                elif 'GroupName' in props: name = props['GroupName']
                                elif 'DBInstanceIdentifier' in props: name = props['DBInstanceIdentifier']
                                elif 'FunctionName' in props: name = props['FunctionName']
                                elif 'KeyName' in props: name = props['KeyName']
                                
                                parts = p_type.split('::')
                                svc = parts[1].lower() if len(parts) > 1 else 'unknown'
                                typ = parts[2].lower() if len(parts) > 2 else 'unknown'
        
                                # Ensure we get the raw ID for deduplication, not the ARN
                                dedup_id = identifier
                                
                                # Handle CloudControl "PublicIp|AllocationId" format for EIPs
                                if '|' in identifier:
                                    parts_pipe = identifier.split('|')
                                    if len(parts_pipe) == 2 and parts_pipe[1].startswith('eipalloc-'):
                                        identifier = parts_pipe[1] # Set ResourceId to AllocationId
                                        dedup_id = parts_pipe[1]
                                
                                if identifier.startswith('arn:'):
                                    p_arn = self._parse_arn(identifier)
                                    if p_arn['id_value']:
                                        dedup_id = p_arn['id_value']
        
                                results.append({
                                    'ResourceId': identifier,
                                    'ResourceName': name,
                                    'ResourceType': p_type,
                                    'NormalizedResourceType': p_type,
                                    'ResourceTypeSource': 'DynamicDiscovery',
                                    'ResourceGroup': self._get_resource_group(svc),
                                    'Location': self.region,
                                    'Status': 'Active', 
                                    'Details': identifier,
                                    'Tags': str(props.get('Tags', [])),
                                    '_arn': identifier if identifier.startswith('arn:') else None,
                                    '_service': svc,
                                    '_type_id': typ,
                                    '_id_value': dedup_id
                                })
                    except Exception as e:
                        # If it's a structural or configuration error, DO NOT RETRY - terminate early to save time
                        non_retriable_msgs = ["Resource type", "not found", "not supported", "TypeNotFound", "InvalidRequest", "AccessDenied"]
                        if any(m.lower() in str(e).lower() for m in non_retriable_msgs):
                            raise e 
                        
                        if 'Throttling' in str(e) and attempt < max_retries - 1:
                            time.sleep(backoff)
                            backoff *= 2
                            continue
                        raise e
                    break # Success
            except Exception as e:
                err_str = str(e)
                # Silence expected failures for Cloud Control (unsupported LIST, missing IDs, or regional gaps)
                silence_keywords = [
                    "does not support LIST action", 
                    "must not be null", 
                    "required key", 
                    "not found", 
                    "not supported",
                    "UnsupportedActionException",
                    "InvalidRequestException",
                    "TypeNotFoundException",
                    "AccessDeniedException",
                    "ResourceModel property",
                    "InvestigationGroup",
                    "AmplifyUIBuilder",
                    "InternalFailure",
                    "Handler returned status FAILED: null",
                    "Too Many Requests",
                    "Rate exceeded",
                    "ThrottlingException",
                    "not have access",
                    "Invalid request provided",
                    "GeneralServiceException"
                ]

                if not any(k.lower() in err_str.lower() for k in silence_keywords):
                    print(f"⚠️  Discovery throttle/failure for {p_type} after {max_retries} retries: {err_str}", flush=True)
            return results

        # Run in parallel
        # Reduced workers to 5 for stability - AWS Cloud Control is sensitive to aggressive concurrent discovery
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_type = {executor.submit(scan_type, t): t for t in scan_list}
            
            for future in concurrent.futures.as_completed(future_to_type):
                p_type = future_to_type[future]
                try:
                    res_list = future.result(timeout=300) # Increased timeout to 300s for large resource sets (e.g. SSM)
                    
                    if res_list:
                        # Dedup and Add
                        for r in res_list:
                            ident = r['_id_value']
                            if ident not in known_arns:
                                known_arns.add(ident)
                                all_resources.append(r)
                                found_count += 1
                                found_new += 1
                        
                    if found_count > 0:
                        print(f"Scanned {p_type} -> Found {found_count} new resources.", flush=True)
                    
                except Exception:
                    # Silent move forward on per-type failure (e.g. AccessDenied)
                    pass
        
        # Noisy prints removed to satisfy user request for silent "scan if access" behavior.

    def _scan_native_active_services(self, all_resources, known_arns):
        """
        Manually discover core resources using direct service clients.
        Essential for accounts that have tagging API restricted.
        """
        count = 0
        # Helper to add resource safely
        def add_res(res_id, name, cfn_type, svc, typ):
            nonlocal count
            # Use identifier for dedup
            dedup_id = res_id.split(':')[-1] if 'arn:' in res_id else res_id
            if res_id not in known_arns and dedup_id not in known_arns:
                known_arns.add(res_id)
                known_arns.add(dedup_id)
                all_resources.append({
                    'ResourceId': res_id,
                    'ResourceName': name,
                    'ResourceType': cfn_type,
                    'NormalizedResourceType': cfn_type,
                    'ResourceTypeSource': 'NativeScan',
                    'ResourceGroup': self._get_resource_group(svc),
                    'Location': self.region,
                    'Status': 'Active',
                    'Details': res_id,
                    'Tags': '[]', # Native scans don't fetch tags directly
                    '_arn': res_id if res_id.startswith('arn:') else None,
                    '_service': svc,
                    '_type_id': typ,
                    '_id_value': dedup_id
                })
                count += 1

        # Dynamic list of scanners to cover the 15+ services previously seen in the console.
        # This ensures the assessment report matches the console output summary.
        core_scans = [
            ('ec2', 'describe_instances', 'Reservations', 'Instances', 'InstanceId', 'AWS::EC2::Instance', 'ec2', 'instance'),
            ('rds', 'describe_db_instances', 'DBInstances', None, 'DBInstanceIdentifier', 'AWS::RDS::DBInstance', 'rds', 'dbinstance'),
            ('ec2', 'describe_vpcs', 'Vpcs', None, 'VpcId', 'AWS::EC2::VPC', 'ec2', 'vpc'),
            ('ec2', 'describe_subnets', 'Subnets', None, 'SubnetId', 'AWS::EC2::Subnet', 'ec2', 'subnet'),
            ('ec2', 'describe_security_groups', 'SecurityGroups', None, 'GroupId', 'AWS::EC2::SecurityGroup', 'ec2', 'securitygroup'),
            ('ec2', 'describe_volumes', 'Volumes', None, 'VolumeId', 'AWS::EC2::Volume', 'ec2', 'volume'),
            ('lambda', 'list_functions', 'Functions', None, 'FunctionArn', 'AWS::Lambda::Function', 'lambda', 'function'),
            ('sns', 'list_topics', 'Topics', None, 'TopicArn', 'AWS::SNS::Topic', 'sns', 'topic'),
            ('sqs', 'list_queues', 'QueueUrls', None, None, 'AWS::SQS::Queue', 'sqs', 'queue'),
            ('elbv2', 'describe_load_balancers', 'LoadBalancers', None, 'LoadBalancerArn', 'AWS::ElasticLoadBalancingV2::LoadBalancer', 'elasticloadbalancing', 'loadbalancer'),
            ('cloudformation', 'list_stacks', 'StackSummaries', None, 'StackId', 'AWS::CloudFormation::Stack', 'cloudformation', 'stack'),
            ('kms', 'list_keys', 'Keys', None, 'KeyArn', 'AWS::KMS::Key', 'kms', 'key'),
            ('cloudtrail', 'describe_trails', 'trailList', None, 'TrailARN', 'AWS::CloudTrail::Trail', 'cloudtrail', 'trail'),
            ('config', 'describe_configuration_recorders', 'ConfigurationRecorders', None, 'name', 'AWS::Config::ConfigurationRecorder', 'config', 'configrecorder')
        ]

        # 0. Global Scans (Only once per session)
        if not hasattr(self.__class__, '_global_scan_done'):
             setattr(self.__class__, '_global_scan_done', False)
        
        if not getattr(self.__class__, '_global_scan_done'):
            try:
                s3 = self._client('s3')
                for b in s3.list_buckets().get('Buckets', []):
                    add_res(f"arn:aws:s3:::{b['Name']}", b['Name'], 'AWS::S3::Bucket', 's3', 'bucket')
                
                iam = self._client('iam')
                for u in iam.list_users().get('Users', []):
                    add_res(u['Arn'], u['UserName'], 'AWS::IAM::User', 'iam', 'user')
                for r in iam.list_roles().get('Roles', []):
                    add_res(r['Arn'], r['RoleName'], 'AWS::IAM::Role', 'iam', 'role')
                
                cf = self._client('cloudfront')
                distributions = cf.list_distributions().get('DistributionList', {}).get('Items', [])
                for d in distributions:
                    add_res(d['ARN'], d['Id'], 'AWS::CloudFront::Distribution', 'cloudfront', 'distribution')

                r53 = self._client('route53')
                for z in r53.list_hosted_zones().get('HostedZones', []):
                    add_res(z['Id'], z['Name'], 'AWS::Route53::HostedZone', 'route53', 'hostedzone')

                setattr(self.__class__, '_global_scan_done', True)
            except: pass

        # 1. Regional Scans
        for svc, method, key, subkey, id_field, cfn_type, s_short, t_short in core_scans:
            try:
                client = self._client(svc)
                paginator = client.get_paginator(method)
                for page in paginator.paginate():
                    items = page.get(key, [])
                    if subkey:
                        # Flat map if needed (e.g. EC2 Reservations -> Instances)
                        flat_items = []
                        for i in items: flat_items.extend(i.get(subkey, []))
                        items = flat_items
                    
                    for item in items:
                        res_id = item.get(id_field) if id_field else (item if isinstance(item, str) else None)
                        if res_id:
                            name = item.get('ResourceName') or item.get('FunctionName') or item.get('DBInstanceIdentifier') or str(res_id).split('/')[-1]
                            add_res(res_id, name, cfn_type, s_short, t_short)
            except:
                continue # Silent fail as requested


        print(f"✅ Direct Native Discovery found {count} primary resources.")

    def _parse_arn(self, arn):
        # Format: arn:partition:service:region:account-id:resource-type/resource-id
        # Or: arn:partition:service:region:account-id:resource-id
        parts = arn.split(':', 5)
        
        if len(parts) < 6:
            return {'service': 'unknown', 'type_id': 'unknown', 'id_value': arn, 'region': None}

        service = parts[2]
        region = parts[3] if parts[3] else None
        resource_part = parts[5]
        
        r_type = 'unknown'
        r_id = resource_part
        
        # Check classic slashes
        if '/' in resource_part:
            r_type, r_id = resource_part.split('/', 1)
        # Check colons only if we haven't already split correctly? 
        # Actually with split 5, resource_part contains the REST.
        # e.g. "db:mysql-instance1". 
        elif ':' in resource_part:
            r_type, r_id = resource_part.split(':', 1)
        else:
            # No separator. Implicit type.
            # Handle implicit types for known services to avoid hardcoding defaults downstream?
            # User wants NO hardcoding, but implicit types are implicit.
            # We strictly need to infer type for correct CFN mapping.
            # Keep minimal defaults for pure implicit ARN services.
            defaults = {
                's3': 'bucket',
                'sqs': 'queue',
                'sns': 'topic',
                'cloudfront': 'distribution',
                'kinesis': 'stream'
            }
            if service in defaults:
                r_type = defaults[service]
        
        return {
            'service': service,
            'region': region,
            'type_id': r_type, 
            'id_value': r_id 
        }


    def _get_resource_group(self, service):
        groups = {
            'ec2': 'Compute',
            'lambda': 'Compute',
            's3': 'Storage',
            'rds': 'Database',
            'dynamodb': 'Database',
            'vpc': 'Networking',
            'elasticloadbalancing': 'Networking',
            'autoscaling': 'Compute',
            'iam': 'Identity'
        }
        return groups.get(service, 'Other')

    def _format_tags(self, tags_list):
        # API returns list of {Key, Value}, we want string representation
        return str({t['Key']: t['Value'] for t in tags_list})

    def _get_tag_value(self, tags_list, key):
        for t in tags_list:
            if t['Key'] == key:
                return t['Value']
        return None

    def _ensure_cfn_type_cache(self):
        """
        Ensure we have a local cache of all valid CloudFormation Types and Mappings.
        Integrates comprehensive metadata discovery to avoid hardcoding.
        """
        import json
        
        # Path to cache file
        cache_dir = os.path.join(os.path.dirname(__file__), "..", "data")
        os.makedirs(cache_dir, exist_ok=True)
        self.cache_file = os.path.join(cache_dir, "aws_full_resource_metadata.json")
        
        data = None
        
        # 1. Try Loading Cache
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r') as f:
                    data = json.load(f)
                    # Check if it has the structure we need
                    if 'inferred_arn_mappings' in data:
                        print(f"Loaded AWS metadata from cache: {len(data['inferred_arn_mappings'])} mappings.")
                        self.cfn_type_map = data['inferred_arn_mappings']
                        return
            except Exception as e:
                print(f"Cache load failed: {e}. Regenerating...")
                pass
                
        # 2. Regenerate Metadata (If missing or invalid)
        print("Fetching full AWS resource metadata (Services & CloudFormation Types)... This may take a minute.", flush=True)
        
        report = {
            "metadata": {
                "generated_at": datetime.now().isoformat(),
                "region": self.region
            },
            "statistics": {},
            "services": [],
            "cloudformation_types": [],
            "inferred_arn_mappings": {}
        }
        
        # A. Get Available Services
        try:
            services = self.session.get_available_services()
            report["services"] = sorted(services)
        except:
            report["services"] = []
            
        # B. Get CloudFormation Types
        cfn_client = self._client('cloudformation')
        cfn_types = []
        try:
            paginator = cfn_client.get_paginator('list_types')
            iterator = paginator.paginate(
                Visibility='PUBLIC', 
                ProvisioningType='FULLY_MUTABLE', 
                Filters={'Category': 'AWS_TYPES'}
            )
            for page in iterator:
                for summary in page['TypeSummaries']:
                    cfn_types.append(summary['TypeName'])
        except Exception as e:
            print(f"Warning: Could not list CFN types: {e}")
            
        report["cloudformation_types"] = cfn_types
        
        # C. Generate Mappings
        # Logic: "service:resource" -> "AWS::Service::Resource"
        mapping = {}
        for full_type in cfn_types:
            parts = full_type.split('::')
            if len(parts) == 3:
                s_cfn = parts[1].lower() # e.g. ec2, elasticloadbalancingv2
                r_cfn = parts[2].lower() # e.g. instance, loadbalancer
                
                # Direct Map
                mapping[f"{s_cfn}:{r_cfn}"] = full_type
                
                # Heuristics for Versions (v2, v3)
                if s_cfn.endswith('v2'):
                    mapping[f"{s_cfn[:-2]}:{r_cfn}"] = full_type
                if s_cfn.endswith('v3'):
                    mapping[f"{s_cfn[:-2]}:{r_cfn}"] = full_type
                    
        report["inferred_arn_mappings"] = mapping
        report["statistics"] = {
            "services": len(report["services"]),
            "types": len(report["cloudformation_types"]),
            "mappings": len(mapping)
        }
        
        # D. Save Cache
        try:
            with open(self.cache_file, 'w') as f:
                json.dump(report, f, indent=2)
            print(f"Generated and saved AWS metadata to {self.cache_file}", flush=True)
        except Exception as e:
            print(f"Warning: Could not save cache: {e}")
            
        # E. Set in memory
        self.cfn_type_map = mapping

    def _enrich_resources(self, resources):
        """
        Enrich resources dynamically using AWS Cloud Control API (Parallelized).
        """
        import concurrent.futures
        import json
        
        # Ensure client exists (though init covers it)
        if not hasattr(self, 'cloud_control'):
             from botocore.config import Config
             fast_fail_config = Config(
                retries={'max_attempts': 2, 'mode': 'standard'},
                connect_timeout=5,
                read_timeout=10
             )
             self.cloud_control = self._client('cloudcontrol', config=fast_fail_config)

        # Use a moderate worker pool for enrichment to speed up extraction.
        # Cloud Control API is generally permissive but can still throttle on account-wide limits.
        print(f"Enriching {len(resources)} resources using Cloud Control API (Parallel Worker Pool=8)...")
        
        def enrich_item(r):
            # Calculate baseline connectivity from discovery data if not already present
            if not r.get('Connectivity'):
                r['Connectivity'] = self._get_connectivity_info(r, r.get('NormalizedResourceType'))

            try:
                # 1. Guess proper CloudFormation Type
                cfn_type = self._guess_cfn_type(r['_service'], r['_type_id'])
                if not cfn_type or 'Unknown' in cfn_type:
                     return
                
                # Ignore noisy services
                if r['_service'] in ['ssm', 'events', 'logs']:
                    return

                # 2. Try enriched data fetch with different Identifier strategies
                candidates = []
                if r.get('_id_value'): candidates.append(r['_id_value'])
                if r.get('ResourceName') and r['ResourceName'] != r.get('_id_value'): candidates.append(r['ResourceName'])
                if r.get('_arn'): candidates.append(r['_arn'])
                
                success = False
                
                for ident in candidates:
                    try:
                        max_retries = 3
                        backoff = 0.5
                        resp = None
                        for attempt in range(max_retries):
                            try:
                                resp = self.cloud_control.get_resource(
                                    TypeName=cfn_type,
                                    Identifier=ident
                                )
                                break # Success
                            except Exception as e:
                                if 'Throttling' in str(e) and attempt < max_retries - 1:
                                    import time
                                    time.sleep(backoff)
                                    backoff *= 2
                                    continue
                                raise e

                        if resp and 'ResourceDescription' in resp and 'Properties' in resp['ResourceDescription']:
                                props = json.loads(resp['ResourceDescription']['Properties'])
                                if 'SecurityGroupIngress' in props:
                                    props['Ingress Rules'] = props.pop('SecurityGroupIngress')
                                if 'SecurityGroupEgress' in props:
                                    props['Egress Rules'] = props.pop('SecurityGroupEgress')

                                r['Connectivity'] = self._get_connectivity_info(props, cfn_type)
                                flat_props = self._flatten_dict(props)
                                # Cloud Control returns the raw CFN representation
                                # (BlockDeviceMappings, CpuOptions, MetadataOptions,
                                # EnclaveOptions, NetworkInterfaces, etc.). The
                                # native scan upstream populates clean curated
                                # fields like NameTag, DataDiskCount, DataDisks,
                                # OSDiskSize(GB), Architecture, etc. We must NOT
                                # let Cloud Control's raw dump overwrite the
                                # curated values — only fill in fields the
                                # native scan didn't already set.
                                for k, v in flat_props.items():
                                    if k not in r or r.get(k) in (None, '', [], {}):
                                        r[k] = v
                                
                                if 'State' in props:
                                    if isinstance(props['State'], str):
                                        r['Status'] = props['State']
                                    elif isinstance(props['State'], dict) and 'Name' in props['State']:
                                         r['Status'] = props['State']['Name']
                                elif 'Status' in props and isinstance(props['Status'], str):
                                     r['Status'] = props['Status']
                                     
                                r['ResourceType'] = cfn_type
                                success = True
                                break
                    except Exception:
                        continue
                
                if not success:
                     # Fallback for EIP
                     if cfn_type == 'AWS::EC2::EIP':
                         try:
                             ec2_client = self._client('ec2')
                             eips = None
                             # Try resolving via available IDs
                             res_id = r.get('_id_value') or r.get('ResourceName')
                             if res_id.startswith('eipalloc-'):
                                 eips = ec2_client.describe_addresses(AllocationIds=[res_id])
                             else:
                                 eips = ec2_client.describe_addresses(PublicIps=[res_id])
                                 
                             if eips and 'Addresses' in eips and eips['Addresses']:
                                 eip = eips['Addresses'][0]
                                 r.update(self._flatten_dict(eip))
                                 r['ResourceType'] = 'AWS::EC2::EIP'
                         except Exception:
                             pass

            except Exception as e:
                pass

        # Execute in parallel (Safe limit)
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            # We want to track progress, so we use as_completed
            futures = [executor.submit(enrich_item, r) for r in resources]
            
            count = 0
            for future in concurrent.futures.as_completed(futures):
                count += 1
                if count % 20 == 0:
                    print(f"Enriched {count}/{len(resources)} resources...", flush=True)
                # We don't really care about the result as enrich_item updates 'r' in place
                try:
                    future.result(timeout=60)
                except Exception:
                    pass
                    
        print("Enrichment complete.")


    def _guess_cfn_type(self, service, resource_type):
        """
        Map simplified ARN service/type to accurate CloudFormation Type
        using the dynamically loaded cache.
        """
        # Ensure we have the map (lazy load if needed, though we call it earlier)
        if not hasattr(self, 'cfn_type_map'):
            self._ensure_cfn_type_cache()
            
        key = f"{service.lower()}:{resource_type.lower()}"
        
        # 1. Exact match from our smart map
        if key in self.cfn_type_map:
            return self.cfn_type_map[key]
            
        # 2. Fallback: Tagging API often returns 'ec2' for 'EIP' but CFN expects 'AWS::EC2::EIP'.
        # The map should cover this if we keyed it right. 
        # But sometimes tagging ARN type is 'elastic-ip' and CFN is 'EIP'.
        # We need a small manual override only for purely mismatched NAMES, not Services.
        
        manual_overrides = {
            'ec2:elastic-ip': 'AWS::EC2::EIP',
            'ec2:eip': 'AWS::EC2::EIP',
            'ec2:key-pair': 'AWS::EC2::KeyPair',
            'ec2:network-interface': 'AWS::EC2::NetworkInterface',
            'ec2:volume': 'AWS::EC2::Volume',
            'ec2:vpc': 'AWS::EC2::VPC',
            'ec2:subnet': 'AWS::EC2::Subnet',
            'ec2:security-group': 'AWS::EC2::SecurityGroup',
            'ec2:instance': 'AWS::EC2::Instance',
            'rds:db': 'AWS::RDS::DBInstance',
            'rds:snapshot': 'AWS::RDS::DBSnapshot',
            'rds:cluster-snapshot': 'AWS::RDS::DBClusterSnapshot',
            'elasticloadbalancing:listener-rule': 'AWS::ElasticLoadBalancingV2::ListenerRule',
            'acm:certificate': 'AWS::CertificateManager::Certificate'
        }
        if key in manual_overrides:
            return manual_overrides[key]
            
        # 3. Fuzzy search in map?
        # e.g. service='ec2', type='internet-gateway'. CFN='AWS::EC2::InternetGateway'.
        # Map key: 'ec2:internetgateway' (we stripped dashes in map generation? No we didn't).
        # Let's try stripping dashes in key.
        
        # Refined lookup key with dashes removed
        key_nodash = f"{service.lower()}:{resource_type.lower().replace('-', '')}"
        
        # We need our map to support dash-less lookup
        # Iterate cache to find approximate match? (Expensive?)
        # Better: Since we only do this once per resource, it's okay.
        
        for map_key, cfn_val in self.cfn_type_map.items():
            # map_key is "service:resource"
            mk_svc, mk_res = map_key.split(':')
            
            # Match Resource Type (ignoring dashes)
            if mk_res.replace('-', '') == resource_type.lower().replace('-', ''):
                # Match Service (Fuzzy for v2/v3 suffix)
                # e.g. 'elasticloadbalancing' matches 'elasticloadbalancingv2'
                if mk_svc == service.lower() or mk_svc.startswith(service.lower()) or service.lower().startswith(mk_svc):
                     return cfn_val
                 
        return None

    def _extract_identifier(self, arn, id_value):
        # Deprecated by loop strategy in calling function
        return id_value

    def _flatten_dict(self, d, parent_key='', sep='.'):
        """
        Recursively flatten a dictionary.
        """
        items = []
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(self._flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                if not v:
                    items.append((new_key, ""))
                elif all(isinstance(x, dict) for x in v):
                    # Optimized formatting for lists of objects (especially SG rules)
                    lines = []
                    for item in v:
                        # Identify Security Group Rules (CloudControl or Native format)
                        if any(key in item for key in ['IpProtocol', 'FromPort', 'CidrIp', 'IpRanges', 'Ipv6Ranges']):
                            proto = str(item.get('IpProtocol', 'TCP')).upper()
                            if proto == '-1': proto = 'ALL'
                            
                            fp = item.get('FromPort')
                            tp = item.get('ToPort')
                            port = f"{fp}" if fp == tp else f"{fp}-{tp}"
                            if fp == -1 or fp is None: port = "ALL"
                            
                            sources = []
                            # Sources from CloudControl/Simplified
                            if 'CidrIp' in item: sources.append(item['CidrIp'])
                            if 'CidrIpv6' in item: sources.append(item['CidrIpv6'])
                            # Sources from Native Boto3
                            for rs in item.get('IpRanges', []): 
                                d_inner = f" ({rs.get('Description')})" if rs.get('Description') else ""
                                sources.append(f"{rs.get('CidrIp')}{d_inner}")
                            for rs in item.get('Ipv6Ranges', []): 
                                d_inner = f" ({rs.get('Description')})" if rs.get('Description') else ""
                                sources.append(f"{rs.get('CidrIpv6')}{d_inner}")
                            for rs in item.get('UserIdGroupPairs', []): 
                                d_inner = f" ({rs.get('Description')})" if rs.get('Description') else ""
                                sources.append(f"{rs.get('GroupId')}{d_inner}")
                            
                            source_str = ", ".join(filter(None, sources)) or "anywhere"
                            desc = f" | {item.get('Description')}" if item.get('Description') and 'CidrIp' in item else ""
                            lines.append(f"{proto}:{port} from {source_str}{desc}")
                        else:
                            # General dict-list item formatting (e.g. Tags, Subnets)
                            parts = [f"{sk}:{sv}" for sk, sv in item.items() if not isinstance(sv, (dict, list))]
                            lines.append(", ".join(parts))
                    items.append((new_key, "\n".join(lines)))
                else:
                    # Simple list of primitives
                    items.append((new_key, ", ".join(map(str, v))))
            else:
                items.append((new_key, v))
        return dict(items)
        
    # Removed specific _enrich_ec2, _enrich_rds etc. to be dynamic

    def _get_tag_from_dict(self, tags, key):
        for t in tags:
            if t['Key'] == key:
                return t['Value']
        return None





def get_azure_mapping(row):
    """
    Maps AWS resources to Azure Service Name (High Level).
    """
    aws_type = row.get('NormalizedResourceType', 'Unknown')
    details = str(row.get('Details', '')).lower()
    engine = str(row.get('Engine', '')).lower()
    
    svc = 'Assess Manually'
    
    # Compute
    # Compute
    # Compute
    if aws_type == 'AWS::EC2::Instance': svc = 'Virtual Machines'
    elif aws_type == 'AWS::Lambda::Function': svc = 'Azure Functions'
    elif any(k in aws_type for k in ['ECS::Service', 'ECS::Cluster', 'EKS::Cluster']): svc = 'Azure Kubernetes Service (AKS)'
    elif aws_type == 'AWS::AutoScaling::AutoScalingGroup': svc = 'Azure VM Scale Sets'
    elif aws_type == 'AWS::Batch::ComputeEnvironment': svc = 'Azure Batch'
    elif 'LaunchTemplate' in aws_type: svc = 'VM Scale Set / Image Template'
    elif 'Spot' in aws_type: svc = 'Azure Spot Virtual Machines'
    elif 'Fleet' in aws_type: svc = 'Virtual Machine Scale Sets'

    # Storage
    elif aws_type == 'AWS::S3::Bucket': svc = 'Storage Accounts'
    elif aws_type == 'AWS::EFS::FileSystem': svc = 'Azure Files'
    elif aws_type == 'AWS::EC2::Volume': svc = 'Managed Disks'
    elif 'DLM::Policy' in aws_type: svc = 'Azure Backup / Snapshot Policy'

    # Database
    elif aws_type == 'AWS::RDS::DBInstance':
        if 'mysql' in engine: svc = 'Azure Database for MySQL'
        elif 'postgres' in engine: svc = 'Azure Database for PostgreSQL'
        elif 'sql' in engine: svc = 'Azure SQL Managed Instance'
        elif 'mariadb' in engine: svc = 'Azure Database for MariaDB'
        else: svc = 'Azure Database (SQL/MySQL/PostgreSQL)'

    elif aws_type == 'AWS::DynamoDB::Table': svc = 'Azure Cosmos DB'
    elif aws_type == 'AWS::Redshift::Cluster': svc = 'Azure Synapse Analytics'
    elif 'ElastiCache' in aws_type or 'MemoryDB' in aws_type: svc = 'Azure Cache for Redis'

    # Networking
    elif aws_type == 'AWS::EC2::VPC': svc = 'Virtual Network'
    elif aws_type == 'AWS::EC2::Subnet': svc = 'Virtual Network'
    elif aws_type == 'AWS::ElasticLoadBalancingV2::LoadBalancer':
        if 'application' in details: svc = 'Azure Application Gateway'
        else: svc = 'Azure Load Balancer'
    elif 'TargetGroup' in aws_type: svc = 'Azure LB/AppGw Component'
    elif 'ListenerRule' in aws_type: svc = 'Azure AppGw Rule'
    elif aws_type == 'AWS::GlobalAccelerator::Accelerator': svc = 'Azure Traffic Manager'
    elif aws_type == 'AWS::Route53::HostedZone': svc = 'Azure DNS'
    elif aws_type == 'AWS::EC2::SecurityGroup': svc = 'Network Security Groups'
    elif 'SecurityGroup' in aws_type and 'Rule' in aws_type: svc = 'Network Security Group Rule'
    elif aws_type == 'AWS::EC2::EIP': svc = 'Public IP Addresses (Static)'
    elif aws_type == 'AWS::EC2::NetworkInterface': svc = 'Network Interface (NIC)'
    elif aws_type == 'AWS::EC2::RouteTable': svc = 'Route Tables'
    elif aws_type == 'AWS::EC2::InternetGateway': svc = 'N/A (Implicit in VNet)'
    elif aws_type == 'AWS::CloudFront::Distribution': svc = 'Azure Front Door / CDN'
    elif 'servicediscovery' in aws_type.lower(): svc = 'Azure Private DNS'
    elif 'NetworkFirewall' in aws_type: svc = 'Azure Firewall'
    elif 'VPCEndpoint' in aws_type: svc = 'Private Endpoints / Private Link'
    elif 'VpcFlowLog' in aws_type: svc = 'NSG Flow Logs'
    elif 'NetworkAcl' in aws_type: svc = 'Network Security Groups'
    elif 'DHCPOptions' in aws_type: svc = 'VNet DNS Settings'
    elif 'PrefixList' in aws_type: svc = 'IP Groups'
    elif 'NetworkInterfaceAttachment' in aws_type: svc = 'N/A (Implicit in NIC)'
        
    # Security / Identity
    elif any(k in aws_type for k in ['IAM::Role', 'IAM::User', 'IAM::Group', 'SSO::Instance']): svc = 'Microsoft Entra ID'
    elif 'Policy' in aws_type and 'IAM' in aws_type: svc = 'Azure RBAC / Policies'
    elif 'InstanceProfile' in aws_type: svc = 'Managed Identity'
    elif aws_type == 'AWS::KMS::Key': svc = 'Azure Key Vault'
    elif aws_type == 'AWS::CertificateManager::Certificate': svc = 'Azure Key Vault Certificates'
    elif 'KeyPair' in aws_type: svc = 'Azure SSH Keys'
    elif 'SecretsManager' in aws_type: svc = 'Azure Key Vault Secrets'
    elif 'WAF' in aws_type: svc = 'Azure WAF'
    elif 'VirtualMFADevice' in aws_type: svc = 'Entra ID MFA'

    # Analytics / Integration
    elif 'Kinesis' in aws_type: svc = 'Azure Event Hubs'
    elif aws_type == 'AWS::SNS::Topic' or aws_type == 'AWS::SQS::Queue': svc = 'Azure Service Bus'
    elif 'Events::Rule' in aws_type or 'EventBridge::Rule' in aws_type: svc = 'Azure Event Grid'
    elif aws_type == 'AWS::CloudWatch::Alarm': svc = 'Azure Monitor Alerts'
    elif 'Athena' in aws_type: svc = 'Azure Synapse / Databricks'
    elif 'Glue' in aws_type: svc = 'Azure Data Factory / Databricks'
    elif 'StepFunctions' in aws_type: svc = 'Azure Logic Apps'

    # Containers
    elif 'ECR' in aws_type: svc = 'Azure Container Registry (ACR)'
    elif 'ECS' in aws_type: svc = 'Azure Container Apps / AKS'

    # Operations / Others
    elif 'snapshot' in aws_type.lower(): svc = 'Azure Snapshots'
    elif 'CapacityProvider' in aws_type: svc = 'AKS Node Pools'
    elif 'TaskDefinition' in aws_type: svc = 'Container App Revisions'
    elif 'SSM::Session' in aws_type: svc = 'N/A (Ephemeral)'
    elif any(k in aws_type for k in ['StaticIp', 'ElasticIP']): svc = 'Public IP Addresses (Static)'
    elif any(k in aws_type for k in ['MediaLive', 'MediaConvert', 'MediaStore']): svc = 'Azure Media Services'
    elif 'CloudFormation' in aws_type: svc = 'Azure Resource Manager (ARM)'
    elif 'CloudTrail' in aws_type: svc = 'Azure Activity Log'
    elif 'Config' in aws_type and 'AWS::' in aws_type: svc = 'Azure Policy'
    elif 'AppConfig' in aws_type: svc = 'Azure App Configuration'
    elif 'DataSync' in aws_type: svc = 'Azure Data Box / Storage Mover'
    elif any(k in aws_type for k in ['CodePipeline', 'CodeBuild', 'CodeDeploy']): svc = 'Azure DevOps / GitHub Actions'
    elif 'mgn' in aws_type.lower() or 'SourceServer' in aws_type: svc = 'Azure Migrate Source Server'
    elif 'Alias' in aws_type: svc = 'Azure Key Vault (Alias)'

    # =====================================================================
    # AI / ML services — AWS to Azure mappings.
    # Pattern: most AWS AI services map to a specific Azure AI offering;
    # SageMaker sub-types map to specific Azure Machine Learning concepts.
    # =====================================================================

    # SageMaker → Azure Machine Learning (the umbrella platform).
    elif aws_type == 'AWS::SageMaker::NotebookInstance':
        svc = 'Azure Machine Learning Compute Instance'
    elif aws_type == 'AWS::SageMaker::Endpoint':
        svc = 'Azure Machine Learning Online Endpoint'
    elif aws_type == 'AWS::SageMaker::Model':
        svc = 'Azure Machine Learning Model Registry'
    elif aws_type == 'AWS::SageMaker::Domain':
        svc = 'Azure Machine Learning Workspace (Studio)'
    elif aws_type == 'AWS::SageMaker::Pipeline':
        svc = 'Azure Machine Learning Pipelines'
    elif aws_type == 'AWS::SageMaker::FeatureGroup':
        svc = 'Azure Machine Learning Feature Store'
    elif aws_type in ('AWS::SageMaker::HumanTaskUi', 'AWS::SageMaker::FlowDefinition'):
        svc = 'Azure Machine Learning Data Labeling'
    elif 'SageMaker' in aws_type:
        svc = 'Azure Machine Learning'

    # Bedrock → Azure OpenAI Service / Azure AI Foundry.
    elif aws_type == 'AWS::Bedrock::CustomModel':
        svc = 'Azure OpenAI Fine-tuned Model'
    elif aws_type == 'AWS::Bedrock::Guardrail':
        svc = 'Azure AI Content Safety'
    elif aws_type == 'AWS::Bedrock::ProvisionedModelThroughput':
        svc = 'Azure OpenAI Provisioned Throughput (PTU)'
    elif aws_type == 'AWS::Bedrock::KnowledgeBase':
        svc = 'Azure AI Search + Azure OpenAI On Your Data'
    elif aws_type == 'AWS::Bedrock::Agent':
        svc = 'Azure AI Foundry Agents / Copilot Studio'
    elif aws_type == 'AWS::Bedrock::Prompt':
        svc = 'Azure AI Foundry Prompt Flow'
    elif 'Bedrock' in aws_type:
        svc = 'Azure OpenAI Service / Azure AI Foundry'

    # Comprehend → Azure AI Language.
    elif aws_type == 'AWS::Comprehend::DocumentClassifier':
        svc = 'Azure AI Language — Custom Text Classification'
    elif aws_type == 'AWS::Comprehend::EntityRecognizer':
        svc = 'Azure AI Language — Custom Named Entity Recognition'
    elif aws_type == 'AWS::Comprehend::Endpoint':
        svc = 'Azure AI Language — Deployed Model'
    elif 'Comprehend' in aws_type:
        svc = 'Azure AI Language'

    # Rekognition → Azure AI Vision.
    elif aws_type == 'AWS::Rekognition::Collection':
        svc = 'Azure AI Vision — Face API'
    elif aws_type == 'AWS::Rekognition::StreamProcessor':
        svc = 'Azure AI Video Indexer'
    elif aws_type == 'AWS::Rekognition::Project':
        svc = 'Azure AI Custom Vision'
    elif 'Rekognition' in aws_type:
        svc = 'Azure AI Vision'

    # Polly → Azure AI Speech (TTS).
    elif aws_type == 'AWS::Polly::Lexicon':
        svc = 'Azure AI Speech — Custom Lexicon'
    elif 'Polly' in aws_type:
        svc = 'Azure AI Speech (Text-to-Speech)'

    # Transcribe → Azure AI Speech (STT).
    elif aws_type == 'AWS::Transcribe::Vocabulary':
        svc = 'Azure AI Speech — Custom Speech Vocabulary'
    elif aws_type == 'AWS::Transcribe::TranscriptionJob':
        svc = 'Azure AI Speech — Batch Transcription'
    elif aws_type == 'AWS::Transcribe::LanguageModel':
        svc = 'Azure AI Speech — Custom Speech Model'
    elif aws_type == 'AWS::Transcribe::VocabularyFilter':
        svc = 'Azure AI Speech — Profanity Filter'
    elif 'Transcribe' in aws_type:
        svc = 'Azure AI Speech (Speech-to-Text)'

    # Translate → Azure AI Translator.
    elif aws_type in ('AWS::Translate::Terminology', 'AWS::Translate::ParallelData'):
        svc = 'Azure AI Translator — Custom Translator'
    elif 'Translate' in aws_type:
        svc = 'Azure AI Translator'

    # Lex → Azure Bot Service / Copilot Studio.
    elif aws_type == 'AWS::Lex::Bot' or 'Lex::' in aws_type:
        svc = 'Azure Bot Service / Microsoft Copilot Studio'

    # Kendra → Azure AI Search.
    elif aws_type == 'AWS::Kendra::Index' or 'Kendra' in aws_type:
        svc = 'Azure AI Search (Cognitive Search)'

    # Textract → Azure AI Document Intelligence.
    elif aws_type == 'AWS::Textract::Adapter':
        svc = 'Azure AI Document Intelligence — Custom Model'
    elif 'Textract' in aws_type:
        svc = 'Azure AI Document Intelligence (Form Recognizer)'

    # Personalize → Azure ML Recommendations (Personalizer being retired).
    elif aws_type in ('AWS::Personalize::DatasetGroup', 'AWS::Personalize::Solution', 'AWS::Personalize::Campaign'):
        svc = 'Azure Machine Learning — Recommendations / Personalizer (retired)'
    elif 'Personalize' in aws_type:
        svc = 'Azure Machine Learning — Recommendations'

    # Forecast → Azure ML AutoML Forecasting.
    elif aws_type in ('AWS::Forecast::Dataset', 'AWS::Forecast::Predictor', 'AWS::Forecast::Forecast') or 'Forecast' in aws_type:
        svc = 'Azure Machine Learning — AutoML Forecasting'

    # Lookout family → Azure AI Anomaly Detector / Custom Vision.
    elif 'LookoutEquipment' in aws_type:
        svc = 'Azure AI Anomaly Detector (Multivariate)'
    elif 'LookoutMetrics' in aws_type:
        svc = 'Azure AI Anomaly Detector (Multivariate)'
    elif 'LookoutVision' in aws_type:
        svc = 'Azure AI Custom Vision — Anomaly Detection'

    # DevOps Guru → Azure Monitor / App Insights Smart Detection.
    elif 'DevOpsGuru' in aws_type:
        svc = 'Azure Application Insights — Smart Detection'

    # CodeGuru → Azure DevOps tooling.
    elif 'CodeGuruReviewer' in aws_type:
        svc = 'GitHub Advanced Security / SonarCloud (Azure DevOps)'
    elif 'CodeGuruProfiler' in aws_type:
        svc = 'Azure Application Insights — Profiler'
    elif 'CodeGuru' in aws_type:
        svc = 'Azure DevOps Code Quality Tools'

    # Fraud Detector → Azure AI Anomaly Detector / Custom ML.
    elif 'FraudDetector' in aws_type:
        svc = 'Azure AI Anomaly Detector / Azure ML Custom Model'

    # HealthLake → Azure Health Data Services (FHIR).
    elif 'HealthLake' in aws_type:
        svc = 'Azure Health Data Services (FHIR)'

    # HealthOmics → Azure Genomics.
    elif aws_type in ('AWS::Omics::SequenceStore', 'AWS::Omics::ReferenceStore'):
        svc = 'Azure Genomics — Data Storage'
    elif aws_type == 'AWS::Omics::Workflow':
        svc = 'Azure Genomics / Azure ML Pipelines'
    elif 'Omics' in aws_type:
        svc = 'Azure Genomics'

    # Amazon Q Business → Microsoft Copilot / Azure AI Foundry.
    elif aws_type == 'AWS::QBusiness::Application' or 'QBusiness' in aws_type or 'Q::' in aws_type:
        svc = 'Microsoft 365 Copilot / Copilot Studio'

    # Augmented AI runtime (A2I) — caught by SageMaker patterns above, but
    # in case the type comes in differently:
    elif 'A2I' in aws_type or 'augmented' in aws_type.lower():
        svc = 'Azure Machine Learning Data Labeling'

    return svc

def generate_aws_report(resources, duration=None):
    """
    Generate Excel report from extracted AWS resources.
    Follows STRICT rules:
    1. Identity = ResourceId (Deduplication Key)
    2. Type = NormalizedResourceType (Grouping Key)
    3. Name = Friendly Name (Presentation Only)
    """
    import pandas as pd
    import os
    from datetime import datetime
    from collections import defaultdict

    # Generate path with UUID to prevent collisions in multi-user environments
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_id = str(uuid.uuid4())[:8]
    filename = f"AWS_Assessment_{timestamp}_{unique_id}.xlsx"
    filepath = os.path.join(os.path.dirname(__file__), "..", "data", filename)
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    if not resources:
        df = pd.DataFrame(columns=['ResourceId', 'ResourceName', 'NormalizedResourceType', 'ResourceType', 'ResourceGroup', 'Location', 'Status', 'Details', 'Tags'])
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Summary', index=False)
            df.to_excel(writer, sheet_name='All Resources', index=False)
        return filepath

    # 0. STRICT Global Deduplication
    # We must normalize the ID column to handle mixed ARNs and IDs
    # E.g. "arn:aws:ec2:us-east-1:123:eni/eni-123" vs "eni-123"
    
    seen_ids = set()
    unique_resources = []
    
    # Debug: Check p-mysql before dedup
    for r in resources:
        if 'p-mysql' in r.get('ResourceName', '') or 'p-mysql' in r.get('ResourceId', ''):
             print(f"DEBUG PRE-DEDUP: p-mysql found. ID={r.get('ResourceId')} Type={r.get('NormalizedResourceType')}")

    for r in resources:
        # Determine the "Core ID"
        rid = r.get('ResourceId', '')
        core_id = rid
        
        # If valid ARN, extract the last valid ID part
        # Logic: Normalizing separators (:) to (/) and taking the last segment handles
        # most AWS ARN formats correctly for ID extraction:
        # - arn:aws:ec2:...:instance/i-123 -> i-123
        # - arn:aws:rds:...:db:my-db -> my-db
        # - arn:aws:s3:::my-bucket -> my-bucket
        if rid.startswith('arn:'):
            # Special handling for S3 to preserve bucket name exactly if needed, 
            # but generally split works.
            clean_rid = rid.replace(':', '/')
            core_id = clean_rid.split('/')[-1]
            
            # Fallback for empty (trailing slash?)
            if not core_id:
                core_id = clean_rid.split('/')[-2]
                
        # Clean core_id further? e.g. if it somehow still had path
        if '/' in core_id:
             core_id = core_id.split('/')[-1]
             
        if core_id not in seen_ids:
            seen_ids.add(core_id)
            seen_ids.add(rid) 
            # Handle standard AWS ID variations
            if core_id.startswith('eni-'): seen_ids.add(f"network-interface/{core_id}")
            if core_id.startswith('eipalloc-'): seen_ids.add(f"allocation-id/{core_id}")
            
            unique_resources.append(r)
        else:
             if 'p-mysql' in rid:
                 print(f"DEBUG: DROPPED p-mysql due to duplicate CoreID: {core_id}")
            
    df = pd.DataFrame(unique_resources)
    
    # Drop columns where ALL values are effectively empty: [], "[]", None, "", "{}", "None", "nan", "Unknown"
    def is_empty_val(v):
        if v is None: return True
        # Handle pandas NaN
        if isinstance(v, float) and pd.isna(v): return True
        if isinstance(v, list) and not v: return True
        if isinstance(v, dict) and not v: return True
        sv = str(v).strip().lower()
        if sv in ["", "[]", "{}", "none", "nan", "unknown"]: return True
        return False

    cols_to_drop = []
    for col in df.columns:
        if df[col].apply(is_empty_val).all():
            cols_to_drop.append(col)
    
    # Don't drop base columns even if empty, to maintain structure
    protected_cols = ['ResourceId', 'ResourceName', 'NormalizedResourceType', 'ResourceGroup', 'Location', 'Status', 'Connectivity', 'Ingress Rules', 'Egress Rules', 'Details', 'Tags']
    cols_to_drop = [c for c in cols_to_drop if c not in protected_cols]
    
    if cols_to_drop:
        print(f"Dropping {len(cols_to_drop)} empty columns: {cols_to_drop}")
        df.drop(columns=cols_to_drop, inplace=True)

    # Remove duplicate columns if they exist (fixes the repetitive header bug)
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # 1. Deduplicate by ResourceId (Redundant but safe)

    # Ensure clean columns
    # Ensure clean columns and promote 'Connectivity' to the top for visibility
    base_cols = [
        'ResourceId', 'ResourceName', 'Connectivity', 'NormalizedResourceType', 'ResourceType',
        'ResourceGroup', 'Location', 'Status', 'Ingress Rules', 'Egress Rules', 'Details', 'Tags'
    ]
    for col in base_cols:
        if col not in df.columns:
            df[col] = ''

    # Drop internal-only columns from EVERY sheet (Summary, All Resources, and
    # type-specific). ResourceTypeSource is a discovery-source marker
    # ("NativeScan" / "Native" / "DynamicDiscovery") used internally and
    # has no user-facing value.
    _internal_cols = ['ResourceTypeSource']
    df.drop(columns=[c for c in _internal_cols if c in df.columns], inplace=True, errors='ignore')
    
    # Sanitize datetimes
    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)
        elif df[col].dtype == 'object':
            df[col] = df[col].apply(lambda x: x.replace(tzinfo=None) if isinstance(x, datetime) else x)

    def get_friendly_name(resource_type):
        mapping = {
            'AWS::EC2::Instance': 'EC2 Instances',
            'AWS::EC2::Volume': 'EBS Volumes',
            'AWS::RDS::DBInstance': 'RDS Databases',
            'AWS::S3::Bucket': 'S3 Buckets',
            'AWS::Lambda::Function': 'Lambda Functions',
            'AWS::EC2::VPC': 'VPCs',
            'AWS::EC2::SecurityGroup': 'Security Groups',
            'AWS::ElasticLoadBalancingV2::LoadBalancer': 'Load Balancers',
            'AWS::DynamoDB::Table': 'DynamoDB Tables',
            'AWS::RDS::DBCluster': 'RDS clusters', 
            'AWS::RDS::DBSnapshot': 'RDS Snapshots',
            'AWS::RDS::DBClusterSnapshot': 'RDS Cluster Snapshots',
            'AWS::EC2::EIP': 'EC2 Elastic IPs',
            'AWS::EC2::NetworkInterface': 'EC2 Network Interfaces',
            'AWS::EC2::Subnet': 'EC2 Subnets',
            'AWS::EC2::RouteTable': 'EC2 Route Tables',
            'AWS::EC2::InternetGateway': 'EC2 Internet Gateways',
            'AWS::EC2::NatGateway': 'EC2 NAT Gateways',
            'AWS::EC2::VPCGatewayAttachment': 'EC2 Gateway Attachments',
            'AWS::EC2::VPCEndpoint': 'EC2 VPC Endpoints',
            'AWS::EC2::LaunchTemplate': 'EC2 Launch Templates',
            'AWS::EC2::SecurityGroupIngress': 'EC2 Security Group Rules',
            'AWS::EC2::SecurityGroupEgress': 'EC2 Security Group Rules',
            'AWS::IAM::Role': 'IAM Roles',
            'AWS::IAM::User': 'IAM Users',
            'AWS::IAM::Policy': 'IAM Policies',
            'AWS::CloudWatch::Alarm': 'CloudWatch Alarms',
            'AWS::Logs::LogGroup': 'CloudWatch Log Groups',
            'AWS::SNS::Topic': 'SNS Topics',
            'AWS::SQS::Queue': 'SQS Queues',
            'AWS::KMS::Key': 'KMS Keys',
            'AWS::CertificateManager::Certificate': 'ACM Certificates',
            'AWS::ApiGateway::RestApi': 'API Gateway REST APIs',
            'AWS::ApiGatewayV2::Api': 'API Gateway HTTP/WebSocket APIs',
            'AWS::CloudFront::Distribution': 'CloudFront Distributions',
            'AWS::AutoScaling::AutoScalingGroup': 'Auto Scaling Groups',
            'AWS::ElasticLoadBalancing::LoadBalancer': 'Classic Load Balancers',
            'AWS::ElasticLoadBalancingV2::TargetGroup': 'Load Balancer Target Groups',
            'AWS::ECR::Repository': 'ECR Repositories',
            'AWS::ECS::Cluster': 'ECS Clusters',
            'AWS::ECS::Service': 'ECS Services',
            'AWS::ECS::TaskDefinition': 'ECS Task Definitions',
            'AWS::EKS::Cluster': 'EKS Clusters',
            'AWS::EFS::FileSystem': 'EFS File Systems',
            'AWS::ElastiCache::CacheCluster': 'ElastiCache Clusters',
            'AWS::ElastiCache::ReplicationGroup': 'ElastiCache Replication Groups',
            'AWS::Redshift::Cluster': 'Redshift Clusters',
            'AWS::OpenSearchService::Domain': 'OpenSearch Domains',
            'AWS::MSK::Cluster': 'MSK Clusters',
            'AWS::AppSync::GraphQLApi': 'AppSync GraphQL APIs',
            'AWS::Glue::Job': 'Glue Jobs',
            'AWS::Glue::Crawler': 'Glue Crawlers',
            'AWS::DataPipeline::Pipeline': 'Data Pipelines',
            'AWS::StepFunctions::StateMachine': 'Step Functions State Machines',
            'AWS::Events::Rule': 'EventBridge Rules',
            'AWS::Config::ConfigurationRecorder': 'Config Recorders',
            'AWS::Config::DeliveryChannel': 'Config Delivery Channels',
            'AWS::Config::ConfigRule': 'Config Rules',
            'AWS::WAFv2::WebACL': 'WAFv2 Web ACLs',
            'AWS::Shield::Protection': 'Shield Protections',
            'AWS::Route53::HostedZone': 'Route 53 Hosted Zones',
            'AWS::Route53::RecordSet': 'Route 53 Record Sets',
            'AWS::CloudFormation::Stack': 'CloudFormation Stacks',
            'AWS::CodeBuild::Project': 'CodeBuild Projects',
            'AWS::CodePipeline::Pipeline': 'CodePipeline Pipelines',
            'AWS::CodeDeploy::Application': 'CodeDeploy Applications',
            'AWS::CodeCommit::Repository': 'CodeCommit Repositories',
            'AWS::SSM::Parameter': 'SSM Parameters',
            'AWS::SSM::Document': 'SSM Documents',
            'AWS::SecretsManager::Secret': 'Secrets Manager Secrets',
            'AWS::AppRunner::Service': 'App Runner Services',
            'AWS::Kinesis::Stream': 'Kinesis Streams',
            'AWS::KinesisFirehose::DeliveryStream': 'Kinesis Firehose Delivery Streams',
            'AWS::Cognito::UserPool': 'Cognito User Pools',
            'AWS::Cognito::UserPoolClient': 'Cognito User Pool Clients',
            'AWS::Amplify::App': 'Amplify Apps',
            'AWS::AppConfig::Application': 'AppConfig Applications',
            'AWS::AppFlow::Flow': 'AppFlow Flows',
            'AWS::AppStream::Fleet': 'AppStream Fleets',
            'AWS::Athena::WorkGroup': 'Athena WorkGroups',
            'AWS::Backup::BackupVault': 'Backup Vaults',
            'AWS::Batch::ComputeEnvironment': 'Batch Compute Environments',
            'AWS::Batch::JobQueue': 'Batch Job Queues',
            'AWS::BillingConductor::BillingGroup': 'Billing Groups',
            'AWS::Cassandra::Keyspace': 'Cassandra Keyspaces',
            'AWS::Cloud9::EnvironmentEC2': 'Cloud9 Environments',
            'AWS::CloudTrail::Trail': 'CloudTrail Trails',
            'AWS::CodeGuruReviewer::RepositoryAssociation': 'CodeGuru Reviewer Associations',
            'AWS::Comprehend::DocumentClassifier': 'Comprehend Classifiers',
            'AWS::Connect::Instance': 'Connect Instances',
            'AWS::DataSync::Agent': 'DataSync Agents',
            'AWS::DMS::ReplicationInstance': 'DMS Replication Instances',
            'AWS::DocDB::DBInstance': 'DocumentDB Instances',
            'AWS::EC2::ClientVpnEndpoint': 'EC2 Client VPN Endpoints',
            'AWS::EC2::CustomerGateway': 'EC2 Customer Gateways',
            'AWS::EC2::TransitGateway': 'EC2 Transit Gateways',
            'AWS::EC2::TransitGatewayAttachment': 'EC2 Transit Gateway Attachments',
            'AWS::EC2::VpnConnection': 'EC2 VPN Connections',
            'AWS::EC2::VpnGateway': 'EC2 VPN Gateways',
            'AWS::EKS::FargateProfile': 'EKS Fargate Profiles',
            'AWS::EMR::Cluster': 'EMR Clusters',
            'AWS::Evidently::Project': 'Evidently Projects',
            'AWS::FSx::FileSystem': 'FSx File Systems',
            'AWS::GameLift::Build': 'GameLift Builds',
            'AWS::GlobalAccelerator::Accelerator': 'Global Accelerator Accelerators',
            'AWS::GreengrassV2::ComponentVersion': 'GreengrassV2 Components',
            'AWS::GroundStation::Config': 'Ground Station Configs',
            'AWS::GuardDuty::Detector': 'GuardDuty Detectors',
            'AWS::ImageBuilder::ImagePipeline': 'Image Builder Pipelines',
            'AWS::IoT::Thing': 'IoT Things',
            'AWS::IoT::TopicRule': 'IoT Topic Rules',
            'AWS::IoTAnalytics::Channel': 'IoT Analytics Channels',
            'AWS::IoTCoreDeviceAdvisor::SuiteDefinition': 'IoT Device Advisor Suite Definitions',
            'AWS::IoTWireless::ServiceProfile': 'IoT Wireless Service Profiles',
            'AWS::IVS::Channel': 'IVS Channels',
            'AWS::Kendra::Index': 'Kendra Indexes',
            'AWS::LakeFormation::DataLakeSettings': 'Lake Formation Settings',
            'AWS::Lex::Bot': 'Lex Bots',
            'AWS::LicenseManager::LicenseConfiguration': 'License Configurations',
            'AWS::Lightsail::Instance': 'Lightsail Instances',
            'AWS::Location::Tracker': 'Location Trackers',
            'AWS::LookoutMetrics::AnomalyDetector': 'LookoutMetrics Detectors',
            'AWS::M2::Application': 'Mainframe Modernization Applications',
            'AWS::ManagedBlockchain::Node': 'Managed Blockchain Nodes',
            'AWS::MediaConnect::Flow': 'MediaConnect Flows',
            'AWS::MediaLive::Channel': 'MediaLive Channels',
            'AWS::MediaPackage::Channel': 'MediaPackage Channels',
            'AWS::MediaTailor::PlaybackConfiguration': 'MediaTailor Playback Configs',
            'AWS::MemoryDB::Cluster': 'MemoryDB Clusters',
            'AWS::Neptune::DBCluster': 'Neptune DB Clusters',
            'AWS::NetworkFirewall::Firewall': 'Network Firewall Firewalls',
            'AWS::NetworkManager::GlobalNetwork': 'Network Manager Global Networks',
            'AWS::NimbleStudio::Studio': 'Nimble Studio Studios',
            'AWS::OpenSearchService::Domain': 'OpenSearch Domains', # Already there, but good to keep
            'AWS::OpsWorks::Stack': 'OpsWorks Stacks',
            'AWS::Panorama::ApplicationInstance': 'Panorama Application Instances',
            'AWS::Pinpoint::App': 'Pinpoint Apps',
            'AWS::Proton::ServiceTemplate': 'Proton Service Templates',
            'AWS::QLDB::Ledger': 'QLDB Ledgers',
            'AWS::RAM::ResourceShare': 'RAM Resource Shares',
            'AWS::RUM::AppMonitor': 'RUM App Monitors',
            'AWS::RedshiftServerless::Workgroup': 'Redshift Serverless Workgroups',
            'AWS::Rekognition::Collection': 'Rekognition Collections',
            'AWS::ResilienceHub::App': 'Resilience Hub Apps',
            'AWS::ResourceExplorer2::View': 'Resource Explorer Views',
            'AWS::RoboMaker::RobotApplication': 'RoboMaker Robot Applications',
            'AWS::Route53Resolver::ResolverEndpoint': 'Route 53 Resolver Endpoints',
            'AWS::SageMaker::Domain': 'SageMaker Domains',
            'AWS::Scheduler::Schedule': 'Scheduler Schedules',
            'AWS::ServiceCatalog::CloudFormationProduct': 'Service Catalog Products',
            'AWS::ServiceDiscovery::Service': 'Service Discovery Services',
            'AWS::Signer::SigningProfile': 'Signer Signing Profiles',
            'AWS::SimSpaceWeaver::Simulation': 'SimSpace Weaver Simulations',
            'AWS::Transfer::Server': 'Transfer Family Servers',
            'AWS::VerifiedPermissions::PolicyStore': 'Verified Permissions Policy Stores',
            'AWS::VoiceID::Domain': 'Voice ID Domains',
            'AWS::WAF::WebACL': 'WAF Classic Web ACLs',
            'AWS::WorkSpaces::Workspace': 'WorkSpaces',
            'AWS::XRay::Group': 'X-Ray Groups',
            # Native-scan additions (use these exact type names in
            # _scan_native_essentials so sheet labels stay consistent).
            'AWS::EC2::EgressOnlyInternetGateway': 'EC2 Egress-Only Internet Gateways',
            'AWS::EC2::NetworkAcl': 'EC2 Network ACLs',
            'AWS::EC2::VPCPeeringConnection': 'EC2 VPC Peering Connections',
            'AWS::EC2::Snapshot': 'EBS Snapshots',
            'AWS::EC2::Image': 'AMIs',
            'AWS::EC2::VPNGateway': 'EC2 VPN Gateways',
            'AWS::EC2::VPNConnection': 'EC2 VPN Connections',
            'AWS::Glue::Database': 'Glue Databases',
            'AWS::IAM::Group': 'IAM Groups',
            'AWS::IAM::ManagedPolicy': 'IAM Managed Policies',
            'AWS::Backup::BackupPlan': 'Backup Plans',
            'AWS::DirectConnect::Connection': 'Direct Connect Connections',
            # AI/ML services
            'AWS::SageMaker::NotebookInstance': 'SageMaker Notebooks',
            'AWS::SageMaker::Endpoint': 'SageMaker Endpoints',
            'AWS::SageMaker::Model': 'SageMaker Models',
            'AWS::SageMaker::Domain': 'SageMaker Studio Domains',
            'AWS::SageMaker::Pipeline': 'SageMaker Pipelines',
            'AWS::SageMaker::FeatureGroup': 'SageMaker Feature Groups',
            'AWS::SageMaker::HumanTaskUi': 'SageMaker Human Task UIs (A2I)',
            'AWS::SageMaker::FlowDefinition': 'SageMaker Flow Definitions (A2I)',
            'AWS::Bedrock::CustomModel': 'Bedrock Custom Models',
            'AWS::Bedrock::Guardrail': 'Bedrock Guardrails',
            'AWS::Bedrock::ProvisionedModelThroughput': 'Bedrock Provisioned Throughputs',
            'AWS::Bedrock::KnowledgeBase': 'Bedrock Knowledge Bases',
            'AWS::Bedrock::Agent': 'Bedrock Agents',
            'AWS::Bedrock::Prompt': 'Bedrock Prompts',
            'AWS::Comprehend::DocumentClassifier': 'Comprehend Document Classifiers',
            'AWS::Comprehend::EntityRecognizer': 'Comprehend Entity Recognizers',
            'AWS::Comprehend::Endpoint': 'Comprehend Endpoints',
            'AWS::Rekognition::Collection': 'Rekognition Collections',
            'AWS::Rekognition::StreamProcessor': 'Rekognition Stream Processors',
            'AWS::Rekognition::Project': 'Rekognition Projects',
            'AWS::Polly::Lexicon': 'Polly Lexicons',
            'AWS::Transcribe::Vocabulary': 'Transcribe Vocabularies',
            'AWS::Transcribe::TranscriptionJob': 'Transcribe Transcription Jobs',
            'AWS::Transcribe::LanguageModel': 'Transcribe Language Models',
            'AWS::Transcribe::VocabularyFilter': 'Transcribe Vocabulary Filters',
            'AWS::Translate::Terminology': 'Translate Terminologies',
            'AWS::Translate::ParallelData': 'Translate Parallel Data',
            'AWS::Textract::Adapter': 'Textract Adapters',
            'AWS::Personalize::DatasetGroup': 'Personalize Dataset Groups',
            'AWS::Personalize::Solution': 'Personalize Solutions',
            'AWS::Personalize::Campaign': 'Personalize Campaigns',
            'AWS::Forecast::Dataset': 'Forecast Datasets',
            'AWS::Forecast::Predictor': 'Forecast Predictors',
            'AWS::Forecast::Forecast': 'Forecast Forecasts',
            'AWS::LookoutEquipment::Dataset': 'Lookout for Equipment Datasets',
            'AWS::LookoutEquipment::Model': 'Lookout for Equipment Models',
            'AWS::LookoutMetrics::AnomalyDetector': 'Lookout for Metrics Detectors',
            'AWS::LookoutVision::Project': 'Lookout for Vision Projects',
            'AWS::DevOpsGuru::ResourceCollection': 'DevOps Guru Resource Collections',
            'AWS::CodeGuruProfiler::ProfilingGroup': 'CodeGuru Profiler Groups',
            'AWS::FraudDetector::Detector': 'Fraud Detector Detectors',
            'AWS::FraudDetector::Model': 'Fraud Detector Models',
            'AWS::HealthLake::FHIRDatastore': 'HealthLake FHIR Datastores',
            'AWS::Omics::SequenceStore': 'HealthOmics Sequence Stores',
            'AWS::Omics::ReferenceStore': 'HealthOmics Reference Stores',
            'AWS::Omics::Workflow': 'HealthOmics Workflows',
            'AWS::QBusiness::Application': 'Amazon Q Business Applications',
        }
        
        # Exact match
        if resource_type in mapping:
            return mapping[resource_type]
            
        # Parse CFN type
        parts = resource_type.split('::')
        if len(parts) >= 3:
             # Default to Service Type
             return f"{parts[1]} {parts[2]}".replace('V2', '').strip()
        return resource_type

    # Create Excel Writer
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        
        # --- Formatting Helper ---
        fill_aws = PatternFill(start_color="BE5014", end_color="BE5014", fill_type="solid")
        fill_azure = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        def format_sheet(ws_name):
            if ws_name not in writer.sheets: return
            ws = writer.sheets[ws_name]
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                
                # Header
                header = col[0]
                header.font = font_white
                header.border = thin_border
                if str(header.value).startswith('Azure'):
                    header.fill = fill_azure
                else:
                    header.fill = fill_aws
                
                # Data Cells
                for cell in col[1:]:
                    cell.border = thin_border
                    try:
                        val = str(cell.value)
                        if len(val) > max_length: max_length = len(val)
                    except: pass
                
                # Header length
                if len(str(header.value)) > max_length: max_length = len(str(header.value))
                
                # Cap width
                ws.column_dimensions[column].width = min(max_length + 2, 60)

        
        # 1. Summary Sheet
        # Group by NormalizedResourceType
        summary = df.groupby('NormalizedResourceType').size().reset_index(name='Count')

        # SecurityGroupIngress / SecurityGroupEgress are *rules*, not standalone
        # resources from the user's perspective — they're already exploded into
        # rule-rows inside the Security Groups sheet by the type-specific
        # handler. Counting them again at the top level creates 3 duplicate
        # rows in the Summary ("EC2 Security Group Rules" twice + "Security
        # Groups" once) that don't match the single SG sheet. Drop them here.
        rule_types = {
            'AWS::EC2::SecurityGroupIngress',
            'AWS::EC2::SecurityGroupEgress',
        }
        summary = summary[~summary['NormalizedResourceType'].isin(rule_types)].reset_index(drop=True)

        # Add Friendly Name
        summary['Friendly Name'] = summary['NormalizedResourceType'].apply(get_friendly_name)

        # Add Azure Service Mapping (Representative)
        def get_svc_rep(rtype):
            rows = df[df['NormalizedResourceType'] == rtype]
            if not rows.empty:
                return get_azure_mapping(rows.iloc[0])
            return ''

        summary['Azure Service'] = summary['NormalizedResourceType'].apply(get_svc_rep)

        # If two different NormalizedResourceTypes happen to share the same
        # Friendly Name + Azure Service (defensive — e.g., a future renaming
        # could collide), collapse them so the Summary never shows duplicate
        # rows with the same label.
        summary = (
            summary.groupby(['Friendly Name', 'Azure Service'], as_index=False)['Count'].sum()
        )

        # Sort
        summary.sort_values('Count', ascending=False, inplace=True)
        # Reorder
        summary = summary[['Friendly Name', 'Azure Service', 'Count']]

        summary.to_excel(writer, sheet_name='Summary', index=False)
        
        # Add Duration & Regions info if provided
        ws_summary = writer.sheets['Summary']
        last_row = summary.shape[0] + 3
        
        # Duration row
        if duration:
            ws_summary.cell(row=last_row, column=1, value="Assessment Time")
            ws_summary.cell(row=last_row, column=2, value=duration)
            ws_summary.cell(row=last_row, column=1).font = Font(bold=True)
            last_row += 1
            
        # Regions row
        unique_regions = sorted(df['Location'].unique())
        ws_summary.cell(row=last_row, column=1, value="Regions Scanned")
        ws_summary.cell(row=last_row, column=2, value=", ".join(unique_regions))
        ws_summary.cell(row=last_row, column=1).font = Font(bold=True)

        format_sheet('Summary')
        
        # 2. All Resources Sheet
        # Define Global Drop List (Used here and in individual sheets)
        drop_cols_global = [
            'ResourceId', 'NormalizedResourceType', 'ResourceTypeSource', 'ResourceGroup',
            'DisableApiTermination', 'EbsOptimized', 'ImageId', 'InstanceId', 'InstanceInitiatedShutdownBehavior',
            'KeyName', 'PrivateIp', 'Arn', 'ResourceArn', 'CreationDate', 'CreateTime', 'LaunchTime',
            'AllocationId', 'AssociationId', 'NetworkInterfaceOwnerId', 'PublicIpv4Pool'
        ]
        
        # Filter base_cols for All Resources sheet
        master_cols = [c for c in base_cols if c in df.columns and c not in drop_cols_global]
        master_df = df[master_cols].copy()
        
        # Aggressively drop empty columns in Master sheet
        master_cols_to_drop = []
        for col in master_df.columns:
            if master_df[col].apply(is_empty_val).all() and col not in protected_cols:
                master_cols_to_drop.append(col)
        if master_cols_to_drop:
            master_df.drop(columns=master_cols_to_drop, inplace=True)

        # Sort (safe check if col exists)
        sort_keys = [c for c in ['ResourceGroup', 'NormalizedResourceType'] if c in master_df.columns]
        if sort_keys:
             master_df.sort_values(sort_keys, inplace=True)
             
        master_df.to_excel(writer, sheet_name='All Resources', index=False)
        format_sheet('All Resources')

        # 3. Type-Specific Sheets (One per NormalizedResourceType)
        # 3. Type-Specific Sheets (One per NormalizedResourceType)
        # 3. Type-Specific Sheets (One per NormalizedResourceType)
        unique_types = df['NormalizedResourceType'].unique()
        # Drop standalone ingress/egress rules as independent sheets per user request
        unique_types = [t for t in unique_types if t not in ['AWS::EC2::SecurityGroupIngress', 'AWS::EC2::SecurityGroupEgress']]
        
        sheet_mapping = {}

        for r_type in unique_types:
            type_df = df[df['NormalizedResourceType'] == r_type].copy()
            type_df.dropna(axis=1, how='all', inplace=True)
            
            # Special Handling for Security Groups: EXPLODE rules into separate rows
            if r_type == 'AWS::EC2::SecurityGroup':
                def _split_rules(raw):
                    """Split a multi-line rule blob, filtering out empty
                    strings and pandas' NaN-as-string artifacts. Without this
                    filter, str(NaN).split('\\n') yields ['nan'] which the
                    downstream parser would treat as a real rule, producing
                    2 phantom rule rows per SG (one ingress, one egress)."""
                    if raw is None:
                        return []
                    if isinstance(raw, float) and pd.isna(raw):
                        return []
                    parts = str(raw).split('\n')
                    return [
                        p for p in parts
                        if p.strip() and p.strip().lower() not in ('nan', 'none')
                    ]

                rows = []
                for _, sg_row in type_df.iterrows():
                    # Parse summarized Ingress and Egress columns
                    # Format: "TCP:80 from 0.0.0.0/0 | Description"
                    ingress = _split_rules(sg_row.get('Ingress Rules', '')) if 'Ingress Rules' in sg_row else []
                    egress = _split_rules(sg_row.get('Egress Rules', '')) if 'Egress Rules' in sg_row else []

                    all_rules = []
                    for r_str in ingress:
                         if r_str.strip(): all_rules.append({'Direction': 'Inbound', 'Rule': r_str})
                    for r_str in egress:
                         if r_str.strip(): all_rules.append({'Direction': 'Outbound', 'Rule': r_str})
                    
                    if not all_rules:
                        # Keep the SG even if no rules — still populate rule
                        # columns (empty) so downstream reorder doesn't KeyError
                        # when every SG in scope has no rules.
                        new_row = sg_row.to_dict()
                        new_row['Rule Direction'] = ''
                        new_row['Rule Protocol'] = ''
                        new_row['Rule Port Range'] = ''
                        new_row['Rule Source/Dest'] = ''
                        new_row['Rule Description'] = ''
                        rows.append(new_row)
                    else:
                        def _port_value(s):
                            """Convert single-port strings to ints so Excel
                            writes them as numeric cells (no yellow ! "number
                            stored as text" warning). Port ranges ("80-90")
                            and keywords ("ALL") stay as strings."""
                            s = str(s or '').strip()
                            if s.isdigit():
                                try:
                                    return int(s)
                                except (ValueError, TypeError):
                                    return s
                            return s

                        for rule_data in all_rules:
                            new_row = sg_row.to_dict()
                            new_row['Rule Direction'] = rule_data['Direction']

                            # Parse: "TCP:80 from 0.0.0.0/0 | Desc"
                            r_text = rule_data['Rule']
                            proto_port, other = r_text.split(' from ', 1) if ' from ' in r_text else (r_text, '')
                            proto, port = proto_port.split(':', 1) if ':' in proto_port else (proto_port, 'ALL')

                            source, desc = other.split('|', 1) if '|' in other else (other, '')

                            new_row['Rule Protocol'] = proto.strip()
                            new_row['Rule Port Range'] = _port_value(port)
                            new_row['Rule Source/Dest'] = source.strip()
                            new_row['Rule Description'] = desc.strip()
                            rows.append(new_row)

                type_df = pd.DataFrame(rows)
                # Ensure new columns appear at the start for visibility
                sc_cols = ['Rule Direction', 'Rule Protocol', 'Rule Port Range', 'Rule Source/Dest', 'Rule Description']
                # Guard against empty rows or missing keys (defensive: handles
                # the case where no SGs were processed at all).
                for c in sc_cols:
                    if c not in type_df.columns:
                        type_df[c] = ''
                cols = [c for c in type_df.columns if c not in sc_cols]
                type_df = type_df[sc_cols + cols]
                # Drop original summarized columns
                type_df.drop(columns=['Ingress Rules', 'Egress Rules'], errors='ignore', inplace=True)
            
            # --- Azure Mapping / Enrichment ---
            type_df['Azure Service'] = type_df['NormalizedResourceType'].apply(lambda x: get_azure_mapping({'NormalizedResourceType': x}))
            
            def get_az(row):
                return row.get('AvailabilityZone') or row.get('AvailabilityZoneId') or row.get('Placement', {}).get('AvailabilityZone') or ''
            
            az_series = type_df.apply(get_az, axis=1)
            if not az_series.astype(str).str.strip().eq('').all():
                type_df['Availability Zone'] = az_series
                def get_azure_zone(az):
                    # Return as int (not str) so Excel renders it as a number
                    # and doesn't show the "number stored as text" yellow !
                    # warning triangle. None for unmappable values.
                    az = str(az).strip().lower()
                    if not az or not az[-1].isalpha():
                        return None
                    char = az[-1]
                    if char >= 'a' and char <= 'f':
                        return (ord(char) - ord('a')) % 3 + 1
                    return None
                type_df['Azure Zone'] = pd.to_numeric(
                    type_df['Availability Zone'].apply(get_azure_zone),
                    errors='coerce',
                ).astype('Int64')  # nullable int — keeps blanks blank, not NaN

            def get_storage(row):
                 return row.get('Size') or row.get('AllocatedStorage') or row.get('Size (GB)') or row.get('TotalStorage(GB)') or ''
            
            storage_series = type_df.apply(get_storage, axis=1)
            if not storage_series.astype(str).str.strip().replace('0', '').eq('').all():
                type_df['Storage Size (GB)'] = storage_series
                type_df['Azure Storage Size (GB)'] = type_df['Storage Size (GB)']

            def get_azure_redundancy(row):
                if str(row.get('MultiAZ')).lower() == 'true' or str(row.get('MultiAZEnabled')).lower() == 'true':
                    return 'Zone Redundant (HA)'
                sc = str(row.get('StorageClass', ''))
                if 'ONEZONE' in sc: return 'LRS'
                if 'GLACIER' in sc or 'DEEP_ARCHIVE' in sc: return 'LRS'
                if 'STANDARD' in sc: return 'ZRS' 
                return ''
            
            redundancy_series = type_df.apply(get_azure_redundancy, axis=1)
            if not redundancy_series.astype(str).str.strip().eq('').all():
                type_df['Azure Redundancy'] = redundancy_series

            # 3. Resource-Specific Enrichment
            if r_type == 'AWS::EC2::Instance':
                type_df['Azure SKU'] = (type_df['InstanceType'] if 'InstanceType' in type_df.columns else pd.Series([''] * len(type_df))).apply(lambda x: get_azure_sku_recommendation(x, 'ec2'))
                type_df['Azure OS'] = (type_df['Platform'] if 'Platform' in type_df.columns else pd.Series(['linux'] * len(type_df))).apply(lambda x: 'Windows' if 'windows' in str(x).lower() else 'Linux')

                # Azure OS Disk Size — round to next Azure managed-disk
                # SKU tier (module-level _round_to_azure_disk_tier).
                # e.g. AWS 80 GB → Azure 128 GB (P10); AWS 60 GB → 64 GB (P6).
                if 'OSDiskSize(GB)' in type_df.columns:
                    type_df['Azure OS Disk Size'] = type_df['OSDiskSize(GB)'].apply(
                        lambda x: (f"{_round_to_azure_disk_tier(x)} GB"
                                   if _round_to_azure_disk_tier(x) else '128 GB')
                    )
                else:
                    type_df['Azure OS Disk Size'] = '128 GB'

                # Azure Data Disk Size — PER-DISK Azure tier mapping, comma
                # separated. The native scan populates 'AzureDataDiskSizes'
                # by rounding each AWS data disk's size individually. E.g.,
                # AWS [25 GB, 100 GB, 500 GB] → Azure "32, 128, 512".
                if 'AzureDataDiskSizes' in type_df.columns:
                    type_df['Azure Data Disk Size'] = type_df['AzureDataDiskSizes']
                elif 'DataDiskTotalSize(GB)' in type_df.columns:
                    # Fallback: total rounded if per-disk wasn't captured.
                    type_df['Azure Data Disk Size'] = type_df['DataDiskTotalSize(GB)'].apply(
                        lambda x: _round_to_azure_disk_tier(x) if x else 0
                    )
                else:
                    type_df['Azure Data Disk Size'] = 0
                type_df['Azure Disk Tier'] = 'Premium SSD'

            elif r_type == 'AWS::RDS::DBInstance':
                type_df['Azure SKU'] = (type_df['DBInstanceClass'] if 'DBInstanceClass' in type_df.columns else pd.Series([''] * len(type_df))).apply(lambda x: get_azure_sku_recommendation(x, 'rds'))
                def map_engine(e):
                    e = str(e).lower()
                    if 'aurora-mysql' in e or 'mysql' in e: return 'Azure Database for MySQL Flexible Server'
                    if 'aurora-postgresql' in e or 'postgres' in e: return 'Azure Database for PostgreSQL Flexible Server'
                    if 'mariadb' in e: return 'Azure Database for MariaDB'
                    if 'oracle' in e: return 'Azure SQL Managed Instance (Oracle Migration)'
                    if 'sqlserver' in e or 'sql' in e: return 'Azure SQL Managed Instance'
                    return f"Azure Database ({e})"
                type_df['Azure Engine'] = (type_df['Engine'] if 'Engine' in type_df.columns else pd.Series(['Unknown'] * len(type_df))).apply(map_engine)
                type_df['Azure Service'] = type_df['Azure Engine']

            elif r_type == 'AWS::Redshift::Cluster':
                 type_df['Azure SKU'] = (type_df['NodeType'] if 'NodeType' in type_df.columns else pd.Series([''] * len(type_df))).apply(lambda x: get_azure_sku_recommendation(x, 'ec2'))

            elif r_type == 'AWS::OpenSearchService::Domain':
                 type_df['Azure SKU'] = (type_df['InstanceType'] if 'InstanceType' in type_df.columns else pd.Series([''] * len(type_df))).apply(lambda x: get_azure_sku_recommendation(x, 'opensearch'))

            # Sequence Cols — build with a `seen` set so a column added twice
            # (e.g., 'Azure Zone' appears in both `azure_cols` and the
            # `other_at_end` hint list) shows up only ONCE in the output.
            azure_cols = [c for c in type_df.columns if c.startswith('Azure')]
            other_at_end_hint = ['Availability Zone', 'Azure Zone', 'Azure Redundancy', 'Storage Size (GB)', 'Azure Storage Size (GB)']
            other_at_end = [c for c in other_at_end_hint if c in type_df.columns and c not in azure_cols]

            seen = set()
            final_cols = []
            def _add_col(c):
                if c in type_df.columns and c not in seen:
                    seen.add(c)
                    final_cols.append(c)
            for c in base_cols: _add_col(c)
            for c in type_df.columns:
                if c not in azure_cols and c not in other_at_end:
                    _add_col(c)
            for c in azure_cols: _add_col(c)
            for c in other_at_end: _add_col(c)

            # Drop Junk - Stopped dropping BlockDeviceMappings and NetworkInterfaces to preserve disk info
            drop_prefixes = ['NetworkInterfaceId','PrivateDnsName','Arn','ResponseMetadata','RequestId']
            explicit_drop = [
                'AmiLaunchIndex', 'Monitoring', 'EbsOptimized', 'KeyName', 'Placement',
                'CreationDate', 'LaunchTime', 'AvailabilityZone',
                # NOTE: 'OSDiskSize(GB)' used to be dropped here, but it's the
                # only column carrying the AWS OS-disk size as a raw number
                # (without the "GB" suffix). Keep it so users can sort/filter,
                # alongside the formatted "Azure OS Disk Size" ("80.0 GB").
                # Noise / duplicates exposed by the multi-source merge:
                'ResourceTypeSource',       # internal flag ("NativeScan") — not user-facing
                'NameTag',                  # duplicate of ResourceName after the Name-tag upgrade
                'Storage Size (GB)',        # generic storage indicator — Azure-specific columns cover this
                'Azure Storage Size (GB)',  # redundant; OS Disk + Data Disk give the breakdown
                'AzureDataDiskSizes',       # intermediate field — surfaced as "Azure Data Disk Size"
            ]
            cols_to_drop = []
            for c in final_cols:
                if any(c.startswith(p) for p in drop_prefixes) or c in explicit_drop:
                    cols_to_drop.append(c)
                    continue
                if type_df[c].apply(is_empty_val).all():
                    cols_to_drop.append(c)
            
            final_cols = [c for c in final_cols if c not in cols_to_drop]
            type_df = type_df[final_cols]
            
            # Data Sanitization Loop: Clean strings of '!' and emojis, and normalize numeric types
            # Use iloc (column position) to avoid crashes from duplicate column names
            # Columns that should stay as-is (human-readable strings like '30 GB')
            skip_numeric = {'Azure OS Disk Size', 'Azure Data Disk Size', 'Azure Disk Tier'}
            for i, col in enumerate(type_df.columns):
                series = type_df.iloc[:, i]
                # 1. Clean '!' and emojis from all object/string columns
                if series.dtype == 'object':
                    type_df.iloc[:, i] = series.apply(lambda x: str(x).replace('!', '').replace('⚠️', '').replace('✅', '').replace('❌', '') if pd.notnull(x) else x)
                
                # 2. Convert common numeric-looking columns to numeric types (fixes Excel warnings)
                # Skip Azure disk columns that should remain as readable strings
                if col not in skip_numeric and any(k in col for k in ['Size', 'GB', 'Count', 'RulesCount', 'Nodes', 'Memory']):
                    try:
                        # Re-read series after emoji cleaning
                        fresh_series = type_df.iloc[:, i]
                        type_df.iloc[:, i] = pd.to_numeric(fresh_series.astype(str).str.replace(' GB', '', case=False).replace('Unknown', '0').replace('', '0'), errors='coerce')
                    except TypeError:

                        pass  # Column has strict StringDtype, skip numeric conversion

            # Sheet Naming Sanitization
            friendly = get_friendly_name(r_type).replace('!', '').replace('⚠️', '').replace('✅', '').replace('❌', '')
            sheet_name = friendly.replace('/', '-').replace(':', '').replace('[','').replace(']','')[:31]
            sheet_mapping[r_type] = sheet_name

            try:
                type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_sheet(sheet_name)
            except Exception as e:
                print(f"Failed to write [ {r_type} ]: {e}")

        # --- Summary → per-type sheet hyperlinks ---
        # Every Friendly Name in the Summary becomes a clickable link that
        # jumps to its corresponding sheet. Sheet names in Excel can't start
        # with '#', so internal links use the format `#'Sheet Name'!A1`.
        try:
            ws_summary = writer.sheets['Summary']
            # Build friendly-name → sheet-name map from sheet_mapping (which
            # is keyed by NormalizedResourceType).
            friendly_to_sheet = {}
            for r_type, s_name in sheet_mapping.items():
                friendly = get_friendly_name(r_type).replace('!', '').replace('⚠️', '').replace('✅', '').replace('❌', '')
                friendly_to_sheet[friendly] = s_name

            # Skip header row (row 1); iterate data rows. Friendly Name is in
            # column 1 (the Summary was written with columns
            # ['Friendly Name', 'Azure Service', 'Count']).
            link_font = Font(color='0563C1', underline='single', bold=False)
            link_count = 0
            for row in ws_summary.iter_rows(min_row=2, max_row=ws_summary.max_row, max_col=1):
                cell = row[0]
                if cell.value is None:
                    continue
                friendly = str(cell.value).strip()
                target_sheet = friendly_to_sheet.get(friendly)
                if not target_sheet:
                    continue
                # openpyxl accepts the internal link as a single string with
                # the worksheet name quoted (so spaces / special chars survive).
                cell.hyperlink = f"#'{target_sheet}'!A1"
                cell.font = link_font
                link_count += 1
            print(f"   Summary hyperlinks: {link_count} clickable rows")
        except Exception as e:
            print(f"   Summary hyperlinks skipped: {_short_err(e)}")

        # --- Inter-Sheet Hyperlinking ---
        try:
             import openpyxl
             sg_sheet = sheet_mapping.get('AWS::EC2::SecurityGroup')
             ec2_sheet = sheet_mapping.get('AWS::EC2::Instance')
             if sg_sheet and ec2_sheet and ec2_sheet in writer.sheets:
                 ws_ec2 = writer.sheets[ec2_sheet]
                 # Find Security Group column index (Support multiple names)
                 sg_col_idx = -1
                 for i, c in enumerate(ws_ec2[1]):
                     col_name = str(c.value).strip()
                     if col_name in ['Security Groups', 'SecurityGroupIds', 'SecurityGroups']:
                         sg_col_idx = i + 1
                         break
                 
                 if sg_col_idx != -1:
                     blue_font = openpyxl.styles.Font(color="0000FF", underline="single")
                     for row in range(2, ws_ec2.max_row + 1):
                         cell = ws_ec2.cell(row=row, column=sg_col_idx)
                         if cell.value:
                              cell.hyperlink = f"#'{sg_sheet}'!A1"
                              cell.font = blue_font
        except Exception as e:
             print(f"   ⚠️ Hyperlink creation failed: {_short_err(e)}")
    return filepath
