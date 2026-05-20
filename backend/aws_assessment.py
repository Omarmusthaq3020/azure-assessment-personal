import boto3
import pandas as pd
import os
from datetime import datetime, timedelta
import uuid
from botocore.config import Config
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


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
        # Keep clients cached
        self.tagging_client = self.session.client('resourcegroupstaggingapi')
        # Fail fast configuration for Cloud Control to prevent hanging
        fast_fail_config = Config(
            retries={'max_attempts': 2, 'mode': 'standard'},
            connect_timeout=5,
            read_timeout=10
        )
        self.cloud_control = self.session.client('cloudcontrol', config=fast_fail_config)

    def _prepare_networking_metadata(self):
        """
        Builds a cache of VPC/Subnet relationships and identifies Public vs Private subnets.
        A subnet is considered 'Public' if its associated route table has a route to an Internet Gateway.
        """
        print(f"Pre-fetching networking metadata for {self.region}...", flush=True)
        self._subnet_cache = {}
        try:
            ec2 = self.session.client('ec2')
            
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
            print(f"   ⚠️ Networking metadata fetch failed: {e}")

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
            ec2 = self.session.client('ec2')
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
        
        # 1. Get Master List from Tagging API (Covers ~90% of resources)
        # This is the "Resource Graph" equivalent for AWS
        try:
            paginator = self.tagging_client.get_paginator('get_resources')
            for page in paginator.paginate(ResourcesPerPage=100):
                for resource in page['ResourceTagMappingList']:
                    arn = resource['ResourceARN']
                    parsed = self._parse_arn(arn)
                    
                    # Filter by region
                    if parsed.get('region') and parsed['region'] != self.region:
                        continue

                    known_arns.add(arn)
                    if parsed['id_value']:
                        known_arns.add(parsed['id_value'])
                    
                    tags = self._format_tags(resource.get('Tags', []))
                    name = self._get_tag_value(resource.get('Tags', []), 'Name') or parsed['id_value']
                    
                    group = self._get_resource_group(parsed['service'])
                    cfn_type = self._guess_cfn_type(parsed['service'], parsed['type_id'])
                    if not cfn_type:
                        cfn_type = f"AWS::{parsed['service']}::{parsed['type_id']}"

                    all_resources.append({
                        'ResourceId': arn, 
                        'ResourceName': name,
                        'ResourceType': cfn_type,
                        'NormalizedResourceType': cfn_type,
                        'ResourceTypeSource': 'Tagging',
                        'ResourceGroup': group,
                        'Location': self.region,
                        'Status': 'Unknown',
                        'Details': arn,
                        'Tags': tags,
                        '_arn': arn,         
                        '_service': parsed['service'],
                        '_type_id': parsed['type_id'],
                        '_id_value': parsed['id_value']
                    })
            print(f"✅ Found {len(all_resources)} resources via Tagging API.", flush=True)
        except Exception as e:
            # Common for users without global tagging permissions
            print(f"⚠️ Tagging API Access Denied or failed: {str(e)}. Falling back to direct service scans...", flush=True)
            log_progress("Warning", 10, "Tagging API restricted. Using direct service discovery...")

        # 1.1 Direct Native Scanning for Critical Services (Always run as fallback/safety)
        self._scan_native_active_services(all_resources, known_arns)

        log_progress("Extracting", 20, f"Found {len(all_resources)} resources. Scanning untagged...")
        print(f"Found {len(all_resources)} resources via Tagging API. Scanning core services for untagged resources...")
        
        # 1.5 Explicitly scan Native Essentials (S3, KeyPairs)
        try:
            self._scan_native_essentials(all_resources, known_arns)
        except Exception as e:
            print(f"   ⚠️ Native Essentials fail: {e}")

        # 1.6 Scan core services for untagged resources
        # Tagging API misses resources that have never had tags.
        try:
            self._discover_untagged_resources(all_resources, known_arns)
        except Exception as e:
            print(f"⚠️ Discovery failure: {e}")

        # 1.7 Map SG Rules to named instances (must run after core discovery)
        self._map_security_group_attachments(all_resources)
        
        log_progress("Enriching", 40, f"Discovered {len(all_resources)} total resources. Starting enrichment...")
        print(f"Total resources after discovery: {len(all_resources)}. Starting enrichment...")
        
        # 2. Enrich high-value resources
        # We process the list and update 'Status' and 'Details' for specific types
        try:
            self._enrich_resources(all_resources)
        except Exception as e:
            print(f"⚠️ Enrichment partially failed: {e}")
        
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
        
        # Helper maps for Upsert/Enrichment
        # We index by ARN and ID to match existing Tagging results
        res_map_arn = {r.get('_arn'): r for r in all_resources if r.get('_arn')}
        res_map_id = {r.get('_id_value'): r for r in all_resources if r.get('_id_value')}

        def upsert_resource(arn, name, type_norm, group, details_dict, status='Active', tags_list=[]):
            nonlocal found_count
            
            # 2. Check if exists
            # Try ARN match first
            existing = res_map_arn.get(arn)
            if not existing and name:
                 # Try ID match
                 existing = res_map_id.get(name)

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
            s3 = self.session.client('s3')
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
                        cw = self.session.client('cloudwatch')
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
            iam = self.session.client('iam')
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
            ec2 = self.session.client('ec2')
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
            acm = self.session.client('acm')
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
            rds = self.session.client('rds')
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
            elbv2 = self.session.client('elbv2')
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
            lam = self.session.client('lambda')
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
            ec2 = self.session.client('ec2')
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
                        name=v['VolumeId'], # Name tag usually handled by Upsert if present in Tags
                        type_norm='AWS::EC2::Volume',
                        group='Storage',
                        details_dict=props,
                        status=v['State'],
                        tags_list=v.get('Tags', [])
                    )
        except Exception: pass

        # H2. EC2 Instances (Native for OS/Storage details)
        try:
            ec2 = self.session.client('ec2')
            
            # Build volume size cache: describe_instances doesn't include VolumeSize,
            # only VolumeId. We must look up sizes from describe_volumes.
            volume_size_cache = {}
            try:
                vol_paginator = ec2.get_paginator('describe_volumes')
                for vol_page in vol_paginator.paginate():
                    for vol in vol_page['Volumes']:
                        volume_size_cache[vol['VolumeId']] = vol['Size']
            except Exception as e:
                print(f"   Volume cache build failed: {e}")
            
            paginator = ec2.get_paginator('describe_instances')
            for page in paginator.paginate():
                for r in page['Reservations']:
                    for i in r['Instances']:
                        # Calculate total storage and identify OS disk using volume cache
                        total_storage = 0
                        root_device = i.get('RootDeviceName')
                        os_disk_size = 0
                        
                        for bdm in i.get('BlockDeviceMappings', []):
                            if 'Ebs' in bdm:
                                vol_id = bdm['Ebs'].get('VolumeId', '')
                                # Look up actual size from volume cache
                                vol_size = volume_size_cache.get(vol_id, 0)
                                total_storage += vol_size
                                if bdm.get('DeviceName') == root_device:
                                    os_disk_size = vol_size
                        
                        # If we couldn't determine OS disk size, use a sensible default
                        if os_disk_size == 0:
                            os_disk_size = 30  # Common AWS default for Linux AMIs
                        
                        props = {
                            'InstanceType': i['InstanceType'],
                            'Platform': i.get('Platform', 'linux'), # 'windows' or empty (linux)
                            'State': i['State']['Name'],
                            'PublicIp': i.get('PublicIpAddress'),
                            'PrivateIp': i.get('PrivateIpAddress'),
                            'OSDiskSize(GB)': os_disk_size,
                            'TotalStorage(GB)': total_storage,
                            'VpcId': i.get('VpcId'),
                            'SubnetId': i.get('SubnetId')
                        }
                        upsert_resource(
                            arn=f"arn:aws:ec2:{self.region}:{self.session.client('sts').get_caller_identity()['Account']}:instance/{i['InstanceId']}",
                            name=i.get('KeyName', i['InstanceId']), # Fallback
                            type_norm='AWS::EC2::Instance',
                            group='Compute',
                            details_dict=props,
                            status=i['State']['Name'],
                            tags_list=i.get('Tags', [])
                        )
        except Exception as e: print(f"Native EC2 Instances failed: {e}")


        try:
            asg_client = self.session.client('autoscaling')
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
        except Exception as e: print(f"Native ASG failed: {e}")

        # J. DynamoDB
        try:
            ddb = self.session.client('dynamodb')
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
        except Exception as e: print(f"Native DynamoDB failed: {e}")

        # K. ElastiCache
        try:
            elasticache = self.session.client('elasticache')
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
        except Exception as e: print(f"Native ElastiCache failed: {e}")

        # L. Redshift Clusters
        try:
            redshift = self.session.client('redshift')
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
                        arn=f"arn:aws:redshift:{self.region}:{self.session.client('sts').get_caller_identity()['Account']}:cluster:{c['ClusterIdentifier']}", 
                        name=c['ClusterIdentifier'],
                        type_norm='AWS::Redshift::Cluster',
                        group='Database',
                        details_dict=props,
                        status=c['ClusterStatus']
                    )
        except Exception as e: print(f"Native Redshift failed: {e}")

        # M. OpenSearch Domains
        try:
            opensearch = self.session.client('opensearch')
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
        except Exception as e: print(f"Native OpenSearch failed: {e}")

        # N. SSM Documents (Custom ones only)
        try:
            ssm = self.session.client('ssm')
            # Fetch Account ID once for ARNs
            account_id = self.session.client('sts').get_caller_identity()['Account']
            
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
        except Exception as e: print(f"Native SSM failed: {e}")

        # O. VPCs and Subnets (Ensure CIDRs are ALWAYS captured)
        try:
            ec2 = self.session.client('ec2')
            account_id = self.session.client('sts').get_caller_identity()['Account']
            
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
                    name=vpc_id,
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
                    name=sub_id,
                    type_norm='AWS::EC2::Subnet',
                    group='Networking',
                    details_dict=props,
                    status=sub.get('State', 'Active'),
                    tags_list=sub.get('Tags', [])
                )

            # 3. Security Groups
            sgs = ec2.describe_security_groups()['SecurityGroups']
            for sg in sgs:
                sg_id = sg['GroupId']
                sg_name = sg.get('GroupName', sg_id)
                
                # A. The Group itself
                props = {
                    'GroupName': sg_name,
                    'Description': sg.get('Description'),
                    'VpcId': sg.get('VpcId'),
                    'IngressRulesCount': len(sg.get('IpPermissions', [])),
                    'EgressRulesCount': len(sg.get('IpPermissionsEgress', []))
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
                    name=eni_id,
                    type_norm='AWS::EC2::NetworkInterface',
                    group='Networking',
                    details_dict=props,
                    status=eni.get('Status', 'Active'),
                    tags_list=eni.get('Tags', [])
                )

        except Exception as e: print(f"Native Networking (VPC/Subnet/SG/ENI) failed: {e}")

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
                s3 = self.session.client('s3')
                for b in s3.list_buckets().get('Buckets', []):
                    add_res(f"arn:aws:s3:::{b['Name']}", b['Name'], 'AWS::S3::Bucket', 's3', 'bucket')
                
                iam = self.session.client('iam')
                for u in iam.list_users().get('Users', []):
                    add_res(u['Arn'], u['UserName'], 'AWS::IAM::User', 'iam', 'user')
                for r in iam.list_roles().get('Roles', []):
                    add_res(r['Arn'], r['RoleName'], 'AWS::IAM::Role', 'iam', 'role')
                
                cf = self.session.client('cloudfront')
                distributions = cf.list_distributions().get('DistributionList', {}).get('Items', [])
                for d in distributions:
                    add_res(d['ARN'], d['Id'], 'AWS::CloudFront::Distribution', 'cloudfront', 'distribution')

                r53 = self.session.client('route53')
                for z in r53.list_hosted_zones().get('HostedZones', []):
                    add_res(z['Id'], z['Name'], 'AWS::Route53::HostedZone', 'route53', 'hostedzone')

                setattr(self.__class__, '_global_scan_done', True)
            except: pass

        # 1. Regional Scans
        for svc, method, key, subkey, id_field, cfn_type, s_short, t_short in core_scans:
            try:
                client = self.session.client(svc)
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
        cfn_client = self.session.client('cloudformation')
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
             self.cloud_control = self.session.client('cloudcontrol', config=fast_fail_config)

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
                                r.update(flat_props)
                                
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
                             ec2_client = self.session.client('ec2')
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
        'ResourceTypeSource', 'ResourceGroup', 'Location', 'Status', 'Ingress Rules', 'Egress Rules', 'Details', 'Tags'
    ]
    for col in base_cols:
        if col not in df.columns:
            df[col] = ''
    
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
        # Add Friendly Name
        summary['Friendly Name'] = summary['NormalizedResourceType'].apply(get_friendly_name)
        
        # Add Azure Service Mapping (Representative)
        def get_svc_rep(rtype):
            rows = df[df['NormalizedResourceType'] == rtype]
            if not rows.empty:
                return get_azure_mapping(rows.iloc[0])
            return ''
            
        summary['Azure Service'] = summary['NormalizedResourceType'].apply(get_svc_rep)

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
                rows = []
                for _, sg_row in type_df.iterrows():
                    # Parse summarized Ingress and Egress columns
                    # Format: "TCP:80 from 0.0.0.0/0 | Description"
                    ingress = str(sg_row.get('Ingress Rules', '')).split('\n') if 'Ingress Rules' in sg_row else []
                    egress = str(sg_row.get('Egress Rules', '')).split('\n') if 'Egress Rules' in sg_row else []
                    
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
                        for rule_data in all_rules:
                            new_row = sg_row.to_dict()
                            new_row['Rule Direction'] = rule_data['Direction']

                            # Parse: "TCP:80 from 0.0.0.0/0 | Desc"
                            r_text = rule_data['Rule']
                            proto_port, other = r_text.split(' from ', 1) if ' from ' in r_text else (r_text, '')
                            proto, port = proto_port.split(':', 1) if ':' in proto_port else (proto_port, 'ALL')

                            source, desc = other.split('|', 1) if '|' in other else (other, '')

                            new_row['Rule Protocol'] = proto.strip()
                            new_row['Rule Port Range'] = port.strip()
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
                     az = str(az).strip().lower()
                     if not az or not az[-1].isalpha(): return ''
                     char = az[-1]
                     if char >= 'a' and char <= 'f':
                         return str((ord(char) - ord('a')) % 3 + 1)
                     return ''
                type_df['Azure Zone'] = type_df['Availability Zone'].apply(get_azure_zone)

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
                
                # Make OS Disk Size dynamic (Uses OSDiskSize(GB) captured earlier)
                if 'OSDiskSize(GB)' in type_df.columns:
                    type_df['Azure OS Disk Size'] = type_df['OSDiskSize(GB)'].apply(lambda x: f"{x} GB" if pd.notnull(x) else '127 GB')
                else:
                    type_df['Azure OS Disk Size'] = '127 GB' 

                type_df['Azure Data Disk Size'] = type_df['TotalStorage(GB)'] if 'TotalStorage(GB)' in type_df.columns else 0
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

            # Sequence Cols
            azure_cols = [c for c in type_df.columns if c.startswith('Azure')]
            other_at_end = [c for c in ['Availability Zone', 'Azure Zone', 'Azure Redundancy', 'Storage Size (GB)', 'Azure Storage Size (GB)'] if c in type_df.columns]
            final_cols = [c for c in base_cols if c in type_df.columns]
            final_cols += [c for c in type_df.columns if c not in final_cols and c not in azure_cols and c not in other_at_end]
            final_cols += azure_cols + other_at_end

            # Drop Junk - Stopped dropping BlockDeviceMappings and NetworkInterfaces to preserve disk info
            drop_prefixes = ['NetworkInterfaceId','PrivateDnsName','Arn','ResponseMetadata','RequestId']
            explicit_drop = ['AmiLaunchIndex', 'Monitoring', 'EbsOptimized', 'KeyName', 'Placement', 'CreationDate', 'LaunchTime', 'AvailabilityZone', 'OSDiskSize(GB)']
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
             print(f"   ⚠️ Hyperlink creation failed: {e}")
    return filepath
