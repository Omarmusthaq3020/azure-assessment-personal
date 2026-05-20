import json
import os
import math

class SkuMapper:
    """
    Provides intelligent attributes-based mapping from AWS SKUs to Azure SKUs.
    Contains built-in database of common Instance Spec definitions.
    """
    
    def __init__(self):
        self._aws_specs = None   # lazy-loaded on first AWS mapping call
        self.azure_specs = self._load_azure_specs()

    @property
    def aws_specs(self):
        """Lazy-load AWS specs only when actually needed (i.e. an AWS SKU mapping)."""
        if self._aws_specs is None:
            self._aws_specs = self._load_aws_specs()
        return self._aws_specs

    def map_sku(self, aws_sku, resource_type='ec2'):
        """
        Returns the best matching Azure SKU string.
        """
        # Fix: Ensure aws_sku is a string to avoid 'float' object has no attribute 'lower'
        if not aws_sku or not isinstance(aws_sku, str):
            # If we have CPU/RAM info, we could try to map, but for now fallback to manual
            return "Assess Manually"
            
        aws_sku = aws_sku.lower()
        
        # 0. Direct/Custom Mapping Overrides
        # Removed B-series per user request - prefer D2s_v5 as base
        overrides = {
            't2.micro': 'Standard_D2s_v5', 
            't3.micro': 'Standard_D2s_v5',
        }
        if aws_sku in overrides: return overrides[aws_sku]

        # 1. Lookup AWS Specs
        # If specific SKU not found, try to infer from family+size string?
        # For now, rely on database.
        spec = self.aws_specs.get(aws_sku)
        
        if not spec:
            # Heuristic Fallback if not in DB
            return f"Standard_{aws_sku.replace('.','_')}_(Est)"

        # 2. Find Candidate in Azure
        # Detect if source is likely AMD (a-suffix like m6a, t3a)
        is_amd = ('.a' in aws_sku or 'a.' in aws_sku or 'epyc' in str(spec.get('cpu_type','')).lower())
        
        candidate = self._find_best_match(spec, is_amd=is_amd)
        
        if candidate:
            return candidate['name']
        
        return f"Approx_{aws_sku.replace('.','_')}"

    def map_gcp_to_azure(self, gcp_machine_type, resource_type='VM'):
        """
        Maps GCP machine types to Azure VM SKUs.
        
        Args:
            gcp_machine_type: GCP machine type (e.g., 'n1-standard-2', 'e2-medium')
            resource_type: Type of resource (VM, DB, etc.)
            
        Returns:
            Azure VM SKU string
        """
        if not gcp_machine_type:
            return "Standard_D2s_v3"
            
        gcp_machine_type = gcp_machine_type.lower()
        
        # Parse GCP machine type to extract specs
        spec = self._parse_gcp_machine_type(gcp_machine_type)
        
        if not spec:
            return f"Standard_{gcp_machine_type.replace('-','_')}"
        
        # Find best Azure match
        candidate = self._find_best_match(spec)
        
        if candidate:
            return candidate['name']
            
        return "Standard_D2s_v3"  # Default fallback
    
    def _parse_gcp_machine_type(self, machine_type):
        """
        Parse GCP machine type and return spec dict.
        
        GCP Machine Type Formats:
        - n1-standard-2 (2 vCPU, 7.5 GB)
        - n1-highmem-4 (4 vCPU, 26 GB)
        - n1-highcpu-8 (8 vCPU, 7.2 GB)
        - e2-medium (2 vCPU, 4 GB)
        - e2-standard-4 (4 vCPU, 16 GB)
        - n2-standard-8 (8 vCPU, 32 GB)
        """
        # Common GCP machine type mappings
        gcp_specs = {
            # E2 Series (Shared-core and standard)
            'e2-micro': {'vcpu': 2, 'ram': 1, 'family': 'burstable'},
            'e2-small': {'vcpu': 2, 'ram': 2, 'family': 'burstable'},
            'e2-medium': {'vcpu': 2, 'ram': 4, 'family': 'burstable'},
            'e2-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'general'},
            'e2-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'general'},
            'e2-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'general'},
            'e2-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'general'},
            'e2-highmem-2': {'vcpu': 2, 'ram': 16, 'family': 'memory'},
            'e2-highmem-4': {'vcpu': 4, 'ram': 32, 'family': 'memory'},
            'e2-highmem-8': {'vcpu': 8, 'ram': 64, 'family': 'memory'},
            'e2-highcpu-2': {'vcpu': 2, 'ram': 4, 'family': 'compute'},
            'e2-highcpu-4': {'vcpu': 4, 'ram': 8, 'family': 'compute'},
            'e2-highcpu-8': {'vcpu': 8, 'ram': 16, 'family': 'compute'},
            
            # N1 Series (First generation)
            'n1-standard-1': {'vcpu': 1, 'ram': 3.75, 'family': 'general'},
            'n1-standard-2': {'vcpu': 2, 'ram': 7.5, 'family': 'general'},
            'n1-standard-4': {'vcpu': 4, 'ram': 15, 'family': 'general'},
            'n1-standard-8': {'vcpu': 8, 'ram': 30, 'family': 'general'},
            'n1-standard-16': {'vcpu': 16, 'ram': 60, 'family': 'general'},
            'n1-highmem-2': {'vcpu': 2, 'ram': 13, 'family': 'memory'},
            'n1-highmem-4': {'vcpu': 4, 'ram': 26, 'family': 'memory'},
            'n1-highmem-8': {'vcpu': 8, 'ram': 52, 'family': 'memory'},
            'n1-highcpu-2': {'vcpu': 2, 'ram': 1.8, 'family': 'compute'},
            'n1-highcpu-4': {'vcpu': 4, 'ram': 3.6, 'family': 'compute'},
            'n1-highcpu-8': {'vcpu': 8, 'ram': 7.2, 'family': 'compute'},
            
            # N2 Series (Second generation)
            'n2-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'general'},
            'n2-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'general'},
            'n2-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'general'},
            'n2-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'general'},
            'n2-highmem-2': {'vcpu': 2, 'ram': 16, 'family': 'memory'},
            'n2-highmem-4': {'vcpu': 4, 'ram': 32, 'family': 'memory'},
            'n2-highmem-8': {'vcpu': 8, 'ram': 64, 'family': 'memory'},
            'n2-highcpu-2': {'vcpu': 2, 'ram': 4, 'family': 'compute'},
            'n2-highcpu-4': {'vcpu': 4, 'ram': 8, 'family': 'compute'},
            'n2-highcpu-8': {'vcpu': 8, 'ram': 16, 'family': 'compute'},
            
            # N2D Series (AMD EPYC)
            'n2d-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'general'},
            'n2d-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'general'},
            'n2d-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'general'},
            'n2d-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'general'},
            'n2d-standard-32': {'vcpu': 32, 'ram': 128, 'family': 'general'},
            'n2d-standard-48': {'vcpu': 48, 'ram': 192, 'family': 'general'},
            'n2d-standard-64': {'vcpu': 64, 'ram': 256, 'family': 'general'},
            'n2d-standard-96': {'vcpu': 96, 'ram': 384, 'family': 'general'},
            'n2d-highmem-2': {'vcpu': 2, 'ram': 16, 'family': 'memory'},
            'n2d-highmem-4': {'vcpu': 4, 'ram': 32, 'family': 'memory'},
            'n2d-highmem-8': {'vcpu': 8, 'ram': 64, 'family': 'memory'},
            'n2d-highmem-16': {'vcpu': 16, 'ram': 128, 'family': 'memory'},
            'n2d-highmem-32': {'vcpu': 32, 'ram': 256, 'family': 'memory'},
            'n2d-highmem-48': {'vcpu': 48, 'ram': 384, 'family': 'memory'},
            'n2d-highmem-64': {'vcpu': 64, 'ram': 512, 'family': 'memory'},
            'n2d-highmem-96': {'vcpu': 96, 'ram': 768, 'family': 'memory'},
            'n2d-highcpu-2': {'vcpu': 2, 'ram': 4, 'family': 'compute'},
            'n2d-highcpu-4': {'vcpu': 4, 'ram': 8, 'family': 'compute'},
            'n2d-highcpu-8': {'vcpu': 8, 'ram': 16, 'family': 'compute'},
            'n2d-highcpu-16': {'vcpu': 16, 'ram': 32, 'family': 'compute'},
            'n2d-highcpu-32': {'vcpu': 32, 'ram': 64, 'family': 'compute'},
            'n2d-highcpu-48': {'vcpu': 48, 'ram': 96, 'family': 'compute'},
            'n2d-highcpu-64': {'vcpu': 64, 'ram': 128, 'family': 'compute'},
            'n2d-highcpu-96': {'vcpu': 96, 'ram': 192, 'family': 'compute'},

            # C2D Series (AMD Compute Optimized)
            'c2d-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'compute'},
            'c2d-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'compute'},
            'c2d-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'compute'},
            'c2d-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'compute'},
            'c2d-standard-32': {'vcpu': 32, 'ram': 128, 'family': 'compute'},
            'c2d-highcpu-2': {'vcpu': 2, 'ram': 4, 'family': 'compute'},
            'c2d-highcpu-4': {'vcpu': 4, 'ram': 8, 'family': 'compute'},
            'c2d-highcpu-8': {'vcpu': 8, 'ram': 16, 'family': 'compute'},
            'c2d-highcpu-16': {'vcpu': 16, 'ram': 32, 'family': 'compute'},
            'c2d-highcpu-32': {'vcpu': 32, 'ram': 64, 'family': 'compute'},
            'c2d-highmem-2': {'vcpu': 2, 'ram': 16, 'family': 'memory'},
            'c2d-highmem-4': {'vcpu': 4, 'ram': 32, 'family': 'memory'},
            'c2d-highmem-8': {'vcpu': 8, 'ram': 64, 'family': 'memory'},
            'c2d-highmem-16': {'vcpu': 16, 'ram': 128, 'family': 'memory'},
            'c2d-highmem-32': {'vcpu': 32, 'ram': 256, 'family': 'memory'},

            # T2A Series (Ampere ARM)
            't2a-standard-1': {'vcpu': 1, 'ram': 4, 'family': 'general'},
            't2a-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'general'},
            't2a-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'general'},
            't2a-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'general'},
            't2a-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'general'},
            't2a-standard-32': {'vcpu': 32, 'ram': 128, 'family': 'general'},
            't2a-standard-48': {'vcpu': 48, 'ram': 192, 'family': 'general'},

            # N2 larger variants
            'n2-standard-2': {'vcpu': 2, 'ram': 8, 'family': 'general'},
            'n2-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'general'},
            'n2-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'general'},
            'n2-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'general'},
            'n2-standard-32': {'vcpu': 32, 'ram': 128, 'family': 'general'},
            'n2-standard-48': {'vcpu': 48, 'ram': 192, 'family': 'general'},
            'n2-standard-64': {'vcpu': 64, 'ram': 256, 'family': 'general'},
            'n2-standard-80': {'vcpu': 80, 'ram': 320, 'family': 'general'},
            'n2-highmem-2': {'vcpu': 2, 'ram': 16, 'family': 'memory'},
            'n2-highmem-4': {'vcpu': 4, 'ram': 32, 'family': 'memory'},
            'n2-highmem-8': {'vcpu': 8, 'ram': 64, 'family': 'memory'},
            'n2-highmem-16': {'vcpu': 16, 'ram': 128, 'family': 'memory'},
            'n2-highmem-32': {'vcpu': 32, 'ram': 256, 'family': 'memory'},
            'n2-highmem-48': {'vcpu': 48, 'ram': 384, 'family': 'memory'},
            'n2-highmem-64': {'vcpu': 64, 'ram': 512, 'family': 'memory'},
            'n2-highcpu-2': {'vcpu': 2, 'ram': 4, 'family': 'compute'},
            'n2-highcpu-4': {'vcpu': 4, 'ram': 8, 'family': 'compute'},
            'n2-highcpu-8': {'vcpu': 8, 'ram': 16, 'family': 'compute'},
            'n2-highcpu-16': {'vcpu': 16, 'ram': 32, 'family': 'compute'},
            'n2-highcpu-32': {'vcpu': 32, 'ram': 64, 'family': 'compute'},
            'n2-highcpu-48': {'vcpu': 48, 'ram': 96, 'family': 'compute'},
            'n2-highcpu-64': {'vcpu': 64, 'ram': 128, 'family': 'compute'},
            'n2-highcpu-80': {'vcpu': 80, 'ram': 160, 'family': 'compute'},

            # C2 Series (Compute optimized)
            'c2-standard-4': {'vcpu': 4, 'ram': 16, 'family': 'compute'},
            'c2-standard-8': {'vcpu': 8, 'ram': 32, 'family': 'compute'},
            'c2-standard-16': {'vcpu': 16, 'ram': 64, 'family': 'compute'},
            'c2-standard-30': {'vcpu': 30, 'ram': 120, 'family': 'compute'},
            'c2-standard-60': {'vcpu': 60, 'ram': 240, 'family': 'compute'},

            # M1/M2 Series (Memory optimized)
            'm1-megamem-96': {'vcpu': 96, 'ram': 1433.6, 'family': 'memory'},
            'm1-ultramem-40': {'vcpu': 40, 'ram': 961, 'family': 'memory'},
            'm1-ultramem-80': {'vcpu': 80, 'ram': 1922, 'family': 'memory'},
            'm1-ultramem-160': {'vcpu': 160, 'ram': 3844, 'family': 'memory'},
            'm2-ultramem-208': {'vcpu': 208, 'ram': 5888, 'family': 'memory'},
            'm2-megamem-416': {'vcpu': 416, 'ram': 5888, 'family': 'memory'},
        }
        
        return gcp_specs.get(machine_type)

    def _find_best_match(self, source_spec, is_amd=False):
        """
        Finds the smallest Azure SKU that meets requirements.
        """
        s_cpu = source_spec['vcpu']
        s_ram = source_spec['ram']
        s_fam = source_spec['family'] # 'burstable', 'general', 'compute', 'memory'
        
        # Define Family Preferences
        family_priority = []
        if s_fam == 'burstable': 
            family_priority = ['general'] # No B-series fallback per user request
        elif s_fam == 'general': 
            family_priority = ['general'] 
        elif s_fam == 'compute': 
            family_priority = ['compute', 'general']
        elif s_fam == 'memory': 
            family_priority = ['memory', 'general']
        else: 
            family_priority = ['general']
        
        # Filter
        candidates = []
        for az in self.azure_specs:
            if az['vcpu'] >= s_cpu and az['ram'] >= s_ram:
                candidates.append(az)
                
        if not candidates:
            return None
            
        # Score candidates
        scored = []
        for cand in candidates:
            score = (cand['vcpu'] - s_cpu) * 10 + (cand['ram'] - s_ram) * 2
            
            # Family Penalty
            if cand['family'] == s_fam:
                score += 0
            elif cand['family'] in family_priority:
                score += 50 # Acceptable substitute
            else:
                score += 1000 # Bad substitute
            
            # Architecture Bonus/Penalty
            if is_amd:
                if 'as_v5' in cand['name'].lower() or 'eas_v5' in cand['name'].lower():
                    score -= 20 # Prefer AMD for AMD
            else:
                if 'as_v5' in cand['name'].lower() or 'eas_v5' in cand['name'].lower():
                    score += 10 # Prefer Intel for Intel (standard v5)
            
            scored.append((score, cand))
            
        # Sort by lowest score
        scored.sort(key=lambda x: x[0])
        
        return scored[0][1]

    def _load_aws_specs(self):
        """
        Loads AWS specs. Priority:
        1. Cached JSON file (sku_cache_aws.json)
        2. Fetch from GitHub (vantage-sh/ec2instances.info)
        3. Hardcoded Seed (Fallback)
        """
        cache_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sku_cache_aws.json')
        
        # 1. Try Cache
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r') as f:
                    data = json.load(f)
                    if data: return data
            except: pass
            
        # 2. Try Web Fetch
        print("Fetching latest AWS SKU data from ec2instances.info...")
        try:
            import requests
            url = "https://raw.githubusercontent.com/vantage-sh/ec2instances.info/master/instances.json"
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200:
                raw_list = resp.json()
                db = {}
                for i in raw_list:
                    name = i.get('instance_type')
                    if not name: continue
                    
                    vcpu = i.get('vCPU', 1)
                    ram = i.get('memory', 0)
                    fam_raw = i.get('family', '').lower()
                    
                    fam = 'general'
                    if 'burstable' in fam_raw or 'micro' in fam_raw: fam = 'burstable'
                    elif 'compute' in fam_raw: fam = 'compute'
                    elif 'memory' in fam_raw: fam = 'memory'
                    elif 'storage' in fam_raw: fam = 'storage'
                    elif 'gpu' in fam_raw or 'accelerated' in fam_raw: fam = 'gpu'
                    
                    db[name] = {'vcpu': vcpu, 'ram': ram, 'family': fam}
                    
                    # Add db. prefix for RDS
                    db[f"db.{name}"] = db[name]

                # Save Cache
                try:
                    with open(cache_file, 'w') as f:
                        json.dump(db, f)
                except: pass
                
                print(f"Successfully loaded {len(db)} AWS SKUs from web.")
                return db
        except Exception as e:
            print(f"Failed to fetch AWS SKUs: {e}. Using Hardcoded Seed.")

        # 3. Hardcoded Seed
        db = {}
        
        # T2/T3/T3a (Burstable) - Format: name, cpu, ram
        # T2 is 1 vCPU for small sizes, T3 is 2 vCPU base
        burst = [
            # T2
            ('t2.nano', 1, 0.5), ('t2.micro', 1, 1), ('t2.small', 1, 2), ('t2.medium', 2, 4), ('t2.large', 2, 8), ('t2.xlarge', 4, 16),
            # T3
            ('t3.nano', 2, 0.5), ('t3.micro', 2, 1), ('t3.small', 2, 2), ('t3.medium', 2, 4), ('t3.large', 2, 8), ('t3.xlarge', 4, 16), ('t3.2xlarge', 8, 32),
            # T3a (AMD) - Same specs as T3
            ('t3a.nano', 2, 0.5), ('t3a.micro', 2, 1), ('t3a.small', 2, 2), ('t3a.medium', 2, 4), ('t3a.large', 2, 8), ('t3a.xlarge', 4, 16), ('t3a.2xlarge', 8, 32),
            # T4g (ARM) - Same specs as T3 approx (2 vCPU baseline for most)
            ('t4g.nano', 2, 0.5), ('t4g.micro', 2, 1), ('t4g.small', 2, 2), ('t4g.medium', 2, 4), ('t4g.large', 2, 8), ('t4g.xlarge', 4, 16), ('t4g.2xlarge', 8, 32),
        ]
        for n,c,r in burst: db[n] = {'vcpu':c, 'ram':r, 'family':'burstable'}

        # M5/M5a/M6i/M6a (General)
        # Ratio 1:4
        general = [
            ('m4.large', 2, 8), ('m4.xlarge', 4, 16), ('m4.2xlarge', 8, 32),
            ('m5.large', 2, 8), ('m5.xlarge', 4, 16), ('m5.2xlarge', 8, 32), ('m5.4xlarge', 16, 64),
            ('m5a.large', 2, 8), ('m5a.xlarge', 4, 16), ('m5a.2xlarge', 8, 32), ('m5a.4xlarge', 16, 64),
            ('m6i.large', 2, 8), ('m6i.xlarge', 4, 16), ('m6i.2xlarge', 8, 32),
            ('m6a.large', 2, 8), ('m6a.xlarge', 4, 16), ('m6a.2xlarge', 8, 32), ('m6a.4xlarge', 16, 64), ('m6a.8xlarge', 32, 128),
        ]
        for n,c,r in general: db[n] = {'vcpu':c, 'ram':r, 'family':'general'}

        # C5/C5a/C6i/C6a (Compute) - Ratio 1:2
        compute = [
            ('c5.large', 2, 4), ('c5.xlarge', 4, 8), ('c5.2xlarge', 8, 16), ('c5.4xlarge', 16, 32),
            ('c5a.large', 2, 4), ('c5a.xlarge', 4, 8), ('c5a.2xlarge', 8, 16),
            ('c6i.large', 2, 4), ('c6i.xlarge', 4, 8), ('c6i.2xlarge', 8, 16),
            ('c6a.large', 2, 4), ('c6a.xlarge', 4, 8), ('c6a.2xlarge', 8, 16), ('c6a.4xlarge', 16, 32),
        ]
        for n,c,r in compute: db[n] = {'vcpu':c, 'ram':r, 'family':'compute'}

        # R5/R5a/R6/R6a (Memory) - Ratio 1:8
        memory = [
            ('r5.large', 2, 16), ('r5.xlarge', 4, 32), ('r5.2xlarge', 8, 64),
            ('r5a.large', 2, 16), ('r5a.xlarge', 4, 32), ('r5a.2xlarge', 8, 64), ('r5a.4xlarge', 16, 128), ('r5a.8xlarge', 32, 256),
            ('r6i.large', 2, 16), ('r6i.xlarge', 4, 32),
            ('r6a.large', 2, 16), ('r6a.xlarge', 4, 32), ('r6a.2xlarge', 8, 64), ('r6a.4xlarge', 16, 128), ('r6a.8xlarge', 32, 256),
            ('r7g.large', 2, 16), ('r7g.xlarge', 4, 32), ('r7g.2xlarge', 8, 64), ('r7g.4xlarge', 16, 128),
        ]
        for n,c,r in memory: db[n] = {'vcpu':c, 'ram':r, 'family':'memory'}
        
        # Add DB prefixes
        db_keys = list(db.keys())
        for k in db_keys:
            db[f"db.{k}"] = db[k] # Alias for RDS

        return db

    def _load_azure_specs(self):
        """
        Returns list of Azure specs. Removed B-series per user request.
        """
        data = []
        
        # D-Series v5 (General Intel) - 1:4 Ratio
        data.extend([
            {'name': 'Standard_D2s_v5', 'vcpu': 2, 'ram': 8, 'family': 'general'},
            {'name': 'Standard_D4s_v5', 'vcpu': 4, 'ram': 16, 'family': 'general'},
            {'name': 'Standard_D8s_v5', 'vcpu': 8, 'ram': 32, 'family': 'general'},
            {'name': 'Standard_D16s_v5', 'vcpu': 16, 'ram': 64, 'family': 'general'},
            {'name': 'Standard_D32s_v5', 'vcpu': 32, 'ram': 128, 'family': 'general'},
            {'name': 'Standard_D48s_v5', 'vcpu': 48, 'ram': 192, 'family': 'general'},
            {'name': 'Standard_D64s_v5', 'vcpu': 64, 'ram': 256, 'family': 'general'},
            {'name': 'Standard_D96s_v5', 'vcpu': 96, 'ram': 384, 'family': 'general'},
        ])

        # Das-Series v5 (General AMD EPYC) - 1:4 Ratio
        data.extend([
            {'name': 'Standard_D2as_v5', 'vcpu': 2, 'ram': 8, 'family': 'general'},
            {'name': 'Standard_D4as_v5', 'vcpu': 4, 'ram': 16, 'family': 'general'},
            {'name': 'Standard_D8as_v5', 'vcpu': 8, 'ram': 32, 'family': 'general'},
            {'name': 'Standard_D16as_v5', 'vcpu': 16, 'ram': 64, 'family': 'general'},
            {'name': 'Standard_D32as_v5', 'vcpu': 32, 'ram': 128, 'family': 'general'},
            {'name': 'Standard_D48as_v5', 'vcpu': 48, 'ram': 192, 'family': 'general'},
            {'name': 'Standard_D64as_v5', 'vcpu': 64, 'ram': 256, 'family': 'general'},
            {'name': 'Standard_D96as_v5', 'vcpu': 96, 'ram': 384, 'family': 'general'},
        ])

        # F-Series v2 (Compute Intel) - 1:2 Ratio
        data.extend([
            {'name': 'Standard_F2s_v2', 'vcpu': 2, 'ram': 4, 'family': 'compute'},
            {'name': 'Standard_F4s_v2', 'vcpu': 4, 'ram': 8, 'family': 'compute'},
            {'name': 'Standard_F8s_v2', 'vcpu': 8, 'ram': 16, 'family': 'compute'},
            {'name': 'Standard_F16s_v2', 'vcpu': 16, 'ram': 32, 'family': 'compute'},
            {'name': 'Standard_F32s_v2', 'vcpu': 32, 'ram': 64, 'family': 'compute'},
            {'name': 'Standard_F48s_v2', 'vcpu': 48, 'ram': 96, 'family': 'compute'},
            {'name': 'Standard_F64s_v2', 'vcpu': 64, 'ram': 128, 'family': 'compute'},
            {'name': 'Standard_F72s_v2', 'vcpu': 72, 'ram': 144, 'family': 'compute'},
        ])
        
        # E-Series v5 (Memory Intel) - 1:8 Ratio
        data.extend([
            {'name': 'Standard_E2s_v5', 'vcpu': 2, 'ram': 16, 'family': 'memory'},
            {'name': 'Standard_E4s_v5', 'vcpu': 4, 'ram': 32, 'family': 'memory'},
            {'name': 'Standard_E8s_v5', 'vcpu': 8, 'ram': 64, 'family': 'memory'},
            {'name': 'Standard_E16s_v5', 'vcpu': 16, 'ram': 128, 'family': 'memory'},
            {'name': 'Standard_E32s_v5', 'vcpu': 32, 'ram': 256, 'family': 'memory'},
            {'name': 'Standard_E48s_v5', 'vcpu': 48, 'ram': 384, 'family': 'memory'},
            {'name': 'Standard_E64s_v5', 'vcpu': 64, 'ram': 512, 'family': 'memory'},
            {'name': 'Standard_E96s_v5', 'vcpu': 96, 'ram': 672, 'family': 'memory'},
        ])

        # Eas-Series v5 (Memory AMD EPYC) - 1:8 Ratio
        data.extend([
            {'name': 'Standard_E2as_v5', 'vcpu': 2, 'ram': 16, 'family': 'memory'},
            {'name': 'Standard_E4as_v5', 'vcpu': 4, 'ram': 32, 'family': 'memory'},
            {'name': 'Standard_E8as_v5', 'vcpu': 8, 'ram': 64, 'family': 'memory'},
            {'name': 'Standard_E16as_v5', 'vcpu': 16, 'ram': 128, 'family': 'memory'},
            {'name': 'Standard_E32as_v5', 'vcpu': 32, 'ram': 256, 'family': 'memory'},
            {'name': 'Standard_E48as_v5', 'vcpu': 48, 'ram': 384, 'family': 'memory'},
            {'name': 'Standard_E64as_v5', 'vcpu': 64, 'ram': 512, 'family': 'memory'},
            {'name': 'Standard_E96as_v5', 'vcpu': 96, 'ram': 672, 'family': 'memory'},
        ])
        
        return data
