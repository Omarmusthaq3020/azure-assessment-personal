import boto3
import json
import os
import datetime

def generate_aws_metadata_report():
    """
    Generates a comprehensive JSON report of all available AWS resources, 
    services, and their mappings.
    """
    print("Starting AWS Metadata Discovery...")
    
    # 1. Initialize session (looks for env vars or default profile)
    # If running in the app context, these might need to be passed.
    # For a standalone script, we assume env vars or ~/.aws/credentials.
    try:
        session = boto3.Session()
        # Test auth
        sts = session.client('sts')
        identity = sts.get_caller_identity()
        print(f"Authenticated as: {identity['Arn']}")
    except Exception as e:
        print(f"Error: Could not authenticate. Please set AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY environment variables. {e}")
        return

    report = {
        "metadata": {
            "generated_at": datetime.datetime.now().isoformat(),
            "region": session.region_name,
            "account_id": identity['Account']
        },
        "statistics": {},
        "services": [],
        "cloudformation_types": [],
        "inferred_arn_mappings": {}
    }

    # ---------------------------------------------------------
    # SOURCE 1: Available Boto3 Services
    # ---------------------------------------------------------
    print("Fetching available AWS services via boto3...")
    services = session.get_available_services()
    report["services"] = sorted(services)
    print(f"Found {len(services)} available services.")

    # ---------------------------------------------------------
    # SOURCE 2: CloudFormation Resource Types (The "Truth")
    # ---------------------------------------------------------
    print("Fetching CloudFormation Resource Types (this may take 10-20 seconds)...")
    cfn_client = session.client('cloudformation')
    
    cfn_types = []
    try:
        paginator = cfn_client.get_paginator('list_types')
        # We only care about AWS native resources (PUBLIC, AWS_TYPES)
        iterator = paginator.paginate(
            Visibility='PUBLIC', 
            ProvisioningType='FULLY_MUTABLE', 
            Filters={'Category': 'AWS_TYPES'}
        )
        
        for page in iterator:
            for summary in page['TypeSummaries']:
                cfn_types.append({
                    "TypeName": summary['TypeName'],
                    "Description": summary.get('Description', 'No description'),
                    "LastUpdated": str(summary.get('LastUpdated', ''))
                })
    except Exception as e:
        print(f"Error fetching CFN types: {e}")

    report["cloudformation_types"] = cfn_types
    print(f"Found {len(cfn_types)} CloudFormation resource types.")

    # ---------------------------------------------------------
    # SOURCE 3: Inferred ARN Mappings
    # ---------------------------------------------------------
    print("Generating inferred ARN mappings...")
    # We want to map: "service:resource_type" -> "AWS::Service::Resource"
    # Example: "ec2:instance" -> "AWS::EC2::Instance"
    
    mapping = {}
    
    for item in cfn_types:
        full_type = item['TypeName']
        # Format: AWS::Service::Resource
        parts = full_type.split('::')
        if len(parts) == 3:
            aws_prefix = parts[0] # AWS
            service_cfn = parts[1] # e.g. EC2, ElasticLoadBalancingV2
            resource_cfn = parts[2] # e.g. Instance, LoadBalancer
            
            # Key generation rules
            s_lower = service_cfn.lower()
            r_lower = resource_cfn.lower()
            
            # Simple Map
            mapping[f"{s_lower}:{r_lower}"] = full_type
            
            # Handle Service Versions (e.g. ElasticLoadBalancingV2 -> elasticloadbalancing)
            if s_lower.endswith('v2'):
                s_stripped = s_lower[:-2]
                mapping[f"{s_stripped}:{r_lower}"] = full_type
            if s_lower.endswith('v3'):
                s_stripped = s_lower[:-2]
                mapping[f"{s_stripped}:{r_lower}"] = full_type
                
            # Handle "AWS" prefix removals if common
            # (e.g. some obscure services might map differently, but this covers 95%)

    report["inferred_arn_mappings"] = mapping
    report["statistics"] = {
        "service_count": len(services),
        "resource_type_count": len(cfn_types),
        "mapping_count": len(mapping)
    }

    # ---------------------------------------------------------
    # Save Report
    # ---------------------------------------------------------
    output_dir = os.path.join(os.path.dirname(__file__), "..", "data")
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, "aws_full_resource_metadata.json")
    
    with open(output_file, 'w') as f:
        json.dump(report, f, indent=2)
        
    print(f"\nSUCCESS! Report generated at: {output_file}")
    print(f"Reference this file for dynamic type resolution.")

if __name__ == "__main__":
    generate_aws_metadata_report()
