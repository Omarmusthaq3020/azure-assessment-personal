import boto3
import sys

# Use default credentials or environment
try:
    rds = boto3.client('rds', region_name='ap-south-1')
    print("Scanning RDS in ap-south-1...")
    
    resp = rds.describe_db_instances()
    dbs = resp['DBInstances']
    print(f"Found {len(dbs)} DB Instances:")
    
    for db in dbs:
        print(f"- Name: {db['DBInstanceIdentifier']}")
        print(f"  Status: {db['DBInstanceStatus']}")
        print(f"  Engine: {db['Engine']}")
        print(f"  ARN:    {db['DBInstanceArn']}")
        print("-" * 30)

except Exception as e:
    print(f"Error: {e}")
