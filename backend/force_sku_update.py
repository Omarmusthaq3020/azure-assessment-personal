import requests
import json
import os
import sys

# URL for Vantage EC2 Instances Data (The "Scraping" Source)
URL = "https://raw.githubusercontent.com/vantage-sh/ec2instances.info/master/instances.json"
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sku_cache_aws.json')

def main():
    print(f"--- SKU Fetch Diagnostic Tool ---")
    print(f"Target URL: {URL}")
    print("Attempting download...")
    
    try:
        # Check if requests is installed
        import requests
        
        # SSL Verification can sometimes fail behind corporate proxies. 
        # We try strict first, then warn.
        try:
             response = requests.get(URL, timeout=20)
        except requests.exceptions.SSLError:
             print("SSL Error encountered. Retrying without SSL verification (Warning: Insecure)...")
             response = requests.get(URL, timeout=20, verify=False)
             
        print(f"HTTP Status: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            count = len(data)
            print(f"SUCCESS: Downloaded {count} instance definitions.")
            
            # Transform to Map format
            db = {}
            for i in data:
                name = i.get('instance_type')
                if not name: continue
                vcpu = i.get('vCPU', 1)
                ram = i.get('memory', 0)
                fam_raw = i.get('family', '').lower()
                
                # Simple Family Mapping
                fam = 'general'
                if 'burstable' in fam_raw or 'micro' in fam_raw: fam = 'burstable'
                elif 'compute' in fam_raw: fam = 'compute'
                elif 'memory' in fam_raw: fam = 'memory'
                
                db[name] = {'vcpu': vcpu, 'ram': ram, 'family': fam}
                db[f"db.{name}"] = db[name] # RDS Support
            
            print(f"processed {len(db)} SKU keys (including RDS aliases).")
            
            with open(CACHE_FILE, 'w') as f:
                json.dump(db, f, indent=2)
            print(f"Cache saved to: {CACHE_FILE}")
            print("The main assessment tool will now prioritize this data over hardcoded values.")
            
        else:
            print("FAILED: Server returned error.")
            print(response.text[:200])
            
    except ImportError:
        print("CRITICAL: 'requests' library not installed. Cannot fetch web data.")
        print("Run: pip install requests")
    except Exception as e:
        print(f"FAILED: Connectivity Issue. {str(e)}")
        
if __name__ == "__main__":
    main()
