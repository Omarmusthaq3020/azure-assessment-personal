"""
Scrape Microsoft's official move support documentation to build an accurate
list of resources that cannot be moved across subscriptions.

USAGE:
1. Make sure you have the required packages installed:
   pip install requests beautifulsoup4

2. Run this script:
   python scrape_move_support.py

3. The script will generate 'non_movable_resources.json' in the same directory

This script fetches the official table from Microsoft Learn and generates
a clean JSON file containing only resources with "Subscription Move = No".

Run this script periodically (e.g., weekly) to keep the deny list up to date.
"""

import json
from datetime import datetime
import os

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("❌ Missing required packages. Please install:")
    print("   pip install requests beautifulsoup4")
    exit(1)

def fetch_move_support_table():
    """Fetch and parse the official Microsoft move support table."""
    url = "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
    
    print(f"🌐 Fetching move support data from Microsoft Learn...")
    print(f"   URL: {url}")
    
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        print(f"✅ Successfully fetched page ({len(response.content)} bytes)")
    except Exception as e:
        print(f"❌ Failed to fetch page: {e}")
        return None
    
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Find all headings that look like namespaces
    all_headings = soup.find_all(['h2', 'h3'])
    namespace_headings = []
    for h in all_headings:
        text = h.get_text(strip=True)
        if text.startswith('Microsoft.'):
            namespace_headings.append((h, text))
    
    print(f"📋 Found {len(namespace_headings)} namespace headings")
    
    # Find all tables
    tables = soup.find_all('table')
    print(f"📊 Found {len(tables)} tables on the page")
    
    non_movable_resources = {}
    total_processed = 0
    
    # For each namespace heading, find the next table(s) after it
    for heading, namespace in namespace_headings:
        try:
            # Find all tables that come after this heading
            current = heading
            namespace_tables = []
            
            # Walk forward through siblings and descendants
            for sibling in heading.find_all_next():
                # Stop if we hit another namespace heading
                if sibling.name in ['h2', 'h3'] and sibling.get_text(strip=True).startswith('Microsoft.'):
                    break
                # Collect tables
                if sibling.name == 'table':
                    namespace_tables.append(sibling)
            
            # Process tables for this namespace
            for table in namespace_tables:
                rows = table.find_all('tr')
                
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) < 3:
                        continue
                    
                    # Column 0: Resource type
                    # Column 2: Subscription move
                    resource_type = cells[0].get_text(strip=True)
                    subscription_move = cells[2].get_text(strip=True).lower()
                    
                    # Check if subscription move is "No"
                    if subscription_move == 'no' and resource_type:
                        if namespace not in non_movable_resources:
                            non_movable_resources[namespace] = []
                        
                        if resource_type not in non_movable_resources[namespace]:
                            non_movable_resources[namespace].append(resource_type)
                            total_processed += 1
        
        except Exception as e:
            print(f"⚠️ Error processing namespace {namespace}: {e}")
            continue
    
    print(f"✅ Processed {total_processed} non-movable resource types across {len(non_movable_resources)} namespaces")
    return non_movable_resources

def generate_json_file(non_movable_resources, output_path):
    """Generate the JSON file with non-movable resources."""
    
    # Convert to the expected format
    result = []
    for namespace in sorted(non_movable_resources.keys()):
        types = sorted(list(set(non_movable_resources[namespace])))  # Remove duplicates
        result.append({
            "namespace": namespace,
            "types": types
        })
    
    # Add metadata
    output = {
        "metadata": {
            "generated_at": datetime.utcnow().isoformat() + "Z",
            "source": "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources",
            "description": "Resources that cannot be moved across Azure subscriptions",
            "total_namespaces": len(result),
            "total_resource_types": sum(len(item["types"]) for item in result)
        },
        "non_movable_resources": result
    }
    
    # Write to file
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"\n✅ Successfully generated {output_path}")
    print(f"   📦 Total namespaces: {output['metadata']['total_namespaces']}")
    print(f"   📦 Total resource types: {output['metadata']['total_resource_types']}")
    print(f"   🕒 Generated at: {output['metadata']['generated_at']}")
    
    return output

def main():
    """Main execution function."""
    print("=" * 70)
    print("🔧 Azure Move Support Scraper")
    print("=" * 70)
    print()
    
    # Fetch the data
    non_movable_resources = fetch_move_support_table()
    
    if not non_movable_resources:
        print("\n❌ Failed to fetch move support data. Exiting.")
        return
    
    print(f"\n📋 Found {len(non_movable_resources)} namespaces with non-movable resources")
    
    # Generate the JSON file in the same directory as this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, "non_movable_resources.json")
    
    result = generate_json_file(non_movable_resources, output_path)
    
    # Display sample
    print("\n📄 Sample entries:")
    for item in result["non_movable_resources"][:5]:
        print(f"   • {item['namespace']}: {len(item['types'])} types")
        for type_name in item['types'][:3]:
            print(f"     - {type_name}")
        if len(item['types']) > 3:
            print(f"     ... and {len(item['types']) - 3} more")
    
    print("\n" + "=" * 70)
    print("✅ Scraping complete!")
    print("   📁 Output file: " + output_path)
    print("=" * 70)

if __name__ == "__main__":
    main()
