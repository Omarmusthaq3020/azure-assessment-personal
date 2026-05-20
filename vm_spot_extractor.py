"""
VM Spot Instance Extractor for Azure Assessment Tool
====================================================
Extracts Azure Spot VM details including eviction policy, eviction type, and max price.
"""

def extract_spot_vm_details(resources):
    """
    Extract Spot VM details for all virtual machines.
    
    Args:
        resources: List of resource dictionaries
    """
    print("💰 Extracting VM Spot Instance details...")
    
    spot_vm_count = 0
    for resource in resources:
        if resource.get("type", "").lower() != "microsoft.compute/virtualmachines":
            continue
        
        # Get VM properties
        properties = resource.get("properties", {})
        
        # Check if VM is a Spot instance
        priority = properties.get("priority", "Regular")
        eviction_policy = properties.get("evictionPolicy", "N/A")
        billing_profile = properties.get("billingProfile", {})
        max_price = billing_profile.get("maxPrice", "N/A")
        
        # Determine if it's a Spot VM
        is_spot = priority == "Spot"
        
        # Add Spot VM details to resource
        resource["Azure Spot"] = "Yes" if is_spot else "No"
        resource["VM Priority"] = priority
        resource["Eviction Policy"] = eviction_policy if is_spot else "N/A"
        resource["Eviction Type"] = get_eviction_type(eviction_policy) if is_spot else "N/A"
        resource["Max Price"] = format_max_price(max_price) if is_spot else "N/A"
        
        if is_spot:
            spot_vm_count += 1
    
    print(f"   ✅ Found {spot_vm_count} Spot VMs")


def get_eviction_type(eviction_policy):
    """
    Determine eviction type based on eviction policy.
    
    Args:
        eviction_policy: The eviction policy string
        
    Returns:
        str: Human-readable eviction type
    """
    if eviction_policy == "Deallocate":
        return "Capacity-based (Deallocate)"
    elif eviction_policy == "Delete":
        return "Price or Capacity-based (Delete)"
    else:
        return "Unknown"


def format_max_price(max_price):
    """
    Format max price for display.
    
    Args:
        max_price: The max price value
        
    Returns:
        str: Formatted price string
    """
    if max_price == -1 or max_price == "-1":
        return "Pay up to on-demand price"
    elif max_price == "N/A" or max_price is None:
        return "N/A"
    else:
        try:
            price_float = float(max_price)
            return f"${price_float:.4f} per hour"
        except (ValueError, TypeError):
            return str(max_price)


if __name__ == "__main__":
    # Test with sample data
    test_resources = [
        {
            "type": "Microsoft.Compute/virtualMachines",
            "name": "spot-vm-01",
            "properties": {
                "priority": "Spot",
                "evictionPolicy": "Deallocate",
                "billingProfile": {
                    "maxPrice": 0.05
                }
            }
        },
        {
            "type": "Microsoft.Compute/virtualMachines",
            "name": "regular-vm-01",
            "properties": {
                "priority": "Regular"
            }
        },
        {
            "type": "Microsoft.Compute/virtualMachines",
            "name": "spot-vm-02",
            "properties": {
                "priority": "Spot",
                "evictionPolicy": "Delete",
                "billingProfile": {
                    "maxPrice": -1
                }
            }
        }
    ]
    
    extract_spot_vm_details(test_resources)
    
    print("\n📊 Results:")
    for res in test_resources:
        if res.get("type", "").lower() == "microsoft.compute/virtualmachines":
            print(f"\n   VM: {res['name']}")
            print(f"   Azure Spot: {res.get('Azure Spot', 'N/A')}")
            print(f"   Priority: {res.get('VM Priority', 'N/A')}")
            print(f"   Eviction Policy: {res.get('Eviction Policy', 'N/A')}")
            print(f"   Eviction Type: {res.get('Eviction Type', 'N/A')}")
            print(f"   Max Price: {res.get('Max Price', 'N/A')}")
