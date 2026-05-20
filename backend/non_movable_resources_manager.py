"""
Dynamic Non-Movable Resources Manager
======================================
Manages the non-movable resources list with automatic caching and updates.

This module:
1. Checks if the cached JSON is stale (older than configured days)
2. If stale, triggers the scraper to update the data
3. Falls back to cached data if scraping fails
4. Provides a simple interface to load non-movable resources
"""

import os
import json
from datetime import datetime, timedelta
import sys

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import NonMovableResourcesConfig


def is_cache_stale(json_path: str, cache_duration_days: int = 7) -> bool:
    """
    Check if the cached JSON file is stale.
    
    Args:
        json_path: Path to the JSON file
        cache_duration_days: Number of days before cache is considered stale
    
    Returns:
        True if cache is stale or doesn't exist, False otherwise
    """
    if not os.path.exists(json_path):
        print(f"📭 Cache file not found: {json_path}")
        return True
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Check if metadata exists
        if 'metadata' not in data or 'generated_at' not in data['metadata']:
            print("⚠️ Cache file missing metadata, treating as stale")
            return True
        
        # Parse the timestamp
        generated_at_str = data['metadata']['generated_at']
        # Handle both ISO format with and without 'Z'
        if generated_at_str.endswith('Z'):
            generated_at_str = generated_at_str[:-1]
        
        generated_at = datetime.fromisoformat(generated_at_str)
        age = datetime.utcnow() - generated_at
        
        is_stale = age > timedelta(days=cache_duration_days)
        
        if is_stale:
            print(f"📅 Cache is {age.days} days old (max: {cache_duration_days} days) - stale")
        else:
            print(f"✅ Cache is {age.days} days old (max: {cache_duration_days} days) - fresh")
        
        return is_stale
    
    except Exception as e:
        print(f"⚠️ Error checking cache age: {e}")
        return True


def update_non_movable_resources() -> bool:
    """
    Trigger the scraper to update non-movable resources.
    
    Returns:
        True if update was successful, False otherwise
    """
    try:
        print("🔄 Updating non-movable resources from Microsoft docs...")
        
        # Import the scraper module
        from backend import scrape_move_support
        
        # Run the scraper
        non_movable_resources = scrape_move_support.fetch_move_support_table()
        
        if not non_movable_resources:
            print("❌ Failed to fetch data from Microsoft docs")
            return False
        
        # Generate the JSON file at configured cache location
        output_path = NonMovableResourcesConfig.JSON_FILE_PATH
        if not os.path.isabs(output_path):
            script_dir = os.path.dirname(os.path.abspath(__file__))
            output_path = os.path.join(script_dir, os.path.basename(output_path))
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        scrape_move_support.generate_json_file(non_movable_resources, output_path)
        
        print(f"✅ Successfully updated non-movable resources")
        return True
    
    except Exception as e:
        print(f"❌ Failed to update non-movable resources: {e}")
        import traceback
        traceback.print_exc()
        return False


def load_non_movable_resources(force_update: bool = False):
    """
    Load non-movable resources with automatic cache management.
    
    This function:
    1. Checks if cache is stale
    2. Updates if stale (and auto-update is enabled)
    3. Falls back to cached data if update fails
    4. Returns the non-movable resources list
    
    Args:
        force_update: If True, force an update regardless of cache age
    
    Returns:
        List of non-movable resources in the format:
        [
            {"namespace": "Microsoft.Compute", "types": ["type1", "type2"]},
            ...
        ]
    """
    # Get configuration
    json_path = NonMovableResourcesConfig.JSON_FILE_PATH
    cache_duration = NonMovableResourcesConfig.CACHE_DURATION_DAYS
    auto_update_enabled = NonMovableResourcesConfig.AUTO_UPDATE_ENABLED
    
    # Make path absolute
    if not os.path.isabs(json_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        json_path = os.path.join(script_dir, os.path.basename(json_path))
    
    # Check if update is needed.
    # Always attempt initial generation when cache file is missing, even if auto-update is disabled.
    cache_missing = not os.path.exists(json_path)
    needs_update = force_update or cache_missing or (auto_update_enabled and is_cache_stale(json_path, cache_duration))
    
    if needs_update:
        print(f"🔄 Cache update needed (force={force_update}, auto_update={auto_update_enabled})")
        
        # Try to update
        update_success = update_non_movable_resources()
        
        if update_success:
            print("✅ Using FRESHLY UPDATED data from Microsoft docs")
            sys.stdout.flush()
        else:
            print("⚠️ Update failed, will use CACHED data if available")
            sys.stdout.flush()
    else:
        print(f"💾 Using CACHED data (fresh, no update needed)")
        sys.stdout.flush()
    
    # Load from JSON file (either fresh or cached)
    try:
        if not os.path.exists(json_path):
            print(f"❌ No cache file found at: {json_path}")
            print("   Please run scrape_move_support.py manually to create initial cache")
            return []
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Handle both old format (array) and new format (object with metadata)
        if isinstance(data, list):
            # Old format: direct array
            print(f"📋 Loaded {len(data)} namespaces (old format)")
            return data
        elif isinstance(data, dict) and 'non_movable_resources' in data:
            # New format: object with metadata
            resources = data['non_movable_resources']
            metadata = data.get('metadata', {})
            
            # Parse age for display
            generated_at_str = metadata.get('generated_at', '')
            if generated_at_str:
                if generated_at_str.endswith('Z'):
                    generated_at_str = generated_at_str[:-1]
                try:
                    generated_at = datetime.fromisoformat(generated_at_str)
                    age_hours = (datetime.utcnow() - generated_at).total_seconds() / 3600
                    if age_hours < 1:
                        age_display = f"{int(age_hours * 60)} minutes ago"
                    elif age_hours < 24:
                        age_display = f"{int(age_hours)} hours ago"
                    else:
                        age_display = f"{int(age_hours / 24)} days ago"
                except:
                    age_display = "unknown"
            else:
                age_display = "unknown"
            
            print(f"📋 Loaded non-movable resources:")
            print(f"   - Total namespaces: {metadata.get('total_namespaces', len(resources))}")
            print(f"   - Total types: {metadata.get('total_resource_types', 'unknown')}")
            print(f"   - Cache age: {age_display}")
            
            return resources
        else:
            print(f"⚠️ Unexpected JSON format in {json_path}")
            return []
    
    except Exception as e:
        print(f"❌ Failed to load non-movable resources: {e}")
        import traceback
        traceback.print_exc()
        return []


def get_cache_info() -> dict:
    """
    Get information about the current cache.
    
    Returns:
        Dictionary with cache information
    """
    json_path = NonMovableResourcesConfig.JSON_FILE_PATH
    
    if not os.path.isabs(json_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        json_path = os.path.join(script_dir, os.path.basename(json_path))
    
    if not os.path.exists(json_path):
        return {
            "exists": False,
            "path": json_path
        }
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        metadata = data.get('metadata', {})
        
        generated_at_str = metadata.get('generated_at', '')
        if generated_at_str.endswith('Z'):
            generated_at_str = generated_at_str[:-1]
        
        if generated_at_str:
            generated_at = datetime.fromisoformat(generated_at_str)
            age_days = (datetime.utcnow() - generated_at).days
        else:
            age_days = None
        
        return {
            "exists": True,
            "path": json_path,
            "generated_at": metadata.get('generated_at'),
            "age_days": age_days,
            "is_stale": is_cache_stale(json_path, NonMovableResourcesConfig.CACHE_DURATION_DAYS),
            "total_namespaces": metadata.get('total_namespaces'),
            "total_resource_types": metadata.get('total_resource_types'),
            "source": metadata.get('source')
        }
    
    except Exception as e:
        return {
            "exists": True,
            "path": json_path,
            "error": str(e)
        }


if __name__ == "__main__":
    """
    Command-line interface for managing non-movable resources cache.
    
    Usage:
        python non_movable_resources_manager.py           # Load with auto-update
        python non_movable_resources_manager.py --update  # Force update
        python non_movable_resources_manager.py --info    # Show cache info
    """
    import argparse
    
    parser = argparse.ArgumentParser(description="Manage non-movable resources cache")
    parser.add_argument('--update', action='store_true', help='Force update from Microsoft docs')
    parser.add_argument('--info', action='store_true', help='Show cache information')
    args = parser.parse_args()
    
    if args.info:
        print("=" * 70)
        print("Non-Movable Resources Cache Information")
        print("=" * 70)
        info = get_cache_info()
        for key, value in info.items():
            print(f"{key}: {value}")
        print("=" * 70)
    elif args.update:
        print("=" * 70)
        print("Forcing Update of Non-Movable Resources")
        print("=" * 70)
        resources = load_non_movable_resources(force_update=True)
        print(f"\n✅ Loaded {len(resources)} namespaces")
    else:
        print("=" * 70)
        print("Loading Non-Movable Resources (with auto-update)")
        print("=" * 70)
        resources = load_non_movable_resources()
        print(f"\n✅ Loaded {len(resources)} namespaces")
