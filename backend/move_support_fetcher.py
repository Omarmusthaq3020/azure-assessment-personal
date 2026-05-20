import requests
import json
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import MoveSupportFetcherConfig

_CACHE_FILE_PATH = MoveSupportFetcherConfig.CACHE_FILE_PATH
if os.path.isabs(_CACHE_FILE_PATH):
    CACHE_FILE = _CACHE_FILE_PATH
else:
    CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.path.basename(_CACHE_FILE_PATH))
CACHE_DURATION_DAYS = MoveSupportFetcherConfig.CACHE_DURATION_DAYS
DOCS_URL = "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
_MOVE_SUPPORT_MEMORY_CACHE = None

def fetch_move_support_data():
    """
    Fetches and parses the Microsoft documentation for Azure resource move support.
    Returns a dictionary where keys are Provider Namespaces (e.g., 'Microsoft.Web')
    and values contain the rules text and source URLs.
    """
    
    global _MOVE_SUPPORT_MEMORY_CACHE

    if _MOVE_SUPPORT_MEMORY_CACHE is not None:
        return _MOVE_SUPPORT_MEMORY_CACHE

    # Check cache first
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
                cache_date = datetime.fromisoformat(cache_data.get('timestamp'))
                if datetime.now() - cache_date < timedelta(days=CACHE_DURATION_DAYS):
                    print("✅ Using cached move support data.")
                    print(f"   Cache contains {len(cache_data.get('data', {}))} providers")
                    _MOVE_SUPPORT_MEMORY_CACHE = cache_data.get('data', {})
                    return _MOVE_SUPPORT_MEMORY_CACHE
        except Exception as e:
            print(f"⚠️ Error reading cache: {e}")

    print("🌐 Fetching fresh move support data from Microsoft Docs...")
    print(f"   URL: {DOCS_URL}")
    try:
        # Fetch the main page
        response = requests.get(DOCS_URL, timeout=30)
        response.raise_for_status()
        print(f"   ✅ Successfully fetched page ({len(response.text)} bytes)")
        
        # Parse with BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Build knowledge base
        knowledge_base = parse_docs_with_links(soup, DOCS_URL)
        
        if not knowledge_base:
            print("   ⚠️ WARNING: No providers found in documentation!")
            return {}
        
        # Save to cache
        os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                'timestamp': datetime.now().isoformat(),
                'data': knowledge_base
            }, f, indent=2)
        
        print(f"   ✅ Cached {len(knowledge_base)} providers")
        _MOVE_SUPPORT_MEMORY_CACHE = knowledge_base
        return _MOVE_SUPPORT_MEMORY_CACHE
        
    except Exception as e:
        print(f"❌ Failed to fetch move support data: {e}")
        import traceback
        traceback.print_exc()
        return {}

def parse_move_support_table(table):
    """
    Parse a move support table and convert to readable text.
    Tables typically have columns: Resource type, Resource group, Subscription, Region move
    """
    try:
        rows = table.find_all('tr')
        if not rows or len(rows) < 2:
            return ""
        
        # Get headers
        header_row = rows[0]
        headers = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]
        
        # Parse data rows
        parsed_items = []
        for row in rows[1:]:
            cells = row.find_all(['td', 'th'])
            if len(cells) >= 2:
                resource_type = cells[0].get_text(strip=True)
                
                # Build readable description
                move_info = []
                for i, cell in enumerate(cells[1:], 1):
                    if i < len(headers):
                        header = headers[i]
                        value = cell.get_text(strip=True)
                        
                        # Convert to readable format
                        if value.lower() == 'yes':
                            move_info.append(f"Can move across {header.lower()}")
                        elif value.lower() == 'no':
                            move_info.append(f"Cannot move across {header.lower()}")
                        else:
                            move_info.append(f"{header}: {value}")
                
                if resource_type and move_info:
                    parsed_items.append(f"• {resource_type}: {', '.join(move_info)}")
        
        if parsed_items:
            return "Move support details:\n" + "\n".join(parsed_items)
        
        return ""
    except Exception as e:
        return ""

def parse_docs_with_links(soup, base_url):
    """
    Parses the HTML to extract provider namespaces, their summary text, and follow links.
    Returns a knowledge base with URL tracking.
    """
    knowledge_base = {}
    
    # Find all h2 headers that contain "Microsoft."
    headers = soup.find_all('h2')
    
    for header in headers:
        header_text = header.get_text(strip=True)
        
        # Check if this is a provider namespace
        if header_text.startswith('Microsoft.'):
            provider = header_text
            
            # Extract content until the next h2
            content_parts = []
            detail_url = None
            table_content = []
            
            # Get all siblings until next h2
            for sibling in header.find_next_siblings():
                if sibling.name == 'h2':
                    break
                
                # Parse tables intelligently
                if sibling.name == 'table':
                    parsed_table = parse_move_support_table(sibling)
                    if parsed_table:
                        table_content.append(parsed_table)
                    continue
                
                # Look for "learn more" or "guidance" links
                if sibling.name in ['p', 'div', 'ul', 'ol', 'section']:
                    links = sibling.find_all('a', href=True)
                    for link in links:
                        link_text = link.get_text(strip=True).lower()
                        if any(keyword in link_text for keyword in ['learn more', 'guidance', 'move', 'limitation']):
                            href = link['href']
                            # Make absolute URL
                            if href.startswith('/'):
                                detail_url = f"https://learn.microsoft.com{href}"
                            elif href.startswith('http'):
                                detail_url = href
                            break
                
                # Extract text from paragraphs, lists, etc. (but not tables)
                if sibling.name in ['p', 'ul', 'ol', 'div', 'section']:
                    # Skip if it's primarily a table container
                    tables = sibling.find_all('table')
                    if tables:
                        # Parse tables and remove them
                        for table in tables:
                            parsed_table = parse_move_support_table(table)
                            if parsed_table:
                                table_content.append(parsed_table)
                            table.decompose()
                    
                    text = sibling.get_text(strip=True)
                    if text and len(text) > 5:  # Very lenient - keep almost everything
                        content_parts.append(text)
            
            # Combine regular text and table content
            summary_text = ' '.join(content_parts)
            if table_content:
                summary_text = f"{summary_text}\n\n{chr(10).join(table_content)}".strip()
            
            # If we found a detail link, fetch it
            detail_text = ""
            if detail_url and MoveSupportFetcherConfig.FOLLOW_DETAIL_LINKS:
                print(f"  📄 Following link for {provider}: {detail_url}")
                detail_text = fetch_detail_page(detail_url)
            
            # Combine summary and detail
            full_text = f"{summary_text}\n\n{detail_text}".strip()
            
            # Fallback: If no text was extracted, provide a generic message
            if not full_text or len(full_text) < 20:
                print(f"  ⚠️ Limited content for {provider}, using fallback")
                full_text = f"Refer to Microsoft documentation for {provider} move support details. Check the Azure portal for specific resource limitations."
            
            knowledge_base[provider] = {
                "summary_url": base_url,
                "detail_url": detail_url or base_url,
                "rules_text": full_text
            }
    
    print(f"✅ Built knowledge base for {len(knowledge_base)} providers.")
    return knowledge_base

def fetch_detail_page(url):
    """
    Fetches a detail page and extracts the main content, skipping tables.
    """
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Try to find the main content area
        main_content = soup.find('main') or soup.find('article') or soup.find('div', class_='content')
        
        if main_content:
            # Remove unwanted elements
            for element in main_content(['script', 'style', 'table', 'nav', 'header', 'footer']):
                element.decompose()
            
            # Extract paragraphs and lists for cleaner text
            text_parts = []
            
            # Get all paragraphs
            for p in main_content.find_all('p'):
                text = p.get_text(strip=True)
                if text and len(text) > 20:  # Skip very short paragraphs
                    text_parts.append(text)
            
            # Get list items (often contain important limitations)
            for ul in main_content.find_all(['ul', 'ol']):
                for li in ul.find_all('li', recursive=False):
                    text = li.get_text(strip=True)
                    if text and len(text) > 10:
                        text_parts.append(f"• {text}")
            
            # Join with newlines and limit size
            full_text = '\n\n'.join(text_parts)
            return full_text[:8000]  # Limit to 8000 chars
        
        return ""
    except Exception as e:
        print(f"  ⚠️ Failed to fetch detail page: {e}")
        return ""

if __name__ == "__main__":
    # Test run
    data = fetch_move_support_data()
    print(f"\nFound {len(data)} providers with documentation.")
    
    # Print a sample
    if 'Microsoft.Compute' in data:
        print(f"\n{'='*60}")
        print(f"Microsoft.Compute Sample:")
        print(f"{'='*60}")
        print(f"Summary URL: {data['Microsoft.Compute']['summary_url']}")
        print(f"Detail URL: {data['Microsoft.Compute']['detail_url']}")
        print(f"Rules Text (first 500 chars):\n{data['Microsoft.Compute']['rules_text'][:500]}...")
