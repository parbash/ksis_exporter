import requests
import re
import csv
import argparse
import sys
import os
from datetime import datetime
from bs4 import BeautifulSoup

# Enable ANSI colors on Windows
if sys.platform == "win32":
    os.system("")  # Enable ANSI escape sequences on Windows 10+

# ANSI color codes
class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

# Global debug flag
DEBUG = False
NAME_CACHE = {}  # Cache for multi-word name decisions

def debug_print(message):
    """Print message only if DEBUG is enabled."""
    if DEBUG:
        print(message)

def safe_find_text(soup, tag, default="Unknown"):
    """Safely extract text from a tag, returning default if not found."""
    element = soup.find(tag)
    return element.get_text().strip() if element else default

def parse_date(date_string):
    """Parse date from various formats, returning ISO format (YYYY-MM-DD)."""
    try:
        # Try to find a date pattern like "DD.MM.YYYY"
        date_match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', date_string)
        if date_match:
            day, month, year = date_match.groups()
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        
        # Try alternative format "YYYY-MM-DD"
        date_match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', date_string)
        if date_match:
            return date_match.group(0)
            
    except Exception as e:
        debug_print(f"Warning: Could not parse date from '{date_string}': {e}")
    
    return "Unknown"

def fetch_url(url):
    """Fetch URL with error handling."""
    try:
        # Add headers to mimic a real browser
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        }
        debug_print(f"  Fetching: {url[:80]}...")
        response = requests.get(url, headers=headers, timeout=10)
        debug_print(f"  Status code: {response.status_code}")
        debug_print(f"  Content-Encoding: {response.headers.get('Content-Encoding', 'none')}")
        response.raise_for_status()
        
        # Force proper decoding
        response.encoding = response.apparent_encoding or 'utf-8'
        
        debug_print(f"  Received {len(response.text)} characters")
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"  {Colors.RED}✗ Error fetching URL: {e}{Colors.RESET}")
        return None

def reorder_name(name):
    """Reorder name from 'Last First' to 'First Last' format.
    
    For names with 3+ words, prompts user to determine the split point.
    Names are displayed in their original "Last First" order for clarity.
    Caches decisions to avoid repeated prompts for the same name.
    """
    global NAME_CACHE
    
    # Normalize whitespace
    name = re.sub(r'\s+', ' ', name).strip()
    
    # Check cache first
    if name in NAME_CACHE:
        return NAME_CACHE[name]
    
    parts = name.split()
    
    if len(parts) == 1:
        # Single name, return as-is
        result = name
    elif len(parts) == 2:
        # Simple case: "Last First" -> "First Last"
        result = f"{parts[1]} {parts[0]}"
    else:
        # 3+ words: ask user where to split
        print(f"\n{Colors.YELLOW}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")
        print(f"{Colors.YELLOW}Multiple-word name detected: {Colors.BOLD}{name}{Colors.RESET}")
        print(f"{Colors.CYAN}This name is in 'Last First' format. Where does the LAST name end?{Colors.RESET}")
        
        # Show options (interpreting as "Last... First..." from left to right)
        for i in range(1, len(parts)):
            last_name = ' '.join(parts[:i])
            first_name = ' '.join(parts[i:])
            reordered = f"{first_name} {last_name}"
            print(f"{Colors.CYAN}{i}.{Colors.RESET} Last: {Colors.BOLD}{last_name}{Colors.RESET}, First: {Colors.BOLD}{first_name}{Colors.RESET} → {Colors.GREEN}{reordered}{Colors.RESET}")
        
        # Get user input
        while True:
            choice = input(f"{Colors.CYAN}Enter choice (1-{len(parts)-1}): {Colors.RESET}").strip()
            
            if choice.isdigit() and 1 <= int(choice) < len(parts):
                split_point = int(choice)
                last_name = ' '.join(parts[:split_point])
                first_name = ' '.join(parts[split_point:])
                result = f"{first_name} {last_name}"
                print(f"{Colors.GREEN}✓ Saved as: {result}{Colors.RESET}")
                print(f"{Colors.YELLOW}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")
                break
            else:
                print(f"{Colors.RED}Invalid choice. Please enter a number between 1 and {len(parts)-1}.{Colors.RESET}")
    
    # Cache the result
    NAME_CACHE[name] = result
    return result

def parse_athlete_row(row_data):
    """Parse athlete information from table row."""
    try:
        # Extract athlete info
        row_ath = row_data[2]
        
        debug_print(f"\n=== Raw athlete cell HTML ===")
        debug_print(row_ath[:500])
        debug_print("=" * 30)
        
        # Split by <br/> or <br>
        parts = re.split(r'<br\s*/?>', row_ath, maxsplit=1)
        if len(parts) < 2:
            return None
        
        row_name = parts[0]
        row_club = parts[1]
        
        # Try to extract name from anchor tag with data attributes or onclick
        # Look for potential data in the link that might indicate first/last name split
        soup_cell = BeautifulSoup(row_ath, 'html.parser')
        link = soup_cell.find('a')
        
        if link:
            # Check for onclick or other attributes that might contain name data
            onclick = link.get('onclick', '')
            href = link.get('href', '')
            title = link.get('title', '')
            
            debug_print(f"Link text: {link.get_text()}")
            debug_print(f"Link onclick: {onclick}")
            debug_print(f"Link href: {href}")
            debug_print(f"Link title: {title}")
            debug_print(f"All link attributes: {link.attrs}")
            
            # Extract the visible name text
            row_name = link.get_text().strip()
        else:
            # No link, just extract text
            name_match = re.search(r'>([^<]+)</a', row_name)
            if name_match:
                row_name = name_match.group(1).strip()
            else:
                row_name = re.sub(r'<.*?>', '', row_name).strip()
        
        # Normalize whitespace
        row_name = re.sub(r'\s+', ' ', row_name)
        
        debug_print(f"Extracted name: {row_name}")
        
        # Reorder name from "Last First" to "First Last"
        row_name = reorder_name(row_name)
        debug_print(f"Reordered name: {row_name}")
        
        # Clean club
        row_club = re.sub(r'<.*?>', '', row_club).strip()
        
        # Extract other fields
        row_yob = re.sub(r'<.*?>', '', row_data[3]).strip()
        row_score = re.sub(r'<.*?>', '', row_data[8]).strip()
        
        return [row_name, row_club, row_yob, row_score]
    
    except (IndexError, AttributeError) as e:
        debug_print(f"Warning: Could not parse row: {e}")
        return None

def get_prop_id():
    """Get and validate prop_id from user input."""
    while True:
        prop_id = input(f"{Colors.CYAN}Enter comp prop_id (ex: 8819): {Colors.RESET}").strip()
        
        if not prop_id:
            print(f"{Colors.RED}⚠ Error: prop_id cannot be empty. Please enter a valid prop_id.{Colors.RESET}")
            continue
        
        if not prop_id.isdigit():
            print(f"{Colors.RED}⚠ Error: prop_id must be a number. Please try again.{Colors.RESET}")
            continue
        
        return prop_id

def list_competitions(live_only=False, search_keyword=None):
    """Fetch and display list of competitions from ksis.eu.
    
    Args:
        live_only: If True, only show competitions with LIVE sessions
        search_keyword: If provided, filter competitions by keyword
    """
    filter_desc = ""
    if live_only:
        filter_desc = " (Live competitions only)"
    elif search_keyword:
        filter_desc = f" (Searching for: '{search_keyword}')"
    
    print(f"\n{Colors.CYAN}Fetching competition list (Canada - Women's Artistic Gymnastics){filter_desc}...{Colors.RESET}")
    
    url = "https://ksis.eu/menu.php?akcia=S&oblast=ARTW&country=CAN"
    html_content = fetch_url(url)
    
    if not html_content:
        print(f"{Colors.RED}Failed to fetch competition list.{Colors.RESET}")
        return []
    
    soup = BeautifulSoup(html_content, "html.parser")
    
    # Find all competition links
    all_competitions = []
    links = soup.find_all('a', href=True)
    
    for link in links:
        href = link.get('href')
        # Look for links with id_prop parameter
        match = re.search(r'id_prop=(\d+)', href)
        if match:
            prop_id = match.group(1)
            comp_name = link.get_text().strip()
            
            # Check if this link has a LIVE badge (span with class badge containing "live")
            is_live = False
            parent = link.parent
            if parent:
                badges = parent.find_all('span', class_='badge')
                for badge in badges:
                    if 'live' in badge.get_text().lower():
                        is_live = True
                        break
            
            # Skip empty names or duplicates
            if comp_name and not any(c['prop_id'] == prop_id for c in all_competitions):
                all_competitions.append({
                    'prop_id': prop_id,
                    'name': comp_name,
                    'is_live': is_live
                })
    
    # Apply filters
    competitions = all_competitions
    
    if live_only:
        competitions = [c for c in competitions if c['is_live']]
    
    if search_keyword:
        keyword_lower = search_keyword.lower()
        competitions = [c for c in competitions if keyword_lower in c['name'].lower()]
    
    if not competitions:
        if live_only:
            print(f"{Colors.YELLOW}No live competitions found.{Colors.RESET}")
        elif search_keyword:
            print(f"{Colors.YELLOW}No competitions found matching '{search_keyword}'.{Colors.RESET}")
        else:
            print(f"{Colors.YELLOW}No competitions found.{Colors.RESET}")
        return []
    
    # Display competitions
    print(f"\n{Colors.BOLD}Available Competitions:{Colors.RESET}")
    print(f"{Colors.CYAN}{'ID':<8} {'Competition Name'}{Colors.RESET}")
    print("-" * 80)
    
    for comp in competitions:
        live_indicator = f" {Colors.RED}[LIVE]{Colors.RESET}" if comp['is_live'] else ""
        print(f"{Colors.GREEN}{comp['prop_id']:<8}{Colors.RESET} {comp['name']}{live_indicator}")
    
    print(f"\n{Colors.BOLD}Total: {len(competitions)} competitions{Colors.RESET}")
    
    return competitions

def export_results(prop_id):
    """Export results for a given prop_id."""
    url = f"https://ksis.eu/resultx.php?id_prop={prop_id}"
    
    print(f"\n{Colors.CYAN}Fetching competition data (prop_id: {prop_id})...{Colors.RESET}")
    html_content = fetch_url(url)
    
    if not html_content:
        print(f"{Colors.RED}Failed to fetch competition page. Exiting...{Colors.RESET}")
        return
    
    debug_print(f"✓ Received {len(html_content)} characters of HTML")
    
    soup = BeautifulSoup(html_content, "html.parser")
    
    if DEBUG:
        # Debug: Save the HTML to check what we're receiving
        with open('debug_output.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        debug_print("Saved HTML to debug_output.html for inspection")
        
        # Show first 500 characters of what we received
        debug_print(f"\nFirst 500 characters of response:")
        debug_print("-" * 60)
        debug_print(html_content[:500])
        debug_print("-" * 60)
        debug_print(f"\nSearching for key content...")
        
        # Look for common error messages or indicators
        lower_content = html_content.lower()
        if 'error' in lower_content:
            debug_print("⚠ Found 'error' in page content")
        if 'not found' in lower_content or '404' in lower_content:
            debug_print("⚠ Page indicates content not found")
        if 'competition' in lower_content:
            debug_print("✓ Found 'competition' in page")
        if 'result' in lower_content:
            debug_print("✓ Found 'result' in page")
    
    # Extract competition name
    comp_name = safe_find_text(soup, 'h3', 'Unknown_Competition')
    comp_name = comp_name.replace("/", "-").replace("\\", "-")
    
    # Extract and parse date
    h4_text = safe_find_text(soup, 'h4', '')
    debug_print(f"H4 text: '{h4_text}'")
    comp_date = parse_date(h4_text)
    debug_print(f"Parsed date: {comp_date}")
    
    # Get list of sessions
    print(f'{Colors.CYAN}Parsing sessions...{Colors.RESET}')
    select_tag = soup.find('select', id='id_sut')
    
    if not select_tag:
        print(f"{Colors.RED}✗ Select tag not found.{Colors.RESET}")
        
        # Check if it's because sessions are in progress
        lower_content = html_content.lower()
        if 'in progress' in lower_content or 'live' in lower_content or comp_name != 'Unknown_Competition':
            print(f"{Colors.YELLOW}ℹ The competition sessions may still be in progress.{Colors.RESET}")
            print(f"{Colors.YELLOW}  Results will be available once the sessions are completed.{Colors.RESET}")
        else:
            print(f"{Colors.YELLOW}  Wrong prop_id or the website structure has changed.{Colors.RESET}")
        
        if DEBUG:
            print("\nDebugging information:")
            print(f"  - Page title: {soup.find('title').get_text() if soup.find('title') else 'Not found'}")
            print(f"  - Page contains {len(soup.find_all())} total HTML elements")
            print(f"  - H3 tags found: {len(soup.find_all('h3'))}")
            print(f"  - H4 tags found: {len(soup.find_all('h4'))}")
            print(f"  - Select tags found: {len(soup.find_all('select'))}")
            print("\nCheck debug_output.html to see what the page actually returned.")
            print("The site may require cookies, JavaScript, or specific session handling.")
        else:
            print(f"{Colors.YELLOW}Run with --debug flag for more information.{Colors.RESET}")
        return
    
    all_data = []
    in_progress_count = 0
    options = select_tag.find_all('option')
    
    if not options:
        print(f"{Colors.YELLOW}No sessions found. Exiting...{Colors.RESET}")
        return
    
    for option in options:
        value = option.get('value')
        session_name = option.get_text().strip()
        
        result_url = (
            f"https://ksis.eu/load_result_total_ksismg_art.php?"
            f"lang=en&id_prop={prop_id}&id_sut={value}&rn=null&mn=null"
            f"&state=-1&age_group=&award=-1&nacinie=undefined"
        )
        
        result_content = fetch_url(result_url)
        if not result_content:
            in_progress_count += 1
            print(f"{Colors.YELLOW}  ✗ {session_name}: Could not fetch results{Colors.RESET}")
            continue
        
        soup2 = BeautifulSoup(result_content, "html.parser")
        table = soup2.find('table', attrs={'id': 'myTablePrihlasky'})
        
        if not table:
            debug_print(f"    Warning: No results table found for session {session_name}")
            in_progress_count += 1
            print(f"{Colors.YELLOW}  ⚠ {session_name}: No results table found (session may be in progress){Colors.RESET}")
            continue
        
        # Parse data from table
        row_count = 0
        for row in table.find_all('tr'):
            cells = row.find_all(['td'])
            if len(cells) < 9:  # Ensure we have enough cells
                continue
            
            row_data = [str(cell) for cell in cells]
            athlete_data = parse_athlete_row(row_data)
            
            if athlete_data:
                all_data.append([session_name] + athlete_data + [comp_date])
                row_count += 1
        
        # Check if session is in progress (has LIVE in name or 0 athletes)
        if row_count == 0:
            in_progress_count += 1
            if 'live' in session_name.lower() or 'in progress' in session_name.lower():
                print(f"{Colors.YELLOW}  ⚠ {session_name}: Session is in progress{Colors.RESET}")
            else:
                print(f"{Colors.YELLOW}  ⚠ {session_name}: No athletes found (may be in progress){Colors.RESET}")
        else:
            print(f"  {Colors.GREEN}✓{Colors.RESET} {session_name}: {row_count} athletes")
    
    # Write CSV file
    if all_data:
        # Add timestamp to filename
        timestamp = datetime.now().strftime('%Y%m%d%H%M')
        file_path = f'{comp_name}-{timestamp}.csv'
        
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile, delimiter=',')
            writer.writerow(['Session', 'Name', 'Club', 'YOB', 'Score', 'Date'])  # Header
            writer.writerows(all_data)
        print(f"\n{Colors.GREEN}{Colors.BOLD}✓ Successfully created {file_path} with {len(all_data)} records.{Colors.RESET}")
        
        # Show in-progress sessions count if any
        if in_progress_count > 0:
            print(f"{Colors.MAGENTA}{Colors.BOLD}ℹ {in_progress_count} session(s) still in progress{Colors.RESET}")
    else:
        print(f"\n{Colors.YELLOW}No data was collected. No file created.{Colors.RESET}")
        if in_progress_count > 0:
            print(f"{Colors.MAGENTA}{Colors.BOLD}ℹ {in_progress_count} session(s) still in progress{Colors.RESET}")

def interactive_menu():
    """Display interactive menu and handle user choices."""
    while True:
        print(f"\n{Colors.BOLD}{Colors.CYAN}╔═══════════════════════════════════════╗{Colors.RESET}")
        print(f"{Colors.BOLD}{Colors.CYAN}║    KSIS Competition Results Tool      ║{Colors.RESET}")
        print(f"{Colors.BOLD}{Colors.CYAN}╚═══════════════════════════════════════╝{Colors.RESET}")
        print(f"\n{Colors.BOLD}Select an option:{Colors.RESET}")
        print(f"{Colors.CYAN}1.{Colors.RESET} List all competitions")
        print(f"{Colors.CYAN}2.{Colors.RESET} List live competitions only")
        print(f"{Colors.CYAN}3.{Colors.RESET} Search competitions by keyword")
        print(f"{Colors.CYAN}4.{Colors.RESET} Export results by prop_id")
        print(f"{Colors.CYAN}5.{Colors.RESET} Exit")
        
        choice = input(f"\n{Colors.CYAN}Enter your choice (1-5): {Colors.RESET}").strip()
        
        if choice == '1':
            competitions = list_competitions()
            if competitions:
                print(f"\n{Colors.YELLOW}Tip: Copy a prop_id and use option 4 to export results{Colors.RESET}")
        
        elif choice == '2':
            competitions = list_competitions(live_only=True)
            if competitions:
                print(f"\n{Colors.YELLOW}Tip: Copy a prop_id and use option 4 to export results{Colors.RESET}")
        
        elif choice == '3':
            keyword = input(f"{Colors.CYAN}Enter search keyword: {Colors.RESET}").strip()
            if keyword:
                competitions = list_competitions(search_keyword=keyword)
                if competitions:
                    print(f"\n{Colors.YELLOW}Tip: Copy a prop_id and use option 4 to export results{Colors.RESET}")
            else:
                print(f"{Colors.RED}No keyword provided.{Colors.RESET}")
        
        elif choice == '4':
            prop_id = get_prop_id()
            export_results(prop_id)
        
        elif choice == '5':
            print(f"\n{Colors.CYAN}Goodbye!{Colors.RESET}")
            break
        
        else:
            print(f"{Colors.RED}Invalid choice. Please enter 1-5.{Colors.RESET}")

def main():
    global DEBUG
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Parse KSIS competition results and export to CSV',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python ksis_export.py              # Interactive menu mode
  python ksis_export.py --list       # List competitions and exit
  python ksis_export.py --prop-id 8819  # Export specific competition
  python ksis_export.py --debug --prop-id 8819  # Export with debug output
        """
    )
    parser.add_argument(
        '--debug', '-d',
        action='store_true',
        help='Enable debug output for troubleshooting'
    )
    parser.add_argument(
        '--prop-id',
        type=str,
        help='Competition prop_id to export (skips interactive menu)'
    )
    parser.add_argument(
        '--list', '-l',
        action='store_true',
        help='List available competitions and exit'
    )
    
    args = parser.parse_args()
    DEBUG = args.debug
    
    # Handle different modes
    if args.list:
        # Just list competitions and exit
        list_competitions()
    elif args.prop_id:
        # Export specific competition
        prop_id = args.prop_id.strip()
        if not prop_id.isdigit():
            print(f"{Colors.RED}Error: prop_id must be a number{Colors.RESET}")
            return
        export_results(prop_id)
    else:
        # Interactive menu mode
        interactive_menu()

if __name__ == "__main__":
    main()
