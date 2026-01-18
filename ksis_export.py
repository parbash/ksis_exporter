import requests
import re
import csv
import argparse
import sys
import os
import time
from datetime import datetime
from bs4 import BeautifulSoup

# Note: Pandas is imported only when needed to speed up script startup

# Enable ANSI colors on Windows
if sys.platform == "win32":
    os.system("")

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

# Global flags and caches
DEBUG = False
NAME_CACHE = {}
CLUB_CORRECTIONS = {}
ATHLETE_CORRECTIONS = {}

# Configuration for Columns
# 1. DROP_COLS: These columns are completely removed from the CSV
#    Removed 'SV' from here so we can capture it
DROP_COLS = ['E', 'Bonus', 'Comp', 'ND', 'D'] 

# 2. RENAME_MAP: Renames headers found in HTML -> CSV Header
#    Added 'SV': 'Score' back. This forces the data found in the SV column to appear under 'Score'
RENAME_MAP = {
    'Total': 'Score',
    'SV': 'Score',    # Maps SV column data to the Score column
    'Born': 'YOB', 
    'born': 'YOB'
} 

def debug_print(message):
    """Print message only if DEBUG is enabled."""
    if DEBUG:
        print(f"{Colors.MAGENTA}[DEBUG] {message}{Colors.RESET}")

# ==========================================
#  FILE LOADING / SAVING FUNCTIONS
# ==========================================

def load_excel_corrections(filename):
    """Generic helper to load corrections from Excel (Col A -> Col B)."""
    if not os.path.exists(filename):
        return {}
    
    try:
        import pandas as pd
        # Force reading only the first two columns to prevent "shifting" data
        df = pd.read_excel(filename, header=None, usecols=[0, 1])
        
        if df.empty: return {}
        
        # Manually assign headers to ensure consistency
        df.columns = ['Original', 'Corrected']
        
        # If the file had headers "Original/Corrected", remove that row
        if str(df.iloc[0]['Original']).strip().lower() == 'original':
            df = df.iloc[1:]

        return dict(zip(
            df['Original'].astype(str).str.strip(), 
            df['Corrected'].astype(str).str.strip()
        ))
    except Exception as e:
        debug_print(f"Error loading {filename}: {e}")
        return {}

def load_corrections():
    """Load both Club and Athlete corrections."""
    global CLUB_CORRECTIONS, ATHLETE_CORRECTIONS
    
    debug_print("Loading corrections dictionaries...")
    
    CLUB_CORRECTIONS = load_excel_corrections("Club Name Corrections.xlsx")
    if CLUB_CORRECTIONS and DEBUG:
        debug_print(f"✓ Loaded {len(CLUB_CORRECTIONS)} club corrections.")
        
    ATHLETE_CORRECTIONS = load_excel_corrections("Athlete Name Corrections.xlsx")
    if ATHLETE_CORRECTIONS and DEBUG:
        debug_print(f"✓ Loaded {len(ATHLETE_CORRECTIONS)} athlete name corrections.")

def save_athlete_correction(original, corrected):
    """Appends a new athlete correction to the Excel file."""
    filename = "Athlete Name Corrections.xlsx"
    try:
        import pandas as pd
        new_row = pd.DataFrame([[original, corrected]], columns=['Original', 'Corrected'])
        
        if os.path.exists(filename):
            existing_df = pd.read_excel(filename, header=None, usecols=[0, 1])
            if not existing_df.empty:
                existing_df.columns = ['Original', 'Corrected']
                if str(existing_df.iloc[0]['Original']).strip().lower() == 'original':
                    existing_df = existing_df.iloc[1:]
            else:
                existing_df = pd.DataFrame(columns=['Original', 'Corrected'])
            updated_df = pd.concat([existing_df, new_row], ignore_index=True)
        else:
            updated_df = new_row
            
        updated_df.to_excel(filename, index=False, header=True)
        print(f"{Colors.GREEN}✓ Saved correction to '{filename}'{Colors.RESET}")
        
        # Update memory
        ATHLETE_CORRECTIONS[original] = corrected
    except Exception as e:
        print(f"{Colors.RED}⚠ Error saving correction: {e}{Colors.RESET}")

# ==========================================
#  STANDARDIZATION LOGIC
# ==========================================

def standardize_club(club_name):
    if not club_name: return ""
    clean_name = club_name.strip()
    
    # 1. Dictionary Check
    if clean_name in CLUB_CORRECTIONS:
        clean_name = CLUB_CORRECTIONS[clean_name].strip()
        
    # 2. Suffix Removal
    suffix_pattern = r'\s+(?:Inc\.?\s+ON|ON\s+Inc\.?|Inc\.?|ON)$'
    clean_name = re.sub(suffix_pattern, '', clean_name, flags=re.IGNORECASE).strip()
    
    return clean_name

def reorder_name(name):
    global NAME_CACHE
    
    # Normalize whitespace
    name = re.sub(r'\s+', ' ', name).strip()
    
    # 1. Check Excel Cache
    if name in ATHLETE_CORRECTIONS:
        return ATHLETE_CORRECTIONS[name]
    
    # 2. Check Runtime Cache
    if name in NAME_CACHE:
        return NAME_CACHE[name]
    
    parts = name.split()
    result = name
    
    if len(parts) == 1:
        result = name
    elif len(parts) == 2:
        result = f"{parts[1]} {parts[0]}"
    else:
        print(f"\n{Colors.YELLOW}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")
        print(f"{Colors.YELLOW}Multiple-word name detected: {Colors.BOLD}{name}{Colors.RESET}")
        print(f"{Colors.CYAN}This name is in 'Last First' format. Where does the LAST name end?{Colors.RESET}")
        
        for i in range(1, len(parts)):
            last_name = ' '.join(parts[:i])
            first_name = ' '.join(parts[i:])
            reordered = f"{first_name} {last_name}"
            print(f"{Colors.CYAN}{i}.{Colors.RESET} Last: {Colors.BOLD}{last_name}{Colors.RESET}, First: {Colors.BOLD}{first_name}{Colors.RESET} → {Colors.GREEN}{reordered}{Colors.RESET}")
        
        while True:
            choice = input(f"{Colors.CYAN}Enter choice (1-{len(parts)-1}): {Colors.RESET}").strip()
            if choice.isdigit() and 1 <= int(choice) < len(parts):
                split_point = int(choice)
                last_name = ' '.join(parts[:split_point])
                first_name = ' '.join(parts[split_point:])
                result = f"{first_name} {last_name}"
                
                # Save and Cache
                save_athlete_correction(name, result)
                print(f"{Colors.GREEN}✓ Saved as: {result}{Colors.RESET}")
                print(f"{Colors.YELLOW}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")
                break
            else:
                print(f"{Colors.RED}Invalid choice. Please enter a number between 1 and {len(parts)-1}.{Colors.RESET}")
    
    NAME_CACHE[name] = result
    return result

# ==========================================
#  SCRAPING FUNCTIONS
# ==========================================

def safe_find_text(soup, tag, default="Unknown"):
    element = soup.find(tag)
    return element.get_text().strip() if element else default

def parse_date(date_string):
    try:
        date_match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', date_string)
        if date_match:
            day, month, year = date_match.groups()
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        date_match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', date_string)
        if date_match: return date_match.group(0)
    except Exception: pass
    return datetime.now().strftime("%Y-%m-%d")

def fetch_url(url):
    time.sleep(0.5)
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        }
        debug_print(f"  Fetching: {url[:80]}...")
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or 'utf-8'
        debug_print(f"  Received {len(response.text)} characters")
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"  {Colors.RED}✗ Error fetching URL: {e}{Colors.RESET}")
        return None

def parse_row_data(row, headers):
    """Dynamically map row cells to headers."""
    cells = row.find_all('td')
    if not cells or len(cells) < 4: return None
    
    cell_values = [re.sub(r'\s+', ' ', c.get_text(strip=True)) for c in cells]
    row_dict = {}

    # 1. Handle Athlete/Club Column
    ath_idx = -1
    for i, h in enumerate(headers):
        if 'name' in h.lower() or 'gymnast' in h.lower(): ath_idx = i; break
    if ath_idx == -1 and len(cells) > 2: ath_idx = 2

    # 2. Handle YOB Column
    yob_idx = -1
    for i, h in enumerate(headers):
        if 'born' in h.lower() or 'yob' in h.lower(): yob_idx = i; break
    
    if yob_idx == -1 and len(cells) > 3:
        if cells[3].get_text(strip=True).isdigit() and len(cells[3].get_text(strip=True)) == 4:
            yob_idx = 3

    try:
        if ath_idx < len(cells):
            raw_html = str(cells[ath_idx])
            soup_cell = BeautifulSoup(raw_html, 'html.parser')
            link = soup_cell.find('a')
            if link:
                raw_name = link.get_text().strip()
            else:
                raw_name = BeautifulSoup(re.split(r'<br\s*/?>', raw_html)[0], 'html.parser').get_text().strip()
                
            row_dict['Name'] = reorder_name(raw_name)
            
            parts = re.split(r'<br\s*/?>', raw_html)
            raw_club = ""
            if len(parts) > 1:
                raw_club = BeautifulSoup(parts[1], 'html.parser').get_text().strip()
            row_dict['Club'] = standardize_club(raw_club)
    except Exception:
        row_dict['Name'] = "Unknown"
        row_dict['Club'] = "Unknown"

    if yob_idx != -1 and yob_idx < len(cells):
        row_dict['YOB'] = cell_values[yob_idx]

    # 3. Map remaining columns
    for i, (header, val) in enumerate(zip(headers, cell_values)):
        if i == ath_idx: continue
        if i == yob_idx: continue
        
        # Apply Rename Logic
        clean_header = header.strip()
        if clean_header in RENAME_MAP:
            clean_header = RENAME_MAP[clean_header]
            
        if not clean_header: clean_header = f"Col_{i}"
        
        row_dict[clean_header] = val

    return row_dict

# ==========================================
#  MAIN LOGIC
# ==========================================

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
    """Fetch and display list of competitions from ksis.eu."""
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
    all_competitions = []
    
    for link in soup.find_all('a', href=True):
        if 'id_prop=' in link['href']:
            pid = re.search(r'id_prop=(\d+)', link['href']).group(1)
            name = link.get_text().strip()
            is_live = False
            if link.parent:
                for badge in link.parent.find_all('span', class_='badge'):
                     if 'live' in badge.get_text().lower(): is_live = True

            if name and not any(c['prop_id'] == pid for c in all_competitions):
                all_competitions.append({'prop_id': pid, 'name': name, 'is_live': is_live})

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
    # Ensure corrections are loaded
    load_corrections()
    
    url = f"https://ksis.eu/resultx.php?id_prop={prop_id}"
    print(f"\n{Colors.CYAN}Fetching competition data (prop_id: {prop_id})...{Colors.RESET}")
    
    html_content = fetch_url(url)
    if not html_content:
        print(f"{Colors.RED}Failed to fetch competition page. Exiting...{Colors.RESET}")
        return
    
    soup = BeautifulSoup(html_content, "html.parser")
    comp_name = safe_find_text(soup, 'h3', f'Competition_{prop_id}')
    comp_name = re.sub(r'[\\/*?:"<>|]', "-", comp_name).strip()
    comp_date = parse_date(safe_find_text(soup, 'h4', ''))
    
    print(f'{Colors.CYAN}Parsing sessions...{Colors.RESET}')
    select_tag = soup.find('select', id='id_sut')
    if not select_tag:
        print(f"{Colors.RED}✗ Select tag not found.{Colors.RESET}")
        print(f"{Colors.YELLOW}ℹ The competition sessions may still be in progress or empty.{Colors.RESET}")
        return

    all_rows = []
    # Base headers ensures these exist even if table is empty
    all_headers = set(['Session', 'Name', 'YOB', 'Club', 'Score', 'Date']) 
    in_progress_count = 0
    
    options = select_tag.find_all('option')
    for option in options:
        value = option.get('value')
        session_name = option.get_text().strip()
        if value in ["0", ""]: continue
        
        result_url = (f"https://ksis.eu/load_result_total_ksismg_art.php?lang=en&id_prop={prop_id}&id_sut={value}&rn=null&mn=null&state=-1&age_group=&award=-1&nacinie=undefined")
        result_content = fetch_url(result_url)
        if not result_content:
            in_progress_count += 1
            print(f"{Colors.YELLOW} ✗ {session_name}: Could not fetch results{Colors.RESET}")
            continue
        
        table = BeautifulSoup(result_content, "html.parser").find('table', id='myTablePrihlasky')
        if not table:
            in_progress_count += 1
            print(f"{Colors.YELLOW} ⚠ {session_name}: No results table found (session may be in progress){Colors.RESET}")
            continue

        # Header Parsing
        headers = []
        thead = table.find('thead')
        if thead and thead.find_all('tr'):
             headers = [h.get_text(strip=True) for h in thead.find_all('tr')[-1].find_all(['th', 'td'])]
        if not headers:
            first = table.find('tr')
            if first: headers = [h.get_text(strip=True) for h in first.find_all(['th', 'td'])]

        for h in headers:
            clean = h.strip()
            if clean in RENAME_MAP: clean = RENAME_MAP[clean]
            if clean and clean not in DROP_COLS: all_headers.add(clean)

        data_rows = table.find('tbody').find_all('tr') if table.find('tbody') else table.find_all('tr')
        if not thead and data_rows and len(data_rows[0].find_all('td')) == len(headers):
             data_rows = data_rows[1:]
             
        row_count = 0
        for tr in data_rows:
            row_data = parse_row_data(tr, headers)
            if row_data:
                row_data['Session'] = session_name
                row_data['Date'] = comp_date
                all_rows.append(row_data)
                row_count += 1
        
        if row_count > 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} {session_name}: {row_count} athletes")
        else:
            in_progress_count += 1
            print(f"{Colors.YELLOW}  ⚠ {session_name}: No athletes found (may be in progress){Colors.RESET}")

    if all_rows:
        filename = f'{comp_name}-{datetime.now().strftime("%Y%m%d%H%M")}.csv'
        
        # Priority Order
        priority = ['Session', 'Name', 'YOB', 'Club', 'Score', 'Date']
        others = [h for h in all_headers if h not in priority and h not in DROP_COLS]
        final_headers = priority + sorted(others)
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=final_headers, extrasaction='ignore')
                writer.writeheader()
                writer.writerows(all_rows)
            print(f"\n{Colors.GREEN}{Colors.BOLD}✓ Successfully created {filename} with {len(all_rows)} records.{Colors.RESET}")
        except PermissionError:
             print(f"{Colors.RED}ERROR: Could not write to file. Is it open in Excel?{Colors.RESET}")
        
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
        print(f"{Colors.BOLD}{Colors.CYAN}║     KSIS Competition Results Tool     ║{Colors.RESET}")
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
            export_results(get_prop_id())
        
        elif choice == '5':
            print(f"\n{Colors.CYAN}Goodbye!{Colors.RESET}")
            break
        
        else:
            print(f"{Colors.RED}Invalid choice. Please enter 1-5.{Colors.RESET}")

def main():
    global DEBUG
    
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
    parser.add_argument('--debug', '-d', action='store_true', help='Enable debug output for troubleshooting')
    parser.add_argument('--prop-id', type=str, help='Competition prop_id to export (skips interactive menu)')
    parser.add_argument('--list', '-l', action='store_true', help='List available competitions and exit')
    
    args = parser.parse_args()
    DEBUG = args.debug
    
    if args.list:
        list_competitions()
    elif args.prop_id:
        if not args.prop_id.isdigit():
            print(f"{Colors.RED}Error: prop_id must be a number{Colors.RESET}")
            return
        export_results(args.prop_id)
    else:
        interactive_menu()

if __name__ == "__main__":
    main()
