import requests
import re
import csv
import argparse
import sys
import os
import time
from datetime import datetime
from bs4 import BeautifulSoup

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
DROP_COLS = ['E', 'Bonus', 'Comp', 'ND', 'D', 'SV'] 
RENAME_MAP = {
    'Total': 'Score', 
    'SV': 'Score',
    'Born': 'YOB', 
    'born': 'YOB'
}

def debug_print(message):
    """Print message only if DEBUG is enabled."""
    if DEBUG:
        print(f"{Colors.MAGENTA}[DEBUG] {message}{Colors.RESET}")

# ==========================================
#  DEPENDENCY CHECKING
# ==========================================

def check_pandas():
    """Check if pandas is available for Excel corrections."""
    try:
        import pandas
        import openpyxl
        return True
    except ImportError as e:
        print(f"{Colors.YELLOW}⚠ Optional dependency missing: {e}{Colors.RESET}")
        print(f"{Colors.YELLOW}  Excel corrections feature disabled.{Colors.RESET}")
        print(f"{Colors.CYAN}  To enable: pip install pandas openpyxl{Colors.RESET}")
        return False

# ==========================================
#  FILE LOADING / SAVING FUNCTIONS
# ==========================================

def load_excel_corrections(filename):
    """Generic helper to load corrections from Excel (Col A -> Col B)."""
    if not os.path.exists(filename):
        debug_print(f"Corrections file not found: {filename}")
        return {}
    
    if not check_pandas():
        return {}
    
    try:
        import pandas as pd
        df = pd.read_excel(filename, header=None, usecols=[0, 1])
        
        if df.empty:
            debug_print(f"Empty corrections file: {filename}")
            return {}
        
        df.columns = ['Original', 'Corrected']
        
        # Remove header row if present
        if str(df.iloc[0]['Original']).strip().lower() == 'original':
            df = df.iloc[1:]

        corrections = dict(zip(
            df['Original'].astype(str).str.strip(), 
            df['Corrected'].astype(str).str.strip()
        ))
        
        debug_print(f"Loaded {len(corrections)} corrections from {filename}")
        return corrections
        
    except Exception as e:
        print(f"{Colors.RED}⚠ Error loading {filename}: {e}{Colors.RESET}")
        return {}

def load_corrections():
    """Load both Club and Athlete corrections."""
    global CLUB_CORRECTIONS, ATHLETE_CORRECTIONS
    
    if CLUB_CORRECTIONS or ATHLETE_CORRECTIONS:
        return

    debug_print("Loading corrections dictionaries...")
    
    CLUB_CORRECTIONS = load_excel_corrections("Club Name Corrections.xlsx")
    if CLUB_CORRECTIONS:
        print(f"{Colors.GREEN}✓ Loaded {len(CLUB_CORRECTIONS)} club corrections{Colors.RESET}")
        
    ATHLETE_CORRECTIONS = load_excel_corrections("Athlete Name Corrections.xlsx")
    if ATHLETE_CORRECTIONS:
        print(f"{Colors.GREEN}✓ Loaded {len(ATHLETE_CORRECTIONS)} athlete name corrections{Colors.RESET}")

def save_athlete_correction(original, corrected):
    """Appends a new athlete correction to the Excel file."""
    if not check_pandas():
        print(f"{Colors.YELLOW}⚠ Cannot save correction - pandas not available{Colors.RESET}")
        return
        
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
        
        ATHLETE_CORRECTIONS[original] = corrected
        
    except Exception as e:
        print(f"{Colors.RED}⚠ Error saving correction: {e}{Colors.RESET}")

# ==========================================
#  STANDARDIZATION LOGIC
# ==========================================

def standardize_club(club_name):
    """Standardize club names using corrections and suffix removal."""
    if not club_name:
        return ""
    
    clean_name = club_name.strip()
    
    # Apply dictionary corrections
    if clean_name in CLUB_CORRECTIONS:
        clean_name = CLUB_CORRECTIONS[clean_name].strip()
        debug_print(f"Club corrected: {club_name} → {clean_name}")
    
    # Remove common suffixes
    suffix_pattern = r'\s+(?:Inc\.?\s+ON|ON\s+Inc\.?|Inc\.?|ON)$'
    clean_name = re.sub(suffix_pattern, '', clean_name, flags=re.IGNORECASE).strip()
    
    return clean_name

def reorder_name(name):
    """Reorder name from 'Last First' to 'First Last' format with caching."""
    global NAME_CACHE
    
    name = re.sub(r'\s+', ' ', name).strip()
    
    # Check Excel corrections first
    if name in ATHLETE_CORRECTIONS:
        debug_print(f"Name from Excel: {name} → {ATHLETE_CORRECTIONS[name]}")
        return ATHLETE_CORRECTIONS[name]
    
    # Check runtime cache
    if name in NAME_CACHE:
        return NAME_CACHE[name]
    
    parts = name.split()
    result = name
    
    if len(parts) == 1:
        result = name
    elif len(parts) == 2:
        result = f"{parts[1]} {parts[0]}"
        debug_print(f"Name reordered: {name} → {result}")
    else:
        # Multi-word names require user input
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
    """Safely extract text from a tag."""
    element = soup.find(tag)
    return element.get_text().strip() if element else default

def parse_date(date_string):
    """Parse date string into YYYY-MM-DD format."""
    if not date_string:
        return None
    
    try:
        # Match DD.MM.YYYY
        date_match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', date_string)
        if date_match:
            day, month, year = date_match.groups()
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        
        # Match YYYY-MM-DD
        date_match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', date_string)
        if date_match:
            return date_match.group(0)
            
    except Exception as e:
        debug_print(f"Date parsing error: {e}")
    
    return None

def fetch_url(url):
    """Fetch URL with rate limiting and error handling."""
    time.sleep(0.5)  # Rate limiting
    
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        }
        
        debug_print(f"Fetching: {url[:80]}...")
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or 'utf-8'
        
        debug_print(f"Received {len(response.text)} characters")
        return response.text
        
    except requests.exceptions.RequestException as e:
        print(f"{Colors.RED}✗ Error fetching URL: {e}{Colors.RESET}")
        return None

def parse_row_data(row, headers):
    """Dynamically map row cells to headers."""
    cells = row.find_all('td')
    if not cells or len(cells) < 4:
        return None
    
    cell_values = [re.sub(r'\s+', ' ', c.get_text(strip=True)) for c in cells]
    row_dict = {}

    # Find athlete/club column
    ath_idx = -1
    for i, h in enumerate(headers):
        if h and ('name' in h.lower() or 'gymnast' in h.lower()):
            ath_idx = i
            break
    if ath_idx == -1 and len(cells) > 2:
        ath_idx = 2  # Fallback to column 2

    # Find YOB column
    yob_idx = -1
    for i, h in enumerate(headers):
        if h and ('born' in h.lower() or 'yob' in h.lower()):
            yob_idx = i
            break
    
    if yob_idx == -1 and len(cells) > 3:
        if cells[3].get_text(strip=True).isdigit() and len(cells[3].get_text(strip=True)) == 4:
            yob_idx = 3

    # Parse athlete and club
    try:
        if ath_idx >= 0 and ath_idx < len(cells):
            raw_html = str(cells[ath_idx])
            soup_cell = BeautifulSoup(raw_html, 'html.parser')
            
            # Extract name
            link = soup_cell.find('a')
            if link:
                raw_name = link.get_text().strip()
            else:
                parts = re.split(r'<br\s*/?>', raw_html)
                raw_name = BeautifulSoup(parts[0], 'html.parser').get_text().strip()
            
            row_dict['Name'] = reorder_name(raw_name)
            
            # Extract club
            parts = re.split(r'<br\s*/?>', raw_html)
            raw_club = ""
            if len(parts) > 1:
                raw_club = BeautifulSoup(parts[1], 'html.parser').get_text().strip()
            row_dict['Club'] = standardize_club(raw_club)
            
    except Exception as e:
        debug_print(f"Error parsing athlete/club: {e}")
        row_dict['Name'] = "Unknown"
        row_dict['Club'] = "Unknown"

    # Parse YOB
    if yob_idx >= 0 and yob_idx < len(cell_values):
        row_dict['YOB'] = cell_values[yob_idx]

    # Map remaining columns
    for i, (header, val) in enumerate(zip(headers, cell_values)):
        if i == ath_idx or i == yob_idx:
            continue
        
        clean_header = header.strip() if header else f"Col_{i}"
        
        # Apply rename logic
        if clean_header in RENAME_MAP:
            clean_header = RENAME_MAP[clean_header]
        
        if clean_header and clean_header not in DROP_COLS:
            row_dict[clean_header] = val

    return row_dict

# ==========================================
#  CORE DATA FETCHING
# ==========================================

def fetch_competition_data(prop_id):
    """
    Fetch all results for a competition.
    Returns: (comp_name, rows, headers, stats_dict)
    """
    load_corrections()
    
    url = f"https://ksis.eu/resultx.php?id_prop={prop_id}"
    print(f"\n{Colors.CYAN}Fetching competition data (prop_id: {prop_id})...{Colors.RESET}")
    
    html_content = fetch_url(url)
    if not html_content:
        return None, [], set(), {'in_progress': 0, 'completed': 0, 'failed': 0}
    
    soup = BeautifulSoup(html_content, "html.parser")
    comp_name = safe_find_text(soup, 'h3', f'Competition_{prop_id}')
    comp_name = re.sub(r'[\\/*?:"<>|]', "-", comp_name).strip()
    
    raw_date = safe_find_text(soup, 'h4', '')
    comp_date = parse_date(raw_date) or datetime.now().strftime("%Y-%m-%d")
    
    print(f'{Colors.CYAN}Parsing sessions for "{comp_name}"...{Colors.RESET}')
    select_tag = soup.find('select', id='id_sut')
    
    if not select_tag:
        print(f"{Colors.RED}✗ Select tag not found.{Colors.RESET}")
        lower_content = html_content.lower()
        if 'in progress' in lower_content or 'live' in lower_content or comp_name != f'Competition_{prop_id}':
            print(f"{Colors.YELLOW}ℹ The competition sessions may still be in progress.{Colors.RESET}")
        return comp_name, [], set(), {'in_progress': 0, 'completed': 0, 'failed': 0}

    all_rows = []
    all_headers = set(['Competition', 'Session', 'Name', 'YOB', 'Club', 'Score', 'Date'])
    stats = {'in_progress': 0, 'completed': 0, 'failed': 0}
    
    options = select_tag.find_all('option')
    
    for option in options:
        value = option.get('value')
        session_name = option.get_text().strip()
        
        if value in ["0", ""]:
            continue
        
        result_url = (
            f"https://ksis.eu/load_result_total_ksismg_art.php?"
            f"lang=en&id_prop={prop_id}&id_sut={value}&rn=null&mn=null"
            f"&state=-1&age_group=&award=-1&nacinie=undefined"
        )
        
        result_content = fetch_url(result_url)
        if not result_content:
            stats['failed'] += 1
            print(f"{Colors.YELLOW}  ✗ {session_name}: Could not fetch results{Colors.RESET}")
            continue
        
        table = BeautifulSoup(result_content, "html.parser").find('table', id='myTablePrihlasky')
        if not table:
            stats['in_progress'] += 1
            print(f"{Colors.YELLOW}  ⚠ {session_name}: No results table found (session may be in progress){Colors.RESET}")
            continue

        # Parse headers
        headers = []
        thead = table.find('thead')
        if thead and thead.find_all('tr'):
            headers = [h.get_text(strip=True) for h in thead.find_all('tr')[-1].find_all(['th', 'td'])]
        if not headers:
            first_row = table.find('tr')
            if first_row:
                headers = [h.get_text(strip=True) for h in first_row.find_all(['th', 'td'])]

        # Add headers to master set
        for h in headers:
            clean = h.strip() if h else ""
            if clean in RENAME_MAP:
                clean = RENAME_MAP[clean]
            if clean and clean not in DROP_COLS:
                all_headers.add(clean)

        # Parse data rows
        data_rows = table.find('tbody').find_all('tr') if table.find('tbody') else table.find_all('tr')
        
        # Skip header row if no thead
        if not thead and data_rows and len(data_rows) > 0:
            first_cells = data_rows[0].find_all('td')
            if len(first_cells) == len(headers):
                data_rows = data_rows[1:]
        
        row_count = 0
        for tr in data_rows:
            row_data = parse_row_data(tr, headers)
            if row_data:
                row_data['Competition'] = comp_name
                row_data['Session'] = session_name
                row_data['Date'] = comp_date
                all_rows.append(row_data)
                row_count += 1
        
        if row_count > 0:
            stats['completed'] += 1
            print(f"  {Colors.GREEN}✓{Colors.RESET} {session_name}: {row_count} athletes")
        else:
            stats['in_progress'] += 1
            if 'live' in session_name.lower():
                print(f"{Colors.YELLOW}  ⚠ {session_name}: Session is in progress{Colors.RESET}")
            else:
                print(f"{Colors.YELLOW}  ⚠ {session_name}: No athletes found (may be in progress){Colors.RESET}")

    return comp_name, all_rows, all_headers, stats

# ==========================================
#  EXPORT LOGIC
# ==========================================

def write_csv(rows, headers, filename):
    """Write rows to CSV with smart column ordering."""
    if not rows:
        print(f"{Colors.YELLOW}No rows to write.{Colors.RESET}")
        return False

    # Priority ordering
    priority = ['Competition', 'Session', 'Name', 'YOB', 'Club', 'Score', 'Date']
    others = sorted([h for h in headers if h not in priority and h not in DROP_COLS])
    final_headers = priority + others
    
    try:
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=final_headers, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(rows)
        return True
        
    except PermissionError:
        print(f"{Colors.RED}✗ ERROR: Could not write to '{filename}'. Is it open in Excel?{Colors.RESET}")
        return False
    except Exception as e:
        print(f"{Colors.RED}✗ ERROR writing file: {e}{Colors.RESET}")
        return False

def export_results_single(prop_id):
    """Export single competition or comma-separated list."""
    # Handle comma-separated list
    if ',' in prop_id:
        ids = [x.strip() for x in prop_id.split(',') if x.strip().isdigit()]
        if ids:
            timestamp = datetime.now().strftime('%Y%m%d%H%M')
            filename = f"Aggregated_Manual_List_{len(ids)}_Comps-{timestamp}.csv"
            export_results_aggregated(ids, filename)
        else:
            print(f"{Colors.RED}No valid prop_ids in list.{Colors.RESET}")
        return
    
    # Single competition
    name, rows, headers, stats = fetch_competition_data(prop_id)
    
    if rows:
        timestamp = datetime.now().strftime('%Y%m%d%H%M')
        filename = f'{name}-{timestamp}.csv'
        
        if write_csv(rows, headers, filename):
            print(f"\n{Colors.GREEN}{Colors.BOLD}✓ Successfully created {filename} with {len(rows)} records.{Colors.RESET}")
            
            if stats['in_progress'] > 0:
                print(f"{Colors.MAGENTA}{Colors.BOLD}ℹ {stats['in_progress']} session(s) still in progress{Colors.RESET}")
    else:
        print(f"\n{Colors.YELLOW}No data collected for prop_id {prop_id}.{Colors.RESET}")
        if stats['in_progress'] > 0:
            print(f"{Colors.MAGENTA}{Colors.BOLD}ℹ {stats['in_progress']} session(s) still in progress{Colors.RESET}")

def export_results_aggregated(prop_ids, output_filename):
    """Fetch multiple competitions and combine into one file."""
    master_rows = []
    master_headers = set(['Competition', 'Session', 'Name', 'YOB', 'Club', 'Score', 'Date'])
    total_stats = {'in_progress': 0, 'completed': 0, 'failed': 0}
    
    print(f"\n{Colors.BOLD}{Colors.CYAN}Starting Aggregated Export for {len(prop_ids)} competitions...{Colors.RESET}")
    
    for pid in prop_ids:
        _, rows, headers, stats = fetch_competition_data(pid)
        if rows:
            master_rows.extend(rows)
            master_headers.update(headers)
        
        # Accumulate stats
        for key in total_stats:
            total_stats[key] += stats[key]
    
    if master_rows:
        if write_csv(master_rows, master_headers, output_filename):
            print(f"\n{Colors.GREEN}{Colors.BOLD}✓ Successfully created {output_filename} with {len(master_rows)} records.{Colors.RESET}")
            print(f"{Colors.CYAN}  Sessions completed: {total_stats['completed']}{Colors.RESET}")
            
            if total_stats['in_progress'] > 0:
                print(f"{Colors.MAGENTA}{Colors.BOLD}ℹ {total_stats['in_progress']} session(s) still in progress{Colors.RESET}")
    else:
        print(f"\n{Colors.YELLOW}No data collected from any competition.{Colors.RESET}")

# ==========================================
#  COMPETITION SEARCH & LISTING
# ==========================================

def get_competitions_from_menu():
    """Scrape menu and return list of competitions with metadata."""
    url = "https://ksis.eu/menu.php?akcia=S&oblast=ARTW&country=CAN"
    print(f"{Colors.CYAN}Fetching competition list...{Colors.RESET}")
    
    html = fetch_url(url)
    if not html:
        return []
    
    soup = BeautifulSoup(html, "html.parser")
    comps = []
    missing_dates = 0
    
    for link in soup.find_all('a', href=True):
        if 'id_prop=' not in link['href']:
            continue
            
        match = re.search(r'id_prop=(\d+)', link['href'])
        if not match:
            continue
            
        pid = match.group(1)
        name = link.get_text().strip()
        
        if not name:
            continue
        
        # Parse date from table row
        date_str = None
        row = link.find_parent('tr')
        if row:
            cells = row.find_all('td')
            if cells:
                date_text = cells[0].get_text(strip=True)
                date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_text)
                if date_match:
                    d, m, y = date_match.groups()
                    date_str = f"{y}-{m}-{d}"
        
        if not date_str:
            missing_dates += 1
            debug_print(f"No date found for {pid}: {name[:40]}")
        
        # Check for LIVE badge
        is_live = False
        if link.parent:
            for badge in link.parent.find_all('span', class_='badge'):
                if 'live' in badge.get_text().lower():
                    is_live = True
                    break
        
        # Avoid duplicates
        if not any(c['id'] == pid for c in comps):
            comps.append({
                'id': pid,
                'name': name,
                'date': date_str,
                'is_live': is_live
            })
    
    if missing_dates > 0:
        print(f"{Colors.YELLOW}⚠ {missing_dates} competition(s) have no date information{Colors.RESET}")
    
    return comps

def list_competitions(live_only=False, search_keyword=None):
    """Display competition list with optional filters."""
    competitions = get_competitions_from_menu()
    
    if not competitions:
        print(f"{Colors.YELLOW}No competitions found.{Colors.RESET}")
        return []

    # Apply filters
    if live_only:
        competitions = [c for c in competitions if c['is_live']]
    
    if search_keyword:
        k = search_keyword.lower()
        competitions = [c for c in competitions if k in c['name'].lower()]

    if not competitions:
        if live_only:
            print(f"{Colors.YELLOW}No live competitions found.{Colors.RESET}")
        elif search_keyword:
            print(f"{Colors.YELLOW}No competitions matching '{search_keyword}'.{Colors.RESET}")
        return []

    # Display
    print(f"\n{Colors.BOLD}Available Competitions:{Colors.RESET}")
    print(f"{Colors.CYAN}{'ID':<8} {'Date':<12} {'Competition Name'}{Colors.RESET}")
    print("-" * 80)
    
    for c in competitions:
        live_badge = f" {Colors.RED}[LIVE]{Colors.RESET}" if c['is_live'] else ""
        date_display = c['date'] if c['date'] else "Unknown"
        print(f"{Colors.GREEN}{c['id']:<8}{Colors.RESET} {date_display:<12} {c['name']}{live_badge}")
    
    print(f"\n{Colors.BOLD}Total: {len(competitions)} competitions{Colors.RESET}")
    return competitions

def search_by_date_range():
    """Search competitions by date range and offer aggregated export."""
    print(f"\n{Colors.CYAN}{Colors.BOLD}--- Date Range Search ---{Colors.RESET}")
    
    start_str = input(f"{Colors.CYAN}Enter Start Date (YYYY-MM-DD): {Colors.RESET}").strip()
    end_str = input(f"{Colors.CYAN}Enter End Date   (YYYY-MM-DD): {Colors.RESET}").strip()
    
    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_str, "%Y-%m-%d")
    except ValueError:
        print(f"{Colors.RED}Invalid date format. Please use YYYY-MM-DD.{Colors.RESET}")
        return

    all_comps = get_competitions_from_menu()
    
    matches = []
    skipped = 0
    
    for c in all_comps:
        if c['date']:
            try:
                c_date = datetime.strptime(c['date'], "%Y-%m-%d")
                if start_date <= c_date <= end_date:
                    matches.append(c)
            except ValueError:
                debug_print(f"Invalid date format for {c['id']}: {c['date']}")
                skipped += 1
        else:
            skipped += 1

    if skipped > 0:
        print(f"{Colors.YELLOW}⚠ Skipped {skipped} competition(s) with missing/invalid dates{Colors.RESET}")

    if not matches:
        print(f"{Colors.YELLOW}No competitions found in range {start_str} to {end_str}.{Colors.RESET}")
        return

    print(f"\n{Colors.GREEN}Found {len(matches)} competition(s) in range:{Colors.RESET}")
    for m in matches:
        live_badge = f" {Colors.RED}[LIVE]{Colors.RESET}" if m['is_live'] else ""
        print(f"  {Colors.CYAN}{m['date']}{Colors.RESET}: {m['name']} (ID: {m['id']}){live_badge}")
    
    confirm = input(f"\n{Colors.CYAN}Export all {len(matches)} competitions to one file? (y/n): {Colors.RESET}").lower()
    
    if confirm == 'y':
        ids = [m['id'] for m in matches]
        filename = f"Aggregated_Results_{start_str}_to_{end_str}.csv"
        export_results_aggregated(ids, filename)
    else:
        print(f"{Colors.YELLOW}Export cancelled.{Colors.RESET}")

# ==========================================
#  INTERACTIVE MENU
# ==========================================

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
        print(f"{Colors.CYAN}4.{Colors.RESET} Export results by prop_id (single or comma-separated list)")
        print(f"{Colors.CYAN}5.{Colors.RESET} Search & Export by Date Range")
        print(f"{Colors.CYAN}6.{Colors.RESET} Exit")
        
        choice = input(f"\n{Colors.CYAN}Enter your choice (1-6): {Colors.RESET}").strip()
        
        if choice == '1':
            list_competitions()
            
        elif choice == '2':
            list_competitions(live_only=True)
            
        elif choice == '3':
            keyword = input(f"{Colors.CYAN}Enter search keyword: {Colors.RESET}").strip()
            if keyword:
                list_competitions(search_keyword=keyword)
            else:
                print(f"{Colors.RED}No keyword provided.{Colors.RESET}")
                
        elif choice == '4':
            pid = input(f"{Colors.CYAN}Enter prop_id (or comma-separated list like 8819,8820): {Colors.RESET}").strip()
            if pid:
                export_results_single(pid)
            else:
                print(f"{Colors.RED}No prop_id provided.{Colors.RESET}")
                
        elif choice == '5':
            search_by_date_range()
            
        elif choice == '6':
            print(f"\n{Colors.CYAN}Goodbye!{Colors.RESET}")
            break
            
        else:
            print(f"{Colors.RED}Invalid choice. Please enter 1-6.{Colors.RESET}")

# ==========================================
#  MAIN ENTRY POINT
# ==========================================

def main():
    global DEBUG
    
    parser = argparse.ArgumentParser(
        description='Parse KSIS competition results and export to CSV',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python ksis_export.py                    # Interactive menu mode
  python ksis_export.py --list             # List competitions and exit
  python ksis_export.py --prop-id 8819     # Export specific competition
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
        list_competitions()
    elif args.prop_id:
        export_results_single(args.prop_id)
    else:
        interactive_menu()

if __name__ == "__main__":
    main()
