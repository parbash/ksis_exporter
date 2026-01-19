# KSIS Competition Results Tool - User Guide

## Table of Contents
1. [Installation](#installation)
2. [Quick Start](#quick-start)
3. [Features Overview](#features-overview)
4. [Interactive Menu Guide](#interactive-menu-guide)
5. [Command Line Usage](#command-line-usage)
6. [Excel Corrections](#excel-corrections)
7. [Troubleshooting](#troubleshooting)

---

## Installation

### Step 1: Install Python

#### Windows
1. Download Python from [python.org/downloads](https://www.python.org/downloads/)
2. Run the installer
3. **IMPORTANT**: Check "Add Python to PATH" during installation
4. Click "Install Now"
5. Verify installation by opening Command Prompt and typing:
   ```
   python --version
   ```

#### macOS
1. Download Python from [python.org/downloads](https://www.python.org/downloads/)
2. Run the installer package
3. Follow the installation wizard
4. Verify installation by opening Terminal and typing:
   ```
   python3 --version
   ```

#### Linux
Most Linux distributions come with Python pre-installed. If not:
```bash
sudo apt-get update
sudo apt-get install python3 python3-pip
```

### Step 2: Install Required Packages

Open your terminal/command prompt and navigate to the folder containing `ksis_export.py`, then run:

```bash
pip install requests beautifulsoup4
```

**For Excel corrections feature** (optional but recommended):
```bash
pip install pandas openpyxl
```

If you're on macOS/Linux and `pip` doesn't work, try:
```bash
pip3 install requests beautifulsoup4 pandas openpyxl
```

### Step 3: Verify Installation

Run the script to test:
```bash
python ksis_export.py
```

You should see the interactive menu appear.

---

## Quick Start

### Basic Usage (Interactive Mode)

1. Open your terminal/command prompt
2. Navigate to the script folder:
   ```bash
   cd path/to/script/folder
   ```
3. Run the script:
   ```bash
   python ksis_export.py
   ```
4. Choose option **1** to see all available competitions
5. Note the prop_id of the competition you want
6. Choose option **4** and enter the prop_id
7. The CSV file will be created in the same folder

### Example Session
```
$ python ksis_export.py

╔═══════════════════════════════════════╗
║     KSIS Competition Results Tool     ║
╚═══════════════════════════════════════╝

Select an option:
1. List all competitions
2. List live competitions only
3. Search competitions by keyword
4. Export results by prop_id (single or comma-separated list)
5. Search & Export by Date Range
6. Exit

Enter your choice (1-6): 1

[Competition list appears...]

Enter your choice (1-6): 4
Enter prop_id (or comma-separated list like 8819,8820): 8819

[Export begins...]
✓ Successfully created Competition_Name-202501181435.csv with 156 records.
```

---

## Features Overview

### 1. List All Competitions
- Shows all Canadian Women's Artistic Gymnastics competitions
- Displays prop_id, date, and competition name
- Indicates LIVE competitions with a red [LIVE] badge

### 2. List Live Competitions Only
- Filters to show only competitions currently in progress
- Useful for tracking ongoing events

### 3. Search by Keyword
- Search competition names by keyword
- Case-insensitive search
- Example: Search "Provincial" to find all provincial championships

### 4. Export Results
- **Single competition**: Enter one prop_id (e.g., `8819`)
- **Multiple competitions**: Enter comma-separated list (e.g., `8819,8820,8821`)
- Creates timestamped CSV file(s)
- Automatically handles name reordering and club standardization

### 5. Search & Export by Date Range
- Find all competitions within a date range
- Option to export all matches to a single aggregated file
- Perfect for season summaries or regional analysis

### 6. Name Reordering
- Automatically converts "Last First" to "First Last" format
- Simple names (2 words) are handled automatically
- Multi-word names prompt you for confirmation:
  ```
  Multiple-word name detected: Rohas De Suza Mary Elizabeth
  This name is in 'Last First' format. Where does the LAST name end?
  1. Last: Rohas, First: De Suza Mary Elizabeth → De Suza Mary Elizabeth Rohas
  2. Last: Rohas De, First: Suza Mary Elizabeth → Suza Mary Elizabeth Rohas De
  3. Last: Rohas De Suza, First: Mary Elizabeth → Mary Elizabeth Rohas De Suza
  4. Last: Rohas De Suza Mary, First: Elizabeth → Elizabeth Rohas De Suza Mary
  Enter choice (1-4): 3
  ```
- Your choices are saved to avoid repeated prompts

---

## Interactive Menu Guide

### Menu Option 1: List All Competitions
```
Enter your choice (1-6): 1

Available Competitions:
ID       Date         Competition Name
--------------------------------------------------------------------------------
9045     2026-01-18   Ontario Winter Games 2026
9012     2026-01-11   Provincial Championships [LIVE]
8819     2025-12-15   Ottawa Invitational

Total: 3 competitions
```

### Menu Option 2: List Live Competitions
Shows only competitions currently in progress with active sessions.

### Menu Option 3: Search by Keyword
```
Enter your choice (1-6): 3
Enter search keyword: ontario

Available Competitions:
ID       Date         Competition Name
--------------------------------------------------------------------------------
9045     2026-01-18   Ontario Winter Games 2026
8995     2025-11-22   Ontario Cup Finals

Total: 2 competitions
```

### Menu Option 4: Export Results

**Single Competition:**
```
Enter your choice (1-6): 4
Enter prop_id (or comma-separated list like 8819,8820): 8819

Fetching competition data (prop_id: 8819)...
Parsing sessions for "Ottawa Invitational"...
  ✓ Level 4 (Age 10) WAG: 24 athletes
  ✓ Level 5 (Age 11-12) WAG: 18 athletes
  ⚠ Level 6 (Age 13) WAG - LIVE: Session is in progress

✓ Successfully created Ottawa Invitational-202601181435.csv with 42 records.
ℹ 1 session(s) still in progress
```

**Multiple Competitions:**
```
Enter prop_id (or comma-separated list like 8819,8820): 8819,8820,8821

Starting Aggregated Export for 3 competitions...
[Processing each competition...]

✓ Successfully created Aggregated_Manual_List_3_Comps-202601181435.csv with 156 records.
  Sessions completed: 12
ℹ 2 session(s) still in progress
```

### Menu Option 5: Date Range Export
```
Enter your choice (1-6): 5

--- Date Range Search ---
Enter Start Date (YYYY-MM-DD): 2025-12-01
Enter End Date   (YYYY-MM-DD): 2025-12-31

Found 5 competition(s) in range:
  2025-12-05: Winter Cup Qualifier (ID: 8801)
  2025-12-12: Holiday Classic (ID: 8815)
  2025-12-15: Ottawa Invitational (ID: 8819)
  2025-12-20: Christmas Challenge (ID: 8825)
  2025-12-28: Year End Championships (ID: 8830)

Export all 5 competitions to one file? (y/n): y

[Aggregated export begins...]
✓ Successfully created Aggregated_Results_2025-12-01_to_2025-12-31.csv
```

---

## Command Line Usage

### Basic Syntax
```bash
python ksis_export.py [OPTIONS]
```

### Available Options

| Option | Short | Description |
|--------|-------|-------------|
| `--list` | `-l` | List all competitions and exit |
| `--prop-id ID` | | Export specific competition (skip menu) |
| `--debug` | `-d` | Enable detailed debug output |

### Examples

**List competitions:**
```bash
python ksis_export.py --list
```

**Export single competition:**
```bash
python ksis_export.py --prop-id 8819
```

**Export with debug output:**
```bash
python ksis_export.py --debug --prop-id 8819
```

**Export multiple competitions:**
```bash
python ksis_export.py --prop-id 8819,8820,8821
```

### Advanced Usage

**Batch processing (Windows):**
Create a file `export_all.bat`:
```batch
@echo off
python ksis_export.py --prop-id 8819
python ksis_export.py --prop-id 8820
python ksis_export.py --prop-id 8821
pause
```

**Batch processing (macOS/Linux):**
Create a file `export_all.sh`:
```bash
#!/bin/bash
python3 ksis_export.py --prop-id 8819
python3 ksis_export.py --prop-id 8820
python3 ksis_export.py --prop-id 8821
```
Make it executable: `chmod +x export_all.sh`

---

## Excel Corrections

### Overview
The script supports Excel-based correction files to standardize athlete names and club names. This is optional but highly recommended for consistent data.

### Setting Up Corrections

#### Club Name Corrections
Create a file named `Club Name Corrections.xlsx` in the same folder as the script:

| Original | Corrected |
|----------|-----------|
| Club de gymnastique Les Sittelles ON | Les Sittelles |
| Ottawa Gymnastics Centre | OGC |
| Rideau Gymnastics ON Inc. | Rideau Gymnastics |

**File Format:**
- Column A: Original name (exactly as it appears in KSIS)
- Column B: Corrected/standardized name
- First row can optionally be headers ("Original", "Corrected")

#### Athlete Name Corrections
Create a file named `Athlete Name Corrections.xlsx` in the same folder:

| Original | Corrected |
|----------|-----------|
| Rohas De Suza Mary Elizabeth | Mary Elizabeth Rohas De Suza |
| Smith Jones Anna | Anna Smith Jones |

**How It Works:**
1. When the script encounters a multi-word name, it prompts you
2. You make a choice for how to split first/last name
3. Your choice is automatically saved to `Athlete Name Corrections.xlsx`
4. Next time you see that name, it's handled automatically
5. The corrections file builds up over time

### Benefits
- **Consistency**: Same club/athlete names across all exports
- **Efficiency**: No repeated prompts for known names
- **Sharing**: Share correction files with teammates
- **Flexibility**: Edit Excel files manually if needed

### Example: Building Your Corrections Database
First run (new athlete):
```
Multiple-word name detected: Rohas De Suza Mary Elizabeth
[You choose option 3]
✓ Saved correction to 'Athlete Name Corrections.xlsx'
```

Second run (same athlete):
```
[Name is automatically corrected, no prompt]
```

---

## Troubleshooting

### Common Issues

#### "Module not found" Error
**Problem:** Script says `ModuleNotFoundError: No module named 'requests'`

**Solution:**
```bash
pip install requests beautifulsoup4
```

If that doesn't work, try:
```bash
python -m pip install requests beautifulsoup4
```

---

#### "Python is not recognized" Error (Windows)
**Problem:** Command prompt doesn't recognize `python` command

**Solution:**
1. Reinstall Python and check "Add Python to PATH"
2. Or use full path: `C:\Python312\python.exe ksis_export.py`

---

#### No Colors Showing (Windows)
**Problem:** Terminal shows garbled text or no colors

**Solution:**
- Use Windows Terminal (recommended) or PowerShell
- Command Prompt on Windows 10+ should work
- Colors work best in Windows Terminal (free from Microsoft Store)

---

#### "Permission Denied" When Writing CSV
**Problem:** `ERROR: Could not write to file. Is it open in Excel?`

**Solution:**
- Close the CSV file in Excel before running the script
- If file is open, script cannot overwrite it
- Save your work in Excel, close it, then re-run the script

---

#### Empty or Missing Data
**Problem:** Competition shows "No athletes found"

**Possible Causes:**
1. **Session in progress**: Wait for results to be posted
2. **Wrong prop_id**: Verify the ID from the competition list
3. **Website structure changed**: Run with `--debug` flag and report issue

**Debug Mode:**
```bash
python ksis_export.py --debug --prop-id 8819
```
This shows detailed information about what the script is doing.

---

#### Date Not Found for Competitions
**Problem:** Some competitions show "Unknown" for date

**Explanation:**
- Some competitions on KSIS don't have dates in the standard location
- The script will still export these competitions
- Date field will show "Unknown" or current date

---

#### Pandas Not Available Warning
**Problem:** `⚠ Optional dependency missing: No module named 'pandas'`

**Impact:**
- Excel corrections feature won't work
- Basic functionality still works fine
- Name reordering still prompts you, but doesn't save to Excel

**Solution (optional):**
```bash
pip install pandas openpyxl
```

---

### Getting Help

If you encounter issues not covered here:

1. **Run with debug mode:**
   ```bash
   python ksis_export.py --debug --prop-id 8819
   ```

2. **Check the console output** for error messages

3. **Verify files:**
   - Script file: `ksis_export.py`
   - Output location: Same folder as script
   - Permissions: Can write to folder

4. **Test basic functionality:**
   ```bash
   python ksis_export.py --list
   ```
   If this works, Python and modules are installed correctly.

---

## Output Format

### CSV File Structure

The exported CSV files contain the following columns:

| Column | Description |
|--------|-------------|
| Competition | Competition name |
| Session | Session name (e.g., "Level 4 Age 10 WAG") |
| Name | Athlete name (First Last format) |
| YOB | Year of birth |
| Club | Club name (standardized) |
| Score | Final score |
| Date | Competition date (YYYY-MM-DD) |
| [Additional] | Other columns from results (apparatus scores, etc.) |

### File Naming Convention

**Single Competition:**
```
Competition_Name-YYYYMMDDHHMM.csv
```
Example: `Ottawa Invitational-202601181435.csv`

**Aggregated (Manual List):**
```
Aggregated_Manual_List_N_Comps-YYYYMMDDHHMM.csv
```
Example: `Aggregated_Manual_List_3_Comps-202601181435.csv`

**Aggregated (Date Range):**
```
Aggregated_Results_YYYY-MM-DD_to_YYYY-MM-DD.csv
```
Example: `Aggregated_Results_2025-12-01_to_2025-12-31.csv`

### Excel Compatibility

Files are saved with UTF-8 BOM encoding for perfect Excel compatibility:
- Special characters display correctly (é, à, ñ, etc.)
- No encoding issues when opening in Excel
- Works on Windows, macOS, and Linux versions of Excel

---

## Tips and Best Practices

### 1. Regular Backups
Keep your correction files safe:
- `Club Name Corrections.xlsx`
- `Athlete Name Corrections.xlsx`

These build up valuable standardization data over time.

### 2. Batch Processing
For multiple competitions, use comma-separated lists:
```
Enter prop_id: 8819,8820,8821,8825
```
This creates one aggregated file instead of many separate files.

### 3. Monitor In-Progress Sessions
Look for the magenta message:
```
ℹ 3 session(s) still in progress
```
Re-run the export later to get complete results.

### 4. Use Date Range Search
For end-of-season reports or regional summaries:
- Menu Option 5
- Enter date range
- Export all matches to one file

### 5. Verify Competition Names
Before exporting, use Option 1 to:
- Verify the prop_id is correct
- Check if competition is LIVE
- Confirm the competition date

---

## Appendix: Understanding prop_id

**What is a prop_id?**
- Unique identifier for each competition in the KSIS system
- Usually a 4-5 digit number (e.g., 8819, 9045)
- Remains constant for a competition
- Found in the competition URL or from the script's list feature

**How to find prop_id:**
1. Use the script: `python ksis_export.py --list`
2. Or check the KSIS website URL:
   `https://ksis.eu/resultx.php?id_prop=8819`
   The prop_id is `8819`

---

## Version Information

**Script Version:** 2.0  
**Last Updated:** January 2026  
**Compatible with:** Python 3.7+  
**Target Site:** ksis.eu (Canadian Women's Artistic Gymnastics)

---

**Happy exporting!** 🎉

For questions or issues, refer to the Troubleshooting section above.
