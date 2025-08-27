import argparse
import requests
import pandas as pd
import datetime
from openpyxl import load_workbook
import re
from functools import lru_cache
import hashlib
import json
import os

# Global cache for build data to support caching
_build_data_cache = {}

def obligatory_banner():
    ascii_art = r"""
    ____      __                         ____  _____       ________              __            
   /  _/___  / /___  ______  ___        / __ \/ ___/      / ____/ /_  ___  _____/ /_____  _____
   / // __ \/ __/ / / / __ \/ _ \______/ / / /\__ \______/ /   / __ \/ _ \/ ___/ //_/ _ \/ ___/
 _/ // / / / /_/ /_/ / / / /  __/_____/ /_/ /___/ /_____/ /___/ / / /  __/ /__/ ,< /  __/ /    
/___/_/ /_/\__/\__,_/_/ /_/\___/      \____//____/      \____/_/ /_/\___/\___/_/|_|\___/_/     
                                                                                               
by @FlyingPhishy - Why isn't this an accessible feature in Intune?
    """
    print(ascii_art)

def fetch_os_builds(url):
    """Fetch OS builds data from endoflife.date API"""
    print(f"Fetching OS builds data from {url}")
    headers = {"Accept": "application/json"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # Better error handling
    return response.json()

def parse_version_for_os(os_version, os_name):
    """Parse OS version based on OS type - centralized logic"""
    if os_version is None or str(os_version).strip() == '':
        return ''
    
    version_str = str(os_version).strip()
    
    if os_name == 'Windows':
        return '.'.join(version_str.split('.')[:3])
    elif os_name.startswith('Android'):
        # Try major.minor first, fallback to major only
        parts = version_str.split('.')
        return '.'.join(parts[:2]) if len(parts) >= 2 else parts[0]
    elif os_name in ['iOS/iPadOS', 'macOS']:
        return version_str.split('.')[0]
    
    return version_str

def find_matching_build(parsed_version, build_data, os_name):
    """Find matching build data for parsed version"""
    if os_name == 'Windows':
        matching_builds = [b for b in build_data if parsed_version == '.'.join(b['latest'].split('.')[:3])]
        if not matching_builds:
            return None
        # Prefer Windows builds with "(W)" label
        w_build = next((b for b in matching_builds if "(W)" in b.get('releaseLabel', '')), None)
        return w_build or matching_builds[0]
    else:
        # Android, iOS, macOS use cycle matching
        for build in build_data:
            if str(build['cycle']) == parsed_version:
                return build
        
        # Android fallback: try major version only
        if os_name.startswith('Android') and '.' in parsed_version:
            major_version = parsed_version.split('.')[0]
            for build in build_data:
                if str(build['cycle']) == major_version:
                    return build
    
    return None

def calculate_support_status(eol_date):
    """Calculate support status from EOL date"""
    if eol_date is None or eol_date is False:
        return "No EoL"
    
    try:
        eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
        today = datetime.date.today()
        return "Supported" if today <= eol_date else "End of Life"
    except (ValueError, TypeError):
        return "Unknown"

def calculate_version_age(release_date):
    """Calculate human-readable version age from release date"""
    if not release_date or release_date in ['N/A', '', 'null', None]:
        return 'N/A'
    
    try:
        release_date = datetime.datetime.strptime(str(release_date), "%Y-%m-%d").date()
        today = datetime.date.today()
        age_days = (today - release_date).days
        
        if age_days < 0:
            return "Future Release"
        elif age_days == 0:
            return "Today"
        elif age_days == 1:
            return "1 day"
        elif age_days < 30:
            return f"{age_days} days"
        elif age_days < 365:
            months = age_days // 30
            return f"{months} month{'s' if months != 1 else ''}"
        else:
            years = age_days // 365
            remaining_months = (age_days % 365) // 30
            if remaining_months == 0:
                return f"{years} year{'s' if years != 1 else ''}"
            else:
                return f"{years}.{remaining_months} years"
    except (ValueError, TypeError, AttributeError):
        return 'N/A'

@lru_cache(maxsize=1000)  # Cache results to avoid repeated API lookups
def get_os_info(os_version, os_name, build_data_hash):
    """Get comprehensive OS information - cached to prevent duplicate calls"""
    # Retrieve build_data from global cache using hash
    build_data = _build_data_cache.get(build_data_hash, [])
    
    try:
        parsed_version = parse_version_for_os(os_version, os_name)
        
        # Handle empty or invalid parsed versions
        if not parsed_version:
            return {
                'supported': 'Invalid Version',
                'releaseDate': 'N/A',
                'versionAge': 'N/A',
                'eol': 'N/A',
                'releaseLabel': 'N/A',
                'codename': 'N/A',
                'is_latest': False
            }
        
        build = find_matching_build(parsed_version, build_data, os_name)
        
        if not build:
            return {
                'supported': 'Unknown Version',
                'releaseDate': 'N/A',
                'versionAge': 'N/A',
                'eol': 'N/A',
                'releaseLabel': 'N/A',
                'codename': 'N/A',
                'is_latest': False
            }
        
        # Build comprehensive info object
        release_date = build.get('releaseDate', 'N/A')
        info = {
            'supported': calculate_support_status(build.get('eol')),
            'releaseDate': release_date,
            'versionAge': calculate_version_age(release_date),
            'eol': build.get('eol', 'N/A'),
            'releaseLabel': build.get('releaseLabel', 'N/A'),
            'codename': build.get('codename', 'N/A'),
            'is_latest': build.get('latest') == str(os_version)
        }
        
        return info
        
    except (IndexError, AttributeError, ValueError):
        return {
            'supported': 'Invalid Data',
            'releaseDate': 'N/A',
            'versionAge': 'N/A',
            'eol': 'N/A',
            'releaseLabel': 'N/A',
            'codename': 'N/A',
            'is_latest': False
        }

def add_os_columns(df, os_name, build_data):
    """Add OS-specific columns to dataframe - unified approach"""
    # Create a hash of build_data for caching
    build_data_str = json.dumps(build_data, sort_keys=True, default=str)
    build_data_hash = hashlib.md5(build_data_str.encode()).hexdigest()
    
    # Store in global cache
    _build_data_cache[build_data_hash] = build_data
    
    # Get all OS info at once for each version (cached)
    os_info_list = df['OS version'].apply(
        lambda x: get_os_info(str(x), os_name, build_data_hash)
    )
    
    # Add columns based on OS info
    df.loc[:, 'Supported'] = os_info_list.apply(lambda x: x['supported'])
    df.loc[:, 'Release Date'] = os_info_list.apply(lambda x: x['releaseDate'])
    df.loc[:, 'Version Age'] = os_info_list.apply(lambda x: x['versionAge'])
    df.loc[:, 'EOL Date'] = os_info_list.apply(lambda x: x['eol'])
    
    # OS-specific columns
    if os_name == 'Windows':
        df.loc[:, 'Release Label'] = os_info_list.apply(lambda x: x['releaseLabel'])
    elif os_name.startswith('Android'):
        df.loc[:, 'Codename'] = os_info_list.apply(lambda x: x['codename'])
    elif os_name in ['iOS/iPadOS', 'macOS']:
        df.loc[:, 'Latest Version'] = os_info_list.apply(lambda x: x['is_latest'])
    
    return df

def process_excel(file_path, sheet_name, os_data):
    """Process Excel file and add OS support information"""
    print(f"Loading data from {file_path}...")
    
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        # Load the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Validate required columns exist
        required_columns = ['OS', 'OS version']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}. Available columns: {list(df.columns)}")
        
        # Check if dataframe is empty
        if df.empty:
            print("Warning: The input sheet appears to be empty.")
            return
        
        print(f"Found {len(df)} rows with columns: {list(df.columns)}")
        
        for os_name, build_data in os_data.items():
            print(f"Processing {os_name} data...")
            
            # Filter OS-specific data (case-insensitive, handle NaN values)
            os_df = df[df['OS'].astype(str).str.contains(os_name, na=False, case=False)].copy()
            
            if os_df.empty:
                print(f"No {os_name} devices found, skipping...")
                continue
            
            print(f"Found {len(os_df)} {os_name} devices")
            
            # Add all columns at once (efficient, single pass)
            os_df = add_os_columns(os_df, os_name, build_data)
            
            # Write to new sheet
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                clean_os_name = re.sub(r'[^a-zA-Z0-9 ()_-]', '', os_name)
                sheet_name_output = f"{clean_os_name} Versions"
                os_df.to_excel(writer, sheet_name=sheet_name_output, index=False)
                print(f"Created sheet: {sheet_name_output}")
        
        print("Successfully updated the Excel file with new data.")
        
    except FileNotFoundError as e:
        print(f"File error: {e}")
        raise
    except ValueError as e:
        print(f"Data validation error: {e}")
        raise
    except PermissionError:
        print(f"Permission error: Cannot write to {file_path}. Make sure the file is not open in Excel.")
        raise
    except Exception as e:
        print(f"Unexpected error processing Excel file: {e}")
        print(f"Error type: {type(e).__name__}")
        raise

def setup_argparse():
    """Setup command line argument parsing"""
    parser = argparse.ArgumentParser(description='Process OS build numbers to check support status.')
    parser.add_argument('-f', '--file', type=str, required=True, 
                       help='Path to the Excel file to be processed')
    parser.add_argument('-s', '--sheet', type=str, required=True, 
                       help='Name of the sheet to read from')
    return parser

def main():
    """Main execution function"""
    obligatory_banner()
    parser = setup_argparse()
    args = parser.parse_args()
    
    # API endpoints
    api_endpoints = {
        'Windows': "https://endoflife.date/api/windows.json",
        'Android': "https://endoflife.date/api/android.json", 
        'iOS/iPadOS': "https://endoflife.date/api/ios.json",
        'macOS': "https://endoflife.date/api/macos.json"
    }
    
    # Fetch all OS data
    os_data = {}
    for os_name, url in api_endpoints.items():
        try:
            data = fetch_os_builds(url)
            if not data or not isinstance(data, list):
                print(f"Warning: {os_name} API returned invalid data, skipping...")
                continue
            os_data[os_name] = data
            print(f"Successfully fetched {len(data)} {os_name} versions")
        except requests.RequestException as e:
            print(f"Failed to fetch {os_name} data: {e}")
            continue
        except Exception as e:
            print(f"Unexpected error fetching {os_name} data: {e}")
            continue
    
    if not os_data:
        print("No OS data could be fetched. Exiting.")
        return
    
    process_excel(args.file, args.sheet, os_data)

if __name__ == "__main__":
    main()