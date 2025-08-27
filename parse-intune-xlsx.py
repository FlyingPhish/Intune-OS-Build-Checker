import argparse
import requests
import pandas as pd
import datetime
import json
from openpyxl import load_workbook
import re

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
    print(f"Fetching OS builds data from {url}")
    headers = {"Accept": "application/json"}
    response = requests.get(url, headers=headers)
    builds = response.json()
    return builds

def is_supported(os_version, build_data, os_name):
    try:
        today = datetime.date.today()
        if os_name == 'Windows':
            major_build = str(os_version).split('.')[:3]  # Getting the third part of the version number
            matching_builds = [build for build in build_data if major_build == build['latest'].split('.')[:3]]
            
            if len(matching_builds) == 1:
                build = matching_builds[0]
            elif len(matching_builds) > 1:
                w_build = next((build for build in matching_builds if "(W)" in build.get('releaseLabel', '')), None)
                build = w_build if w_build else matching_builds[0]
            else:
                return None
            
            eol_date = build.get('eol')
            if eol_date is None:
                build['supported'] = "No EoL"
            else:
                eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                build['supported'] = "Supported" if today <= eol_date else "End of Life"
            return build
        
        elif os_name.startswith('Android'):
            major_version = '.'.join(str(os_version).split('.')[:2])  # Dropping the third period and anything after
            for build in build_data:
                if str(build['cycle']) == major_version:
                    eol_date = build.get('eol', None)
                    if eol_date is False:  # Checking for 'false' explicitly, as it represents no end of life
                        build['supported'] = "No EoL"
                    elif isinstance(eol_date, str):
                        eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                        build['supported'] = "Supported" if today <= eol_date else "End of Life"
                    else:
                        build['supported'] = "Unknown"
                    return build
                
            if not any(str(build['cycle']) == major_version for build in build_data):
                major_version = str(os_version).split('.')[0]
                for build in build_data:
                    if str(build['cycle']) == major_version:
                        eol_date = build.get('eol', None)
                        if eol_date is False:  # Checking for 'false' explicitly, as it represents no end of life
                            build['supported'] = "No EoL"
                        elif isinstance(eol_date, str):
                            eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                            build['supported'] = "Supported" if today <= eol_date else "End of Life"
                        else:
                            build['supported'] = "Unknown"
                        return build

            return {"supported": "Unknown Version", "releaseDate": "N/A", "eol": "N/A", "codename": "N/A"}
                    
        elif os_name in ['iOS/iPadOS', 'macOS']:
            major_version = str(os_version).split('.')[0]
            for build in build_data:
                if str(build['cycle']) == str(major_version):
                    eol_date = build.get('eol', None)
                    if eol_date is False:
                        build['supported'] = "No EoL"
                    elif isinstance(eol_date, str):
                        eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                        build['supported'] = "Supported" if today <= eol_date else "End of Life"
                    else:
                        build['supported'] = "Unknown"
                    return build
                    
        return {"supported": "Unknown Version", "releaseDate": "N/A", "eol": "N/A", "codename": "N/A"}
    except (IndexError, AttributeError, ValueError):  # Handling cases with empty or malformed data
        return "Invalid Data"


def is_latest(os_version, build_data):
    for build in build_data:
        if build.get('latest') == os_version:
            return True
    return False

def extract_os_version_from_ldap(os_version_attr):
    """Extract version number from LDAP operatingSystemVersion format like '10.0 (19045)'"""
    if isinstance(os_version_attr, list) and len(os_version_attr) > 0:
        version_str = str(os_version_attr[0])
        # Extract version from format like "10.0 (19045)" -> "10.0.19045"
        match = re.match(r"(\d+\.\d+)\s*\((\d+)\)", version_str)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        # Fallback to just the version part before parentheses
        return version_str.split('(')[0].strip()
    return str(os_version_attr)

def load_json_computers(json_file_path):
    """Load and filter JSON data to only include objects with OS information"""
    print(f"Loading computer data from {json_file_path}...")
    
    computers = []
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Filter objects that have both operatingSystem and operatingSystemVersion
    for item in data:
        attributes = item.get('attributes', {})
        if ('operatingSystem' in attributes and 
            'operatingSystemVersion' in attributes):
            
            os_name = attributes['operatingSystem'][0] if isinstance(attributes['operatingSystem'], list) else attributes['operatingSystem']
            os_version = extract_os_version_from_ldap(attributes['operatingSystemVersion'])
            computer_name = attributes.get('cn', ['Unknown'])[0] if isinstance(attributes.get('cn'), list) else attributes.get('cn', 'Unknown')
            
            computers.append({
                'Computer Name': computer_name,
                'OS': os_name,
                'OS version': os_version,
                'Distinguished Name': item.get('dn', 'N/A')
            })
    
    print(f"Found {len(computers)} computers with OS information")
    return pd.DataFrame(computers)

def apply_os_analysis(df, os_data):
    """Apply OS version analysis to DataFrame - reused logic from Excel processing"""
    # Initialize new columns
    df['Supported'] = 'Unknown'
    df['Release Date'] = 'N/A'
    df['EOL Date'] = 'N/A'
    
    for os_name, build_data in os_data.items():
        print(f"Processing {os_name} data...")
        # Get mask for this OS type
        os_mask = df['OS'].str.contains(os_name, na=False, case=False)
        
        if not os_mask.any():
            continue
            
        if os_name == 'Windows':
            # Add Windows-specific columns
            if 'Release Label' not in df.columns:
                df['Release Label'] = 'N/A'
            
            for idx in df[os_mask].index:
                os_version = df.loc[idx, 'OS version']
                support_info = is_supported(os_version, build_data, os_name)
                
                if support_info:
                    df.loc[idx, 'Supported'] = support_info.get('supported', 'Unknown Version')
                    df.loc[idx, 'Release Label'] = support_info.get('releaseLabel', 'N/A')
                    df.loc[idx, 'Release Date'] = support_info.get('releaseDate', 'N/A')
                    df.loc[idx, 'EOL Date'] = support_info.get('eol', 'N/A')
                else:
                    df.loc[idx, 'Supported'] = 'Unknown Version'
                    
        else:
            # For Android, iOS/iPadOS, and macOS
            if os_name.startswith('Android') and 'Codename' not in df.columns:
                df['Codename'] = 'N/A'
            elif os_name in ['iOS/iPadOS', 'macOS'] and 'Latest Version' not in df.columns:
                df['Latest Version'] = False
                
            for idx in df[os_mask].index:
                os_version = df.loc[idx, 'OS version']
                support_info = is_supported(os_version, build_data, os_name)
                
                if support_info:
                    df.loc[idx, 'Supported'] = support_info.get('supported', 'Unknown Version')
                    df.loc[idx, 'Release Date'] = support_info.get('releaseDate', 'N/A')
                    df.loc[idx, 'EOL Date'] = support_info.get('eol', 'N/A')
                    
                    if os_name.startswith('Android'):
                        df.loc[idx, 'Codename'] = support_info.get('codename', 'N/A')
                    elif os_name in ['iOS/iPadOS', 'macOS']:
                        # Extract version for latest check
                        version_match = re.match(r"([\d\.]+)", str(os_version))
                        check_version = version_match.group(1) if version_match else str(os_version)
                        df.loc[idx, 'Latest Version'] = is_latest(check_version, build_data)
                else:
                    df.loc[idx, 'Supported'] = 'Unknown Version'
    
    return df

def process_excel(file_path, sheet_name, os_data):
    print(f"Loading data from {file_path}...")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df = apply_os_analysis(df, os_data)

        # Write results back to Excel
        for os_name in os_data.keys():
            os_df = df[df['OS'].str.contains(os_name, na=False, case=False)].copy()
            if len(os_df) > 0:
                with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    clean_os_name = re.sub(r'[^a-zA-Z0-9 ()_-]', '', os_name)
                    os_df.to_excel(writer, sheet_name=f"{clean_os_name} Versions", index=False)
        
        print("Successfully updated the Excel file with new data.")
    except Exception as e:
        print(f"Failed to process the Excel file: {e}")

def process_json(json_file_path, os_data):
    """Process JSON file and create Excel output with OS analysis"""
    try:
        # Load and filter JSON data
        df = load_json_computers(json_file_path)
        
        if len(df) == 0:
            print("No computers with OS information found in JSON file.")
            return
        
        # Apply OS analysis using existing logic
        df = apply_os_analysis(df, os_data)
        
        # Create output Excel file
        output_file = json_file_path.replace('.json', '_os_analysis.xlsx')
        print(f"Creating Excel output: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write summary sheet with all data
            df.to_excel(writer, sheet_name='All Computers', index=False)
            
            # Write separate sheets per OS type
            for os_name in os_data.keys():
                os_df = df[df['OS'].str.contains(os_name, na=False, case=False)].copy()
                if len(os_df) > 0:
                    clean_os_name = re.sub(r'[^a-zA-Z0-9 ()_-]', '', os_name)
                    os_df.to_excel(writer, sheet_name=f"{clean_os_name} Versions", index=False)
        
        print(f"Successfully created {output_file} with OS analysis.")
        
    except Exception as e:
        print(f"Failed to process the JSON file: {e}")

def setup_argparse():
    parser = argparse.ArgumentParser(description='Process OS build numbers to check support status.')
    
    # Create mutually exclusive group for input file types
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument('-f', '--file', type=str, help='Path to the Excel file to be processed')
    input_group.add_argument('-jf', '--json-file', type=str, help='Path to the JSON file from ldapdomaindump')
    
    parser.add_argument('-s', '--sheet', type=str, help='Name of the sheet to read from (required for Excel files)')
    
    return parser

def main():
    obligatory_banner()
    parser = setup_argparse()
    args = parser.parse_args()
    
    # Validate arguments
    if args.file and not args.sheet:
        parser.error("--sheet is required when using --file")
    
    # Fetch OS data
    os_data = {
        'Windows': fetch_os_builds("https://endoflife.date/api/windows.json"),
        'Android': fetch_os_builds("https://endoflife.date/api/android.json"),
        'iOS/iPadOS': fetch_os_builds("https://endoflife.date/api/ios.json"),
        'macOS': fetch_os_builds("https://endoflife.date/api/macos.json")
    }
    
    # Process based on input type
    if args.file:
        process_excel(args.file, args.sheet, os_data)
    elif args.json_file:
        process_json(args.json_file, os_data)

if __name__ == "__main__":
    main()