import argparse
import requests
import pandas as pd
import datetime
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
            for build in build_data:
                if str(build['cycle']) == str(os_version):
                    eol_date = build.get('eol', None)
                    if eol_date is False:  # Checking for 'false' explicitly, as it represents no end of life
                        return "No EoL"
                    if isinstance(eol_date, str):
                        eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                        return "Supported" if today <= eol_date else "End of Life"
                    
        elif os_name in ['iOS/iPadOS', 'macOS']:
            major_version = str(os_version).split('.')[0]
            for build in build_data:
                if str(build['cycle']) == str(major_version):
                    eol_date = build.get('eol', None)
                    if eol_date is False:
                        return "No EoL"
                    if isinstance(eol_date, str):
                        eol_date = datetime.datetime.strptime(eol_date, "%Y-%m-%d").date()
                        return "Supported" if today <= eol_date else "End of Life"
                    
        return "Unknown Version"
    except (IndexError, AttributeError, ValueError):  # Handling cases with empty or malformed data
        return "Invalid Data"


def is_latest(os_version, build_data):
    for build in build_data:
        if build.get('latest') == os_version:
            return True
    return False

def process_excel(file_path, sheet_name, os_data):
    print(f"Loading data from {file_path}...")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        for os_name, build_data in os_data.items():
            print(f"Processing {os_name} data...")
            os_df = df[df['OS'].str.contains(os_name, na=False, case=False)].copy() 

            if os_name == 'Windows':
                os_df.loc[:, 'Supported'] = os_df['OS version'].apply(
                    lambda x: is_supported(x, build_data, os_name)['supported'] if is_supported(x, build_data, os_name) else "Unknown Version"
                )
                os_df.loc[:, 'Release Label'] = os_df['OS version'].apply(
                    lambda x: is_supported(x, build_data, os_name)['releaseLabel'] if is_supported(x, build_data, os_name) else "N/A"
                )
                os_df.loc[:, 'Release Date'] = os_df['OS version'].apply(
                    lambda x: is_supported(x, build_data, os_name)['releaseDate'] if is_supported(x, build_data, os_name) else "N/A"
                )
                os_df.loc[:, 'EOL Date'] = os_df['OS version'].apply(
                    lambda x: is_supported(x, build_data, os_name)['eol'] if is_supported(x, build_data, os_name) and 'eol' in is_supported(x, build_data, os_name) else "N/A"
                )
            else:
                # For Android, iOS/iPadOS, and macOS, continue using 'cycle'
                os_df.loc[:, 'Supported'] = os_df.apply(lambda x: is_supported(x['OS version'], build_data, os_name), axis=1)
                os_df.loc[:, 'Release Date'] = os_df['OS version'].apply(lambda x: next((build['releaseDate'] for build in build_data if str(build['cycle']) == str(x).split('.')[0]), "N/A"))
                os_df.loc[:, 'EOL Date'] = os_df['OS version'].apply(lambda x: next((build.get('eol', "N/A") for build in build_data if str(build['cycle']) == str(x).split('.')[0]), "N/A"))
                if os_name.startswith('Android'):
                    os_df.loc[:, 'Codename'] = os_df['OS version'].apply(lambda x: next((build.get('codename', "N/A") for build in build_data if str(build['cycle']) == str(x)), "N/A"))
                elif os_name in ['iOS/iPadOS', 'macOS']:
                    os_df.loc[:, 'Latest Version'] = os_df['OS version'].apply(
                    lambda x: is_latest(re.match(r"([\d\.]+)", str(x)).group(1) if re.match(r"([\d\.]+)", str(x)) else str(x), build_data)
                )

            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                clean_os_name = re.sub(r'[^a-zA-Z0-9 ()_-]', '', os_name)
                os_df.to_excel(writer, sheet_name=f"{clean_os_name} Versions", index=False)
        
        print("Successfully updated the Excel file with new data.")
    except Exception as e:
        print(f"Failed to process the Excel file: {e}")


def setup_argparse():
    parser = argparse.ArgumentParser(description='Process OS build numbers to check support status.')
    parser.add_argument('-f', '--file', type=str, required=True, help='Path to the Excel file to be processed')
    parser.add_argument('-s', '--sheet', type=str, required=True, help='Name of the sheet to read from')
    return parser

def main():
    obligatory_banner()
    parser = setup_argparse()
    args = parser.parse_args()
    
    os_data = {
        'Windows': fetch_os_builds("https://endoflife.date/api/windows.json"),
        'Android': fetch_os_builds("https://endoflife.date/api/android.json"),
        'iOS/iPadOS': fetch_os_builds("https://endoflife.date/api/ios.json"),
        'macOS': fetch_os_builds("https://endoflife.date/api/macos.json")
    }
    
    process_excel(args.file, args.sheet, os_data)

if __name__ == "__main__":
    main()