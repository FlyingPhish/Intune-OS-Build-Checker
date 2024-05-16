# Intune OS Build Checker
This Python script analyzes a given XLSX file containing OS version information for devices managed by Microsoft Intune. It updates the file with new sheets that contain additional details such as support status, release label, release date, and end-of-life (EOL) date for each OS version. The script leverages data from the "endoflife.date" API to determine the support status and other relevant information.

## Important!
- This script needs to use an XLSX as CSV isn't suitable for having multiple sheets so convert the CSV file to XLSX before usage.
- Ensure that any date columns are formatted as Date.
- I would highly recommend adding necessary columns via the Intune web GUI before clicking export. I recommend the following:
    - Device Name
    - Compliance
    - Device state
    - EAS activated
    - Encrypted
    - Enrollment date
    - Intune registered
    - Join type
    - Last check-in
    - Managed by
    - Microsoft Entra Device ID
    - Microsoft Entra registered
    - Model
    - OS
    - OS version
    - Ownership
    - Primary user UPN

- **I WOULD HIGHLY AVOID THE FOLLOWING:**
    - Compliance grace period expiration
    - Any of the EAS ones that I haven't mentioned above, especially Last EAS sync time.
        - Why? The data formatting is bollocks and will only increase the amount of time prepping your spreadsheet.
        - If you see any terminal warning output, then take note of the column letter and investigate. I bet it will be due to date/time, or something relating to EAS.

## How to Export Devices CSV
1. Go to All devices within Intune (https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/allDevices).
2. Click Columns and follow my advice (or not) listed in the Important section.
3. Click Export.
4. When downloaded, open the CSV and save as XLSX.
5. Ensure any date columns are properly formatted.

## Script Overview
### Script Help
```
    ____      __                         ____  _____       ________              __            
   /  _/___  / /___  ______  ___        / __ \/ ___/      / ____/ /_  ___  _____/ /_____  _____
   / // __ \/ __/ / / / __ \/ _ \______/ / / /\__ \______/ /   / __ \/ _ \/ ___/ //_/ _ \/ ___/
 _/ // / / / /_/ /_/ / / / /  __/_____/ /_/ /___/ /_____/ /___/ / / /  __/ /__/ ,< /  __/ /    
/___/_/ /_/\__/\__,_/_/ /_/\___/      \____//____/      \____/_/ /_/\___/\___/_/|_|\___/_/     
                                                                                               
by @FlyingPhishy - Why isn't this an accessible feature in Intune?
    
usage: parse-intune-xlsx.py [-h] -f FILE -s SHEET

Process OS build numbers to check support status.

options:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  Path to the Excel file to be processed
  -s SHEET, --sheet SHEET
                        Name of the sheet to read from
```

### Output
```
python3 parse-intune-xlsx.py  -f ~/Desktop/intunev2.xlsx -s All

    ____      __                         ____  _____       ________              __            
   /  _/___  / /___  ______  ___        / __ \/ ___/      / ____/ /_  ___  _____/ /_____  _____
   / // __ \/ __/ / / / __ \/ _ \______/ / / /\__ \______/ /   / __ \/ _ \/ ___/ //_/ _ \/ ___/
 _/ // / / / /_/ /_/ / / / /  __/_____/ /_/ /___/ /_____/ /___/ / / /  __/ /__/ ,< /  __/ /    
/___/_/ /_/\__/\__,_/_/ /_/\___/      \____//____/      \____/_/ /_/\___/\___/_/|_|\___/_/     
                                                                                               
by @FlyingPhishy - Why isn't this an accessible feature in Intune?
    
Fetching OS builds data from https://endoflife.date/api/windows.json
Fetching OS builds data from https://endoflife.date/api/android.json
Fetching OS builds data from https://endoflife.date/api/ios.json
Fetching OS builds data from https://endoflife.date/api/macos.json
Loading data from /Users/X/Desktop/intunev2.xlsx...
Processing Windows data...
Processing Android data...
Processing iOS/iPadOS data...
Processing macOS data...
Successfully updated the Excel file with new data.
```

```md
| Supported       | Release Label | Release Date | EOL Date   |
| --------------- | ------------- | ------------ | ---------- |
| Supported       | 10 22H2       | 2022-10-18   | 2025-10-14 |
| Supported       | 11 23H2 (W)   | 2023-10-31   | 2025-11-11 |
| End of Life     | 10 1809 (W)   | 2018-11-13   | 2020-11-10 |
| Unknown Version | N/A           | N/A          | N/A        |
| End of Life     | 10 1809 (W)   | 2018-11-13   | 2020-11-10 |
| End of Life     | 10 1607 (W)   | 2016-08-02   | 2018-04-10 |
```
## Features

- Fetches the latest OS build information from the "endoflife.date" API for Windows, Android, iOS/iPadOS, and macOS.
- Processes an Excel file containing Intune device information and updates the file with additional columns:
    - Support status (Supported, End of Life, or Unknown Version)
    - Release label (e.g., 11 22H2 (W) for Windows)
    - Release date
    - EOL date
- Handles version matching and selection based on specific criteria (e.g., favors "(W)" release label for Windows versions).
- Generates separate sheets in the Excel file for each OS type (Windows, Android, iOS/iPadOS, macOS) with the updated information.

## Prerequisites

- Python 3.x
- Required Python packages:
    - pandas
    - openpyxl
    - requests

## Usage
### Setup 
1. `python3 -m venv .venv`
2. `source .venv/bin/activate`
3. `pip3 install -r requirements.txt`

### Running
1. `python3 parse-intune-xlsx.py -f file.xlsx -s SheetName`
- The script will process the Excel file, fetch the latest OS build information from the "endoflife.date" API, and update the file with the additional columns.
- Once the script finishes execution, open the updated Excel file to view the results. Each OS type will have a separate sheet with the updated information.

## Credit
Big shoutout to endoflife-date (endoflife.date) and all the contributors for making that useful resource, especially with the API.

- Github: https://github.com/endoflife-date/endoflife.date
- Site: https://endoflife.date/