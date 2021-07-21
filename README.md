# meraki-api-configure-ms-from-cisco-config

Summary:

The purpose of this script is to parse through running configs of Cisco Catalyst switches, convert interface information
and push those settings to Meraki MS series switches via API. Only L2 interface configs are supported at this time (no L3 SVIs,
no routing). This script doesn't support switch stacking. Port channels are experimental (able to parse, but no API call).

Spreadsheet with Cisco/Meraki switch information is required (example spreadsheet is uploaded in this repository). Catalyst 
running configs should be pre-downloaded in the same folder with this script. Naming convention for running config files 
is <ip_address>_show_run.txt. IP addresses of the switch configs have to correspond to IP addresses in the spreadsheet. 
For example: in the spreadsheet we are replacing switches 10.0.0.10 and 10.0.0.11. Config files for those two devices should 
exist in the script folder (10.0.0.10_show_run.txt and 10.0.0.11_show_run.txt). Script will look through all IP addresses
of old switches in the spreadsheet, match them to <ip_address>_show_run.txt files, parse them and use meraki serial number
that is configured in the spreadheet to upload parsed configs via API.

The following Meraki switch port variables are available/configurable via this script: port number, port description,
enable status, port type, access vlan, voice vlan, rstp, stpGuard.

By default, this script will not attempt to claim serial numbers of Meraki devices in the spreadsheet. If you need to claim
those serials and move them to the right network, please see step #4 under How To Run.
 
Requirements:

1) Interpreter: Python 3.8.0+
2) Python Packages: requests, json, openpyxl, re
3) Excel Spreadsheet - .XLSX format
4) API support for the Organization is enabled in Meraki Dashboard. Admin has generated their custom API key.
5) Running config files have been acquired for all old devices that are to be replaced. 

How To Run:

1) Attached is a spreadsheet that can be used as a template. Custom spreadsheet can be used, but column variables must be changed under
   PARAMETERS section. Line 19-24.
2) Open parse_and_copy_config.py with your favorite text editor and edit PARAMETERS sections of the script:
    1) Lines 10-14 is mandatory.
    2) Line 19-24 are required if using custom spreadsheet.
3) Upload all running config files from Cisco switches to the same directory with the script. Spreadsheet should be in this directory as well.
   Run python3 parse_and_copy_config.py in your terminal.
4) By default, script assumes that meraki MS switches have been claimed and moved to their dedicated network in Meraki dashboard.
To claim meraki MS switches into your network using this script, uncomment Line 331 (meraki_claim_serial function).
   