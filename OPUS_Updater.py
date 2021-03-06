from openpyxl import *
import pandas as pd
import os
import re
from traits.api import HasTraits, Str, Range, Bool
from traitsui.api import View, Item

### Read a list of all protocols in a given directory

# change into directory containing OPUS protocol doc
os.chdir(r'N:\Departments\R_D\NexGen\Assay Integration\04-APS File Requests\Gastro Intestinal Protocols')
#returns list of everything in current working directory
dir_list = os.listdir(r'N:\Departments\R_D\NexGen\Assay Integration\04-APS File Requests\Gastro Intestinal Protocols\Requests')
aps_req_list = list()
for req in dir_list:
    if 'GI' in req:
        aps_req_list.append(req)

xl_protocol_log = load_workbook(filename = r'Protocols since Nov 2016.xlsx')

update_protocols = list()

aps_req_list

### Read in OPUS List from Excel Doc

# wb = protocols since Nov 2017 sheet, displays all sheet names
xl_protocol_log.get_sheet_names()
# change to whatever sheet like to have active by sheet name, access like dic keys
prot_log_sheet = xl_protocol_log['Sheet1']
#, this returns tuple of all rows with contents or formatting
prot_log_cells = tuple(prot_log_sheet.rows)
#impt here the value on the cell object shows what
opus_dev_log = list()
#Corrrrrrrect!
#checks if the cell value isn't None and has 'GI' to only include only OPUS names w/ no trivial excel formatting
for cell in prot_log_cells:
    if cell[0].value is not None and 'GI' in cell[0].value:
        opus_dev_log.append(cell[0].value)

logged_opi = set()
for entry in opus_dev_log:
    try:
        opus_req = re.match((".*(GI-\d+)-.*[^OB]"),entry)
        logged_opi.add(opus_req.groups(1)[0])
    except:
        continue
#only return the set of unique values from the list

### Filter out protocols already updated

reqs_to_update = list()
for entry in aps_req_list:
    #checks if no GI skips that iteration to prevent being flagged and time waste
    if 'GI' not in entry or re.search(".*GI-\d\d\d\d.+",entry) is None: continue
    #extracts matching seq as a list, to access raw string b/c = list need index the element
    if re.findall(".*(GI-\d\d\d\d).+",entry)[0] not in logged_opi:
        reqs_to_update.append(entry)

### Access files that've not yet been updated on the dev log sheet

os.chdir(r'N:\Departments\R_D\NexGen\Assay Integration\04-APS File Requests\Gastro Intestinal Protocols\Requests')
# creates the absolute file path for for the protocols to be updated
fi_path = os.path.join(os.getcwd(),reqs_to_update[2])

###Open Request to update

req_bk = load_workbook(filename = fi_path)
req_sheet = req_bk.active

for rows in req_sheet.rows:
    for cell in rows:
        print(cell.value)

# Access specific values from
req_sheet['C8'].value
