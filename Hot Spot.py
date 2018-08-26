import xlrd, xlsxwriter, xlwt
import sys
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

MEP_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/MEP new.xlsx'
save_file_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/For MEP.xlsx'
stimulations_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/Final 18S .xlsx'

needed_sheet_ws = int(input('Input sheet number 1 - ... of MEP data: '))
needed_sheet_stim = int(input('Input sheet number 1 - ... of stimulation data: '))
row_start_stim = int(input('Input the first row number 1 - ... of stimulation data: ')) - 1
column_stim = int(input('Input COLUMN number 1 - ... of stimulation data: ')) - 1

# open my files
table_RAW = xlrd.open_workbook(MEP_path)
sheet_raw = table_RAW.sheet_by_index(needed_sheet_ws - 1)
sheet_method_RAW = table_RAW.sheet_by_index(1)

table_stim = xlrd.open_workbook(stimulations_path)
sheet_stim = table_stim.sheet_by_index(needed_sheet_stim - 1)
sheet_method_stim = table_stim.sheet_by_index(1)

#creation list of stimulation list
stimulations = []
for stim in range(row_start_stim, row_start_stim + 25):
    value1 = sheet_stim.row_values(stim)[column_stim]
    stimulations.append(int(value1))

#creation list of MEPs
MEP = []
first_row_mep = int(input('Type FIRST row where MEP data starts 1-...:')) - 1
last_row_mep = int(input('Type LAST row where MEP data starts 1-...:')) - 1
column_mep = int(input('Type COLUMN MEP data 1-...: ')) - 1
#extracting data and saving to list
for mep in range(first_row_mep, last_row_mep + 1):
    if mep != '':
        value2 = sheet_raw.row_values(mep)[column_mep]  # if row is empty
    else:
        value2 = 0
    MEP.append(int(value2))

key_for_mep =[]
for k in range(len(MEP)):
    key_for_mep.append(k)

MEP_dict = dict(zip(key_for_mep, MEP))

# finding max
MAX_mep = []
MEP_temporal = []
for rows in stimulations:  # 11, 22..
    for r in range(1, int(rows)):  # 0, 2... 10. Because key for the dict is 0-...
        for value3 in MEP_dict[int(r)]:  # value of key
            MEP_temporal.append(value3)  # temporal list with needed interval
        mx = max(MEP_temporal)  # max in needed interval
        MAX_mep.append(mx)
        del MEP_temporal[:]  # clean the list to rewrite it

# getting coordinates of max

_ordA = ord('A')

def find_val_in_workbook(wbname, val):
    wb = xlrd.open_workbook(wbname)
    for sheet in wb.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.ctype != 2:
                    continue
                if cell.value != val:
                    continue
                if colidx > 26:
                    colchar = chr(int(_ordA + colidx / 26))
                else:
                    colchar = ''
                colchar += chr(_ordA + colidx % 26)
                print('{} -> {}{}: {}'.format(sheet.name, colchar, rowidx+1, cell.value))

max_MEP_x = []
max_MEP_y = []
max_MEP_z = []

#for a in stimulations:
    #for m
    #max() #максимальное из списка MEP
print(stimulations)
print(MEP)
print(MAX_mep)



