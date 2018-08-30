import xlrd, xlsxwriter, xlwt
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

# PATHS
MEP_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/MEP new.xlsx'
new_document_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/for hot spot.xlsx'

# getting VARIABLES
subject_number = int(input('Type SUBJECT number: '))
day = int(input('What is DAY? 1 or 2: '))
chanel = int(input('What is CHANEL? 1, 2, or 3: '))
needed_sheet_ws = subject_number

print('Please WAIT')

# open my files
table_RAW = xlrd.open_workbook(MEP_path)
column_mep_sheet = table_RAW.sheet_by_index(0)
sheet_raw = table_RAW.sheet_by_index(needed_sheet_ws)


sheet_stim = table_RAW.sheet_by_index(0)

# creation list of stimulation list
row_start_stim = 1
column_stim = subject_number
stimulations = []
for stim in range(row_start_stim, 26):
    value1 = sheet_stim.row_values(stim)[column_stim]
    stimulations.append(int(value1))

# creation list of MEPs

MEP = []
first_row_mep = 6  # int(input('Type FIRST row where MEP data starts 1-...: ')) - 1
last_row_mep = stimulations[-1] + first_row_mep - 1  # int(input('Type LAST row where MEP data finish 1-...:')) - 1

# obtaining of column number for MEP data


col_ind_col = 0  # column number to receive column index
col_int_row = subject_number + 29
col_ind_EF_max_col = 0
if day == 1:
    col_ind_col = chanel + 2
    col_ind_EF_max_col = 2
else:
    col_ind_col = chanel + 6
    col_ind_EF_max_col = 6

column_mep = int(column_mep_sheet.row_values(col_int_row)[col_ind_col])

# extracting data and saving to MEP list
for mep in range(first_row_mep, last_row_mep + 1):
    value2 = sheet_raw.row_values(mep)[column_mep]  # if row is empty
    if value2 == '' or value2 == '-':
        value2 = 0
    MEP.append(int(value2))

key_for_mep = []
for k in range(len(MEP)):
    key_for_mep.append(k)

MEP_dict = dict(zip(key_for_mep, MEP))


# MAIN PART. STEP 1 - calculation of MAX

# for further calculations whether last stimulation number equals to length of MEP_dict is crucial
# test:

if stimulations[-1] != len(MEP_dict):
    print('MEP data doesnt contain needed stimulation points. \n'
          'Check whether LAST ROW MEP data is determined and typed correct')
    quit()

MAX_mep = []
max_row = []
MEP_temporal = []
row_temporal = []

for rows in stimulations:  # 11, 22..
    for r in range(0, int(rows)):  # 0, 2... 10. Because key for the dict is 0-...
        MEP_temporal.append(MEP_dict[r])  # temporal list with needed interval
    mx = max(MEP_temporal)  # max in needed interval
    MAX_mep.append(mx)
    for index, val in enumerate(MEP_temporal):  # getting index row of max
        if val != mx:
            continue
        else:
            row_temporal.append(index + 7)  # index is added to temporal because of cases when max appears in a list again
    max_row.append(row_temporal[-1])
    del row_temporal[:]
    del MEP_temporal[:]  # clean the list to rewrite it

# MAIN PART. STEP 2 - getting coordinates of max

""" # some code from internet for getting row number 
_ordA = ord('A')

def find_val_in_workbook(wb_path, val, sheetid):
    wb = xlrd.open_workbook(wb_path)
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
"""

# obtaining of 'EF max loc' x;y;z coordinates

EF_max_x_column = int(column_mep_sheet.row_values(col_int_row)[col_ind_EF_max_col])  # getting the information from MEP document
EF_max_y_column, EF_max_z_column = EF_max_x_column + 1, EF_max_x_column + 2

max_MEP_x = []
max_MEP_y = []
max_MEP_z = []

for max_row_index in max_row:
    max_MEP_x.append(sheet_raw.row_values(max_row_index - 1)[EF_max_x_column])
    max_MEP_y.append(sheet_raw.row_values(max_row_index - 1)[EF_max_y_column])
    max_MEP_z.append(sheet_raw.row_values(max_row_index - 1)[EF_max_z_column])

hot_spot_table = [stimulations, MAX_mep, max_MEP_x, max_MEP_y, max_MEP_z, max_row]

# SAVING to a document
new_document_wb = xlrd.open_workbook(new_document_path)
r_sheet = new_document_wb.sheet_by_index(0)
wb = copy(new_document_wb)
wb_sheet = wb.get_sheet(0)

# HEADER writing
wb_sheet.write(0, 0, 'Subject ' + str(subject_number) + ' Day ' + str(day) + ' Ch ' + str(chanel))
header = ['Stimulation', 'Max mep', 'EF loc x', 'EF loc y', 'EF loc z', 'max row ind']

column0 = -1
for name in header:
    column0 += 1
    wb_sheet.write(1, column0, name)

column = -1

# writing of data to a document
for list in hot_spot_table:
    row1 = 1
    column += 1
    for value in list:
        row1 += 1
        wb_sheet.write(row1, column, value)

wb.save(new_document_path)
print('Good job! Let`s open "for hot spot" file through Numbers')
