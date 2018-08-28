import xlrd, xlsxwriter, xlwt
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

# PATHS
MEP_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/MEP new.xlsx'
save_file_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/For MEP.xlsx'
stimulations_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/Final 18S .xlsx'
new_document_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/for hot spot.xlsx'

# getting VARIABLES
subject_number = int(input('Type subject number: '))
day = str(input('What is day? 1 or 2: '))
chanel = str(input('What is chanel? 1, 2, or 3: '))

needed_sheet_ws = int(input('Input sheet number 0 - ... of MEP data: '))

if 1 <= subject_number <= 6:
    needed_sheet_stim = 0
elif 7 <= subject_number <= 12:
    needed_sheet_stim = 1
else:
    needed_sheet_stim = 2

row_start_stim = int(input('Input the first row number 1 - ... of stimulation data: ')) - 1
print('Please WAIT')
column_stim = 0  # CHECK whether the column is actual

# open my files
table_RAW = xlrd.open_workbook(MEP_path)
sheet_raw = table_RAW.sheet_by_index(needed_sheet_ws)

table_stim = xlrd.open_workbook(stimulations_path)
sheet_stim = table_stim.sheet_by_index(needed_sheet_stim)

# creation list of stimulation list

stimulations = []
for stim in range(row_start_stim, row_start_stim + 25):
    value1 = sheet_stim.row_values(stim)[column_stim]
    stimulations.append(int(value1))

# creation list of MEPs

MEP = []
first_row_mep = 6  # int(input('Type FIRST row where MEP data starts 1-...: ')) - 1
last_row_mep = stimulations[-1] + first_row_mep - 1  # int(input('Type LAST row where MEP data finish 1-...:')) - 1


# Calculation for DAY 1 CHANEL 1

column_mep = 0
if day == '1':
    if chanel == '1':
        column_mep = 24
    elif chanel == '2':
        column_mep = 26
    else:
        column_mep = int(input('Type COLUMN MEP data 0-... for ' + 'Day ' + day + ' Ch ' + chanel + ': '))
else:
    column_mep = int(input('Type COLUMN MEP data 0-... for ' + 'Day ' + day + ' Ch ' + chanel + ': '))

# extracting data and saving to MEP list


def extracting_mep_data(first_row, last_row, save_list, sheet, column):
    for mep in range(first_row, last_row + 1):
        value2 = sheet.row_values(mep)[column]  # if row is empty
        if value2 == '' or value2 == '-':
            value2 = 0
        save_list.append(int(value2))


extracting_mep_data(first_row_mep, last_row_mep, MEP, sheet_raw, column_mep)

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

#  NEW Function fpr calculation of max
def calculation_of_MAX(stimulations_list, dict_of_MEP, save_list):
    temporal_list = []
    for rows in stimulations_list:  # 11, 22..
        for r in range(0, int(rows)):  # 0, 2... 10. Because key for the dict is 0-...
            temporal_list.append(dict_of_MEP[r])  # temporal list with needed interval
        mx = max(temporal_list)  # max in needed interval
        save_list.append(mx)
        del temporal_list[:]  # clean the list to rewrite it

MAX_mep = []
calculation_of_MAX(stimulations, MEP_dict, MAX_mep)


# MAIN PART. STEP 2 - getting coordinates of max

# NEW Function to find row number of max

def search_rowind_of_max(list_of_max, orig_doc_with_data, column_ind, save_list):
    for val in list_of_max:
        for rowind, value in enumerate(orig_doc_with_data.col_values(column_ind)):
            if value != val:
                continue
            else:
                save_list.append(rowind + 1)

max_row = []
search_rowind_of_max(MAX_mep, sheet_raw, column_mep, max_row)


# obtaining of 'EF max loc' x;y;z coordinates
EF_max_x_column = 0
if day == '1':
    EF_max_x_column = 19
else:
    EF_max_x_column = int(input('Type COLUMN number for EF MAX X coordinate 0-...: '))

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
"""
new_document_wb = xlrd.open_workbook(new_document_path)
r_sheet = new_document_wb.sheet_by_index(0)
wb = copy(new_document_wb)
wb_sheet = wb.get_sheet(0)
"""
# NEW Function creates a excel document and returns variable with link for the document
def saving_to_doc(path_for_new_doc):
    new_document_wb = xlrd.open_workbook(path_for_new_doc)
    new_document_wb.sheet_by_index(0)
    wb = copy(new_document_wb)
    sheet2 = wb.get_sheet(0)
    return sheet2

# HEADER writing


wb_sheet = saving_to_doc(new_document_path)
wb_sheet.write(0, 0, 'Subject ' + str(subject_number) + ' Day ' + day + 'Ch ' + chanel)
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

