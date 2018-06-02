import xlrd, xlsxwriter, xlwt
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

# TODO Check each list length

# variables:
table_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/Parameters.xlsx'
new_file_path = '/Users/kseniya/Dropbox/project/Filter tms/New file.xlsx'
# table_path = str(input('Input path for original file with data: '))
# new_file_path = str(input('Input path for new excel file to save data: '))

sheetn_in_original_table = int(input('Input SHEET INDEX in the original table (0-...): '))
row_start_data = int(input('Input index of the FIRST ROW of parameters and stimulation in the original table (0-...): '))
row_finish_data = int(input('Input index of LAST ROW of parameters and stimulation in the original table (0-...): '))
column_parameter = int(input('Input index of PARAMETER COLUMN in the original table (0-...): '))
column_stimulation = int(input('Input index of STIMULATION COLUMN in the original table (0-...): '))

# lists:
deviation_list = []  # to save deviation
parameter_list = []  # for iteration of values from parameter column in table
stimulation_list = []  # for iteration of values from stimulation column in table
answer = []  # to save needed stimulation number

# creation of document for data saving
new_book = xlrd.open_workbook(new_file_path)
r_sheet = new_book.sheet_by_index(0)  # read only copy to introspect the file
wb = copy(new_book)  # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0)  # the sheet to write to within the writable copy


# open my file
table = xlrd.open_workbook(table_path)
sheet = table.sheet_by_index(sheetn_in_original_table)
sheet_threshold = table.sheet_by_index(1)

# adding header to the new document
if sheetn_in_original_table == 0:
    parameter_name = str(sheet.row_values(row_start_data - 1)[column_parameter])
else:
    parameter_name = str(input('What is parameter? '))

w_sheet.write(0, 0, 'Deviation ' + parameter_name)  # header 1
w_sheet.write(0, 1, 'Stimulation number')  # header 2

# creation of lists of data from my file basic on information about sheet, column and row

for row in range(row_start_data, row_finish_data + 1):
    if sheet.row_values(row)[column_parameter] != '':  # for case with empty cell
        value1 = sheet.row_values(row)[column_parameter] # if row is empty
        value2 = sheet.row_values(row)[column_stimulation]
        parameter_list.append(value1)
        stimulation_list.append(int(value2))



# list of numbers 0:24 for stimulation dictionary
key = []
for k in range(len(parameter_list)):  # key = 0:24
    key.append(k)

# MAIN PART of searching

deviation = float(0)

while deviation != 100:

    deviation = float(input('Needed deviation from max (in %, mm) '))
    if sheetn_in_original_table == 0:  # case when data is size parameter
        deviation_list.append(int(deviation))
    else:
        deviation_list.append(deviation)

    dictionary_stim = dict(zip(key, stimulation_list))

    # searching of needed stimulations number

    stim = -1  # for searching of needed stimulation number

    for index, value in enumerate(parameter_list):  # index = 0:24
        if sheetn_in_original_table == 0:  # rounding in case with parameter of map size
            value = Decimal(value).quantize(0, ROUND_HALF_UP)  # normal rounding
        else:
            value = Decimal(value).quantize(Decimal("0.1"))  # rounding in case with parameter of map topography
        if value > deviation:
            stim = index

    answer.append(dictionary_stim[stim + 1])

# finish of main part

general_table = [[deviation_list], [answer]]  # is not needed

# getting a code for table
table_code = str(input('Print code for table like "S1D1Ch1": '))
deviation_list.append(table_code)

# data saving to the new file
row1 = 0

for deviat in deviation_list:
    row1 += 1
    w_sheet.write(row1, 0, deviat)

row2 = 0
for stimulation in answer:
    row2 += 1
    w_sheet.write(row2, 1, stimulation)

wb.save(new_file_path)

print(deviation_list, answer)
