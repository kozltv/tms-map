import xlrd, xlsxwriter, xlwt
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

# variables:
table_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/Parameters.xlsx'
new_file_path = '/Users/kseniya/Dropbox/project/Filter tms/New file.xlsx'


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
sheet = table.sheet_by_index(0)

# adding header to the new document
parameter_name = str(sheet.row_values(4 - 1)[29])
w_sheet.write(0, 0, 'Deviation' + parameter_name)  # header 1
w_sheet.write(0, 1, 'Stimulation number')  # header 2

# creation of lists of data from my file basic on information about sheet, column and row
for row in range(4, 29):
    value1 = sheet.row_values(row)[9]
    value2 = sheet.row_values(row)[7]
    parameter_list.append(value1)
    stimulation_list.append(int(value2))

# list of numbers 0:24 for stimulation dictionary
key = []
for k in range(len(parameter_list)):  # key = 0:24
    key.append(k)

# MAIN PART of searching

deviation = 0

while deviation != 100:
    deviation = int(input('Needed deviation from max (in %, mm) '))

    deviation_list.append(deviation)

    dictionary_stim = dict(zip(key, stimulation_list))

    # searching of needed stimulations number

    stim = 0  # for searching of needed stimulation number

    for index, value in enumerate(parameter_list):  # index = 0:24
        value = Decimal(value).quantize(0, ROUND_HALF_UP)  # normal rounding
        if value > deviation:
            stim = index

    if stim != 0:  # for case when needed threshold achieved after first stimulation
        answer.append(dictionary_stim[stim + 1])
    else:
        answer.append(dictionary_stim[stim])

general_table = [[deviation_list], [answer]]  # is not needed

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
