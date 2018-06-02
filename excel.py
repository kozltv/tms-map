import xlrd, xlsxwriter, xlwt
from xlutils.copy import copy
from decimal import Decimal, ROUND_HALF_UP

deviation = int(input('Needed deviation from max (in %, mm) '))
deviation_list = []
deviation_list.append(deviation)

table_path = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/Parameters.xlsx'
new_file_path = '/Users/kseniya/Dropbox/project/Filter tms/New file.xlsx'
table = xlrd.open_workbook(table_path)
sheet = table.sheet_by_index(0)

parameter_list = []
stimulation_list = []

for row in range(4, 29):
    value1 = sheet.row_values(row)[9]
    value2 = sheet.row_values(row)[7]
    parameter_list.append(value1)
    stimulation_list.append(int(value2))

key = []
for k in range(len(parameter_list)):  # key = 0:24
    key.append(k)

dictionary_stim = dict(zip(key, stimulation_list))

stim = 0
answer = []

for index, value in enumerate(parameter_list):  # index = 0:24
    value = Decimal(value).quantize(0, ROUND_HALF_UP)  # normal rounding
    if value > deviation:
        stim = index

if stim != 0:  # for case when needed threshold achieved after first stimulation
    answer.append(dictionary_stim[stim + 1])
else:
    answer.append(dictionary_stim[stim])

general_table = [[deviation_list], [answer]]

# data writing

new_book = xlrd.open_workbook(new_file_path)
r_sheet = new_book.sheet_by_index(0)  # read only copy to introspect the file
wb = copy(new_book)  # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0)  # the sheet to write to within the writable copy

# ADDITIONAL writing of value within lists
# w_sheet.write(row+1, 0, deviation)
# w_sheet.write(row+1, 1, dictionary_stim[stim+1])

w_sheet.write(0, 0, 'Deviation')  # header 1
w_sheet.write(0, 1, 'Stimulation')  # header 2

for deviat in deviation_list:
    row = + 1
    w_sheet.write(row, 0, deviat)

for stimulation in answer:
    row = + 1
    w_sheet.write(row, 1, stimulation)

wb.save(new_file_path)

print(deviation_list, answer)

print(stim)


def name(x):
    print(x)
