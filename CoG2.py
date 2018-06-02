from os import listdir
from os.path import isfile, join
import xlrd, xlwt
from xlutils.copy import copy
import csv
import openpyxl

folder_name = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/for Kseniya/Subj 1.03 arc/Day 1/TMS'  # change to input('Path for a folder') + '/'
list_of_files = [f for f in listdir(folder_name) if isfile(join(folder_name, f))]

new_file = '/Users/kseniya/Desktop/MY/Lab/TMS map subj/test/table.xlsx'  # change to new_file = input('File to save path')

list1 = ['CoG_Sm']
list2 = ['CoG_Ab']
list3 = ['CoG_P']
list4 = ['File_name']

for file_name in list_of_files: 
    try:
        file_name_full = folder_name + file_name
        rb = xlrd.open_workbook(file_name_full)
        sheet = rb.sheet_by_index(3)
        val1 = sheet.row_values(9)[2]
        val2 = sheet.row_values(10)[2]
        val3 = sheet.row_values(11)[2]
        list1.append(val1)
        list2.append(val2)
        list3.append(val3)
        list4.append(file_name)
    except Exception as error:
        print(file_name)

mane_list = []

for index in range(len(list1)):
    a = list1[index]
    b = list2[index]
    c = list3[index]
    d = list4[index]

    a_replace_chars = a.strip('()')
    a_replace_space = ''.join(a_replace_chars.split())
    a_split = a_replace_space.split(';')
    mane_list.append(a_split)

    b_replace_chars = b.strip('()')
    b_replace_space = ''.join(b_replace_chars.split())
    b_split = b_replace_space.split(';')
    mane_list.append(b_split)

    c_replace_chars = c.strip('()')
    c_replace_space = ''.join(c_replace_chars.split())
    c_split = c_replace_space.split(';')
    mane_list.append(c_split)

    d_split = d.split()
    mane_list.append(d_split)

rb = xlrd.open_workbook(new_file)
r_sheet = rb.sheet_by_index(0)  # read only copy to introspect the file
wb = copy(rb)  # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0)  # the sheet to write to within the writable copy

for index, val_list in enumerate(mane_list):
    for index2, column in enumerate(val_list):
        w_sheet.write(index, index2, column)

print(mane_list)

wb.save(new_file)
