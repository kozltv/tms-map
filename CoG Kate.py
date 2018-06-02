from os import listdir
from os.path import isfile, join
import xlrd, xlwt
from xlutils.copy import copy

#File_with_CoG_path = input('Path to file: ')
File_with_CoG_path = '/Users/kseniya/Downloads/all_withSPSSICC_3.xlsm'
# new_file = input('Path to blank excel file for saving: ')

wb = xlrd.open_workbook(File_with_CoG_path)
sheet = wb.sheet_by_index(6)

list1 = []
list2 = []
list3 = []
list4 = []
list5 = []
list6 = []
list7 = []
list8 = []
list9 = []
list10 = []
list11 = []
list12 = []





for column in range(4, 17):
    for n in range(2, 23):
        value = sheet.row_values(n)[column]
        list1.append(value)

print(list1)


# n = -1
#
# for line in range(sheet):
#     n += 1
#     print(sheet.row_values(2)[n])


# for line in xlrd.open_workbook(File_with_CoG_path):





# for index in range(len(list1)):
#     a = list1[index]
#     b = list2[index]
#     c = list3[index]
#     d = list4[index]
#
#     a_replace_chars = a.strip('()')
#     a_replace_space = ''.join(a_replace_chars.split())
#     a_split = a_replace_space.split(';')
#
#     b_replace_chars = b.strip('()')
#     b_replace_space = ''.join(b_replace_chars.split())
#     b_split = b_replace_space.split(';')
#
#     c_replace_chars = c.strip('()')
#     c_replace_space = ''.join(c_replace_chars.split())
#     c_split = c_replace_space.split(';')
#
#     d_split = d.split()
#
#     row_value_lists.append([a_split, b_split, c_split, d_split])
#     print(a_split, b_split, c_split, d_split)
