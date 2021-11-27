import glob
import os
import xlsxwriter

workbook = xlsxwriter.Workbook(r"C:\Users\elis\Desktop\file_list.xlsx")
worksheet = workbook.add_worksheet()

get_file_list = glob.iglob(r"C:\Users\*\Desktop\*", recursive=True)
file_list = []

for file_path in get_file_list:
    file_list.append(file_path)

file_name = []
get_file_names = []

for file in file_list:
    get_file_names = list((os.path.splitext(os.path.basename(file))))

    if os.path.isdir(file):
        get_file_names[0] = get_file_names[0] + get_file_names[1]
        get_file_names[1] = 'Folder'

    file_name.append(get_file_names)

row = 0
col = 0

bold = workbook.add_format({'bold': True})
worksheet.write(row, col, "File name", bold)
worksheet.write(row, col + 1, "File extension", bold)
row += 1

for fn, fe in file_name:
    worksheet.write(row, col,     fn)
    worksheet.write(row, col + 1, fe)
    row += 1

workbook.close()












