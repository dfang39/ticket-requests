import datetime
import os
import openpyxl

desktop = 'C:\\Users\\DanFang\\Desktop'
os.chdir(desktop)

workbook = openpyxl.load_workbook("penguins.xlsx", data_only = True, read_only = True)
sheet = workbook.get_sheet_by_name('Upcoming Games')
print(sheet)


result = '<?xml version="1.0" encoding="UTF-8"?><myItems>'

counter = 1

for row in range(1, sheet.rows):
    print(sheet['A'+str(counter)])
    counter += 1