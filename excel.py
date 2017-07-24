import openpyxl
import os

#os.chdir("\\")

workbook = openpyxl.load_workbook('..//t.xlsx')

print type(workbook)

print workbook.get_sheet_names()


try:
	sheet = workbook.get_sheet_by_name('Sheet1')
	print sheet['C1'].value

except:
	print "Invalid Sheet name"

for i in range(1,8):
	print(i,sheet.cell(row=i,column=2).value)

