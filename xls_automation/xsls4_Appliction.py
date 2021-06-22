
from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.worksheet.table import Table, TableStyleInfo
text_file=open("C:\\Users\\LENOVO_PC\\Desktop\\New folder\\xls_automation\\employees.txt")
records=[]
print(text_file.seek(0))
for record in text_file.readlines():
	records.append(record.rstrip("\n").split(";"))
print(records)
workbook=Workbook()
file_path="C:\\Users\\LENOVO_PC\\Desktop\\New folder\\xls_automation\\MyEmployees.xlsx"
workbook.save(file_path)
sheet=workbook['Sheet']
sheet.title='Employee'
for row in records:
	sheet.append(row)
table=Table(displayName="Table",ref='A1:G11')

style=TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=True)

table.tablestyleInfo=style

sheet.add_table(table)
font=Font(color=colors.BLUE,bold=True,italic=True)

for cell_no in range(2,12):
	if int(sheet['G%s' %(cell_no)].value)>55000:
		sheet['G%s' %(cell_no)].font=font

workbook.save(file_path)
text_file.close()