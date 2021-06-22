import openpyxl

workbook=openpyxl.load_workbook("E:\Employees.xlsx")
sheet=workbook['EmployeeData']

#getting value of a cell
print(sheet['B7'].value)

#getting value of a cell
print(sheet.cell(row=6, column=2).value)

cell=sheet['B9']

print(cell.row,cell.column)

print(cell.coordinate)

#datatype of cell data
print(cell.data_type)

#printing the encoding
print(cell.encoding)


# changing the value of a cell

cell=sheet['B2']
cell.value='Anuj'
workbook.save("E:\Employees.xlsx")

print(cell.parent)

