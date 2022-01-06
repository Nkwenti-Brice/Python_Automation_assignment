import openpyxl
from openpyxl import workbook

#To open workbook
#workbook object is created i.e wb_obj
wb_obj = openpyxl.load_workbook("employeedata.xlsx")

#load worksheet ws
ws = wb_obj['Sheet1']
ws = wb_obj.active

#change worksheet(ws) dimension
ws.column_dimensions['B'].width = 32

column = ws['B2':'B31']

for range in column:
    for output in range:
        printable = output.value
        updated_email = printable.replace('helpinghands.cm', 'handsinhands.org')
        output.value = updated_email

wb_obj.save(filename="employeeupdateddata.xlsx")
