import openpyxl #library that permits you to run excel with python 
from openpyxl import workbook #library that permits python to import the workbook

#To open workbook
#workbook object is created i.e wb_obj
wb_obj = openpyxl.load_workbook("employeedata.xlsx")

#load worksheet ws
ws = wb_obj['Sheet1']
ws = wb_obj.active #creates an active worksheet


#change worksheet(ws) dimension
ws.column_dimensions['B'].width = 32

column = ws['B2':'B31'] # worksheet(ws) takes range (B2:B31) and assigns it to column

for range in column:
    for output in range:
        printable = output.value
        updated_email = printable.replace('helpinghands.cm', 'handsinhands.org')
        output.value = updated_email

#save file that has been edited in new excel and csv sheet respectively
wb_obj.save(filename="employeeupdateddata.xlsx")
wb_obj.save(filename="employeeupdateddata.csv")

