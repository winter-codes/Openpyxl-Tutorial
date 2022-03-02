#Installation
pip install openpyxl
#can also try if the above doesn't work
pip3 install openpyxl

#Import openpyxl module
from openpyxl import Workbook
#Create a workbook
workbook = Workbook()
#Create a sheet
sheet = workbook.active
#Fill the sheet with data
sheet["A1"] = 'Numbers'
sheet["A2"] = 546
sheet['B4'] = 629
sheet['B1'].value = 'More Numbers'
sheet.cell(row = 4, column = 8).value = 'Hello World'
#Save the File
workbook.save(filename="sample.xlsx")