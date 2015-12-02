"""
Need to use xlrd and xlwt to read and write Microsoft Excel Files
And xlutils for some utils


"""

import xlrd, xlwt
import xlutils 
import csv

#opens the workbook
print("Opening Workbook")
wb = xlrd.open_workbook("ag_hr_1998")

#opens the worksheet: can do this by index or name 
print("Opening Worksheets")
icahr = wb.sheet_by_name('ICA HR') # might need to escape space
etaw = wb.sheet_by_name('ETAW HR')

#data cell's values are accessible by sheet.cell(i,j).value
print(icahr.cell(0,0).value) 
print(etaw.cell(0,0).value) 







"""
References: 
http://www.sitepoint.com/using-python-parse-spreadsheet-data/ 

"""