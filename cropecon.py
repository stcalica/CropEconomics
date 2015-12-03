"""
Need to use xlrd and xlwt to read and write Microsoft Excel Files
And xlutils for some utils


"""

import xlrd, xlwt
import xlutils 
import csv

#opens the workbook
print("Opening Workbook")
wb = xlrd.open_workbook("ag_hr_1998.xls")

#opens the worksheet: can do this by index or name 
print("Opening Worksheets")
icahr = wb.sheet_by_name('ICA HR') # might need to escape space
etaw = wb.sheet_by_name('ETAW HR')

#data cell's values are accessible by sheet.cell(row,column).value\
#should print out the first year in both (1998)
print(icahr.cell(1,0).value) 
print(etaw.cell(1,0).value) 

#we can now iterate through the list and make each row a list 


#then we make a dot product between the two lists





"""
References: 
http://www.sitepoint.com/using-python-parse-spreadsheet-data/ 
http://stackoverflow.com/questions/4093989/dot-product-in-python

"""