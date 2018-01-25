import xlrd
import xlwt
import os

xlrd.Book.encoding = "utf8"
test = xlrd.open_workbook('test.xlsx')
table = test.sheet_by_index(0)

print table.nrows
print table.ncols
print test.nsheets

fileNames = os.listdir('./')
tableNames = []

for name in fileNames:
    #if name.endswith('.xls') or name.endswith('.xlsx'):
    if name.endswith('.xls'):
        tableNames.append(name)

print fileNames
print tableNames

nameToTable = dict.fromkeys(tableNames)

for name in tableNames:
    nameToTable[name] = xlrd.open_workbook(name).sheet_by_index(0)

print nameToTable

print nameToTable['CX2100060032.xls'].row_values(0)
