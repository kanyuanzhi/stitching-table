# -*- coding: UTF-8 -*-
from xlrd import open_workbook
from xlwt import Workbook
from getpass import getuser
from os import listdir


def codeChange(itemCode, dictionary):
    if itemCode in dictionary:
        return dictionary[itemCode]
    else:
        return "-"


user = getuser()
tablePath = 'C:/Users/' + user + '/Desktop/three_tables/'
# tablePath = 'two_tables/'
fileNames = listdir(tablePath)

mingxiSheet = open_workbook(tablePath + u'预算.xls').sheet_by_index(0)
codeExchangeSheet = open_workbook(tablePath + u'分项代码转换表.xls').sheet_by_index(0)

mingxiTable = []
codeExchangeTable = []

cols_count = mingxiSheet.ncols
rows_count = mingxiSheet.nrows
mingxiTable = [["" for i in range(cols_count)] for j in range(rows_count)]
for i in range(rows_count):
    for j in range(cols_count):
        mingxiTable[i][j] = mingxiSheet.cell(i, j).value

cols_count = codeExchangeSheet.ncols
rows_count = codeExchangeSheet.nrows
codeExchangeTable = [["" for i in range(cols_count)] for j in range(rows_count)]
for i in range(rows_count):
    for j in range(cols_count):
        codeExchangeTable[i][j] = codeExchangeSheet.cell(i, j).value

codeChangeKey = []
codeChangeValue = []
for row in codeExchangeTable:
    if row[2] != '':
        codeChangeKey.append(row[1])
        codeChangeValue.append(row[2])
codeChangeDictionary = dict(zip(codeChangeKey, codeChangeValue))

mingxiTable[0].append(u'新分项代码')
for i in range(1, len(mingxiTable)):
    mingxiTable[i].append(codeChange(mingxiTable[i][1],codeChangeDictionary))

file = Workbook()
table = file.add_sheet('ceshi', cell_overwrite_ok=True)
for i in range(len(mingxiTable)):
    for j in range(len(mingxiTable[0])):
        table.write(i, j, mingxiTable[i][j])
file.save('C:/Users/' + user + '/Desktop/new_yusuan.xls')
