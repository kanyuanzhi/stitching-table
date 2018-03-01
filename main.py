# -*- coding: UTF-8 -*-
import xlrd
import xlwt
import getpass
import os

user = getpass.getuser()
# tablePath = 'C:/Users/' + user + '/Desktop/tables/'
tablePath = 'tables/'
fileNames = os.listdir(tablePath)
tableNames = []

for name in fileNames:
    if name.endswith('.xls') or name.endswith('.xlsx'):
        # if name.endswith('.xls'):
        tableNames.append(name)

nameToTable = dict.fromkeys(tableNames)
nameToMatrixProcessed = dict.fromkeys(
    tableNames)  # 处理后表格，添加项目代码与项目名称字段，删除前两行表头

for name in tableNames:
    nameToTable[name] = xlrd.open_workbook(tablePath + name).sheet_by_index(0)
    cols_count = nameToTable[name].ncols
    rows_count = nameToTable[name].nrows
    nameToMatrixProcessed[name] = [
        ["" for i in range(cols_count)] for j in range(rows_count - 2)]
    projectCode = nameToTable[name].cell(0, 1).value  # 项目代码
    projectName = nameToTable[name].cell(0, 3).value  # 项目名称
    for i in range(rows_count - 2):
        for j in range(cols_count):
            nameToMatrixProcessed[name][i][j] = nameToTable[name].cell(
                i + 2, j).value
        nameToMatrixProcessed[name][i].append(projectCode)
        nameToMatrixProcessed[name][i].append(projectName)
    if nameToMatrixProcessed[name][rows_count - 3][0] == u'合计':
        nameToMatrixProcessed[name].pop()

##################################################################
# 处理表头
tableHead = []
flag = False
for th in nameToTable[tableNames[0]].row_values(1):
    if th == u"摘要":
        summaryIndex = nameToTable[tableNames[0]].row_values(1).index(th)
        flag = True
    tableHead.append(th)
tableHead.append(u"项目代码")
tableHead.append(u"项目名称")
if flag:
    tableHead.append(u"姓名")
##################################################################

finalMatrix = [tableHead]
for value in nameToMatrixProcessed.itervalues():
    if flag:
        for item in value:
            item.append(item[summaryIndex][0:3])
    finalMatrix.extend(value)

file = xlwt.Workbook()
table = file.add_sheet('ceshi', cell_overwrite_ok=True)
for i in range(len(finalMatrix)):
    for j in range(len(finalMatrix[0])):
        table.write(i, j, finalMatrix[i][j])
file.save('result.xls')
