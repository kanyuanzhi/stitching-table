# -*- coding: UTF-8 -*-
import xlrd
import xlwt
import getpass
import os

user = getpass.getuser()
tablePath = 'C:/Users/' + user + '/Desktop/two_tables/'
# tablePath = 'two_tables/'
fileNames = os.listdir(tablePath)

nameToTable = []
nameToMatrix = []

nameToTable.append(xlrd.open_workbook(tablePath + 'mingxi.xls').sheet_by_index(0))
nameToTable.append(xlrd.open_workbook(tablePath + 'yusuan.xls').sheet_by_index(0))

for m in range(2):
    cols_count = nameToTable[m].ncols
    rows_count = nameToTable[m].nrows
    nameToMatrix.append([["" for i in range(cols_count)] for j in range(rows_count - 1)])
    for i in range(rows_count - 1):
        for j in range(cols_count):
            nameToMatrix[m][i][j] = nameToTable[m].cell(i + 1, j).value

for i in range(len(nameToMatrix[0])):
    projectCode0 = nameToMatrix[0][i][8]
    itemCode0 = nameToMatrix[0][i][0]
    for j in range(len(nameToMatrix[1])):
        projectCode1 = nameToMatrix[1][j][6]
        itemCode1 = nameToMatrix[1][j][0]
        if projectCode0 == projectCode1 and itemCode0 == itemCode1:
            nameToMatrix[0][i].extend(nameToMatrix[1][j])
            break

for i in range(len(nameToMatrix[0])):
    if len(nameToMatrix[0][i]) == 8:
        nameToMatrix[0][i].extend(["-", "-", "-", "-", "-", "-", "-", "-"])

finalTable = [[] for j in range(len(nameToMatrix[0]) + 1)]
finalTable[0].append(u"分项代码")
finalTable[0].append(u"分项名称")
finalTable[0].append(u"预算数")
finalTable[0].append(u"历年累计支出(不含借款)")
finalTable[0].append(u"余额")
finalTable[0].append(u"结余资金占预算比率")
finalTable[0].append(u"报销人员")
finalTable[0].append(u"报销内容")
finalTable[0].append(u"金额")
finalTable[0].append(u"报销金额占总支出比")

for i in range(len(nameToMatrix[0])):
    finalTable[i + 1].append(nameToMatrix[0][i][8])  # 分项代码
    finalTable[i + 1].append(nameToMatrix[0][i][1])  # 分项名称
    finalTable[i + 1].append(nameToMatrix[0][i][13])  # 预算数
    finalTable[i + 1].append(nameToMatrix[0][i][15])  # 历年累计支出(不含借款)
    finalTable[i + 1].append(nameToMatrix[0][i][16])  # 余额
    finalTable[i + 1].append('-')  # 结余资金占预算比率
    finalTable[i + 1].append(nameToMatrix[0][i][10])  # 报销人员
    finalTable[i + 1].append(nameToMatrix[0][i][4])  # 报销内容
    finalTable[i + 1].append(nameToMatrix[0][i][6])  # 金额
    finalTable[i + 1].append('-')  # 报销金额占总支出比
file = xlwt.Workbook()
table = file.add_sheet('ceshi', cell_overwrite_ok=True)
for i in range(len(finalTable)):
    for j in range(len(finalTable[0])):
        table.write(i, j, finalTable[i][j])
file.save('C:/Users/' + user + '/Desktop/merge.xls')
