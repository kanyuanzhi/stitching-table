# -*- coding: UTF-8 -*-
from xlrd import open_workbook
from xlwt import Workbook
from getpass import getuser
from os import listdir

user = getuser()
tablePath = 'C:/Users/' + user + '/Desktop/two_tables/'
# tablePath = 'two_tables/'
fileNames = listdir(tablePath)

nameToTable = []
nameToMatrix = []

nameToTable.append(open_workbook(tablePath + u'明细.xls').sheet_by_index(0))
nameToTable.append(open_workbook(tablePath + u'预算.xls').sheet_by_index(0))

for m in range(2):
    cols_count = nameToTable[m].ncols
    rows_count = nameToTable[m].nrows
    nameToMatrix.append([["" for i in range(cols_count)] for j in range(rows_count)])
    for i in range(rows_count):
        for j in range(cols_count):
            nameToMatrix[m][i][j] = nameToTable[m].cell(i, j).value

for row_index, line in enumerate(nameToMatrix[0]):
    flag1 = False
    for col_index, item in enumerate(line):
        if item == u"分项代码":
            col_index_mx = col_index
            row_index_mx = row_index
            flag1 = True
            break
    flag2 = False
    for col_index, item in enumerate(line):
        if item == u"项目代码":
            col_index_pid_mx = col_index
            flag2 = True
            break
    if flag1 and flag2:
        break

for row_index, line in enumerate(nameToMatrix[1]):
    flag1 = False
    for col_index, item in enumerate(line):
        if item == u"分项代码":
            col_index_ys = col_index
            row_index_ys = row_index
            flag1 = True
            break
    flag2 = False
    for col_index, item in enumerate(line):
        if item == u"项目代码":
            col_index_pid_ys = col_index
            flag2 = True
            break
    if flag1 and flag2:
        break

matrix_mx = nameToMatrix[0][row_index_mx:]
matrix_ys = nameToMatrix[1][row_index_ys:]

used_flag = []
for i in range(1, len(matrix_mx)):
    for j in range(1, len(matrix_ys)):
        if matrix_ys[j][col_index_ys] == matrix_mx[i][col_index_mx] and \
                        matrix_ys[j][col_index_pid_ys] == matrix_mx[i][col_index_pid_mx]:
            matrix_mx[i] = matrix_mx[i] + matrix_ys[j]
            used_flag.append(j)

empty = ["-"] * len(matrix_mx[0])
for j in range(1, len(matrix_ys)):
    if j not in used_flag:
        matrix_mx.append(empty + matrix_ys[j])

matrix_mx[0] = matrix_mx[0] + matrix_ys[0]

file = Workbook()
table = file.add_sheet('ceshi', cell_overwrite_ok=True)
for i in range(len(matrix_mx)):
    for j in range(len(matrix_mx[i])):
        try:
            table.write(i, j, matrix_mx[i][j])
        except IndexError:
            print(i)
            print(j)
file.save('C:/Users/' + user + '/Desktop/merge_original.xls')
