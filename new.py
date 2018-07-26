# -*- coding: UTF-8 -*-
from xlrd import open_workbook
from xlwt import Workbook
from getpass import getuser
from os import listdir,mkdir,path
from shutil import copy,rmtree


class detailMerge:
    def __init__(self, fileNames, detailPath, rootPath):
        self.fileNames = fileNames
        self.detailPath = detailPath
        self.rootPath = rootPath
        self.tables = []
        self.tableHead = self.__tableHead__()
        self.__merge__()

    def __merge__(self):
        for name in self.fileNames:
            path = self.detailPath + name
            sheet = open_workbook(path).sheet_by_index(0)
            [table, projectID] = self.__exchange__(sheet)
            self.tables = self.tables + table
            newPath = self.rootPath + 'detail_new/' + projectID + '_detail.xls'
            copy(path, newPath)

    def save(self):
        file = Workbook()
        table = file.add_sheet('detail', cell_overwrite_ok=True)
        tables = self.tableHead + self.tables
        for i in range(len(tables)):
            for j in range(len(tables[i])):
                try:
                    table.write(i, j, tables[i][j])
                except IndexError:
                    print(i)
                    print(j)
        file.save(self.rootPath + 'detail.xls')

    def __tableHead__(self):
        path = self.detailPath + self.fileNames[0]
        sheet = open_workbook(path).sheet_by_index(0)
        tableHead = [[]]
        cols = sheet.ncols
        for i in range(cols):
            tableHead[0].append(sheet.cell(1, i).value)
        tableHead[0].append(u"项目代码")
        tableHead[0].append(u"分项代码")
        return tableHead

    def __exchange__(self, sheet):
        cols = sheet.ncols
        rows = sheet.nrows
        sheetTitle = sheet.cell(0, 0).value
        projectID = sheetTitle[5:17]
        table = [["" for i in range(cols + 2)] for j in range(rows - 2)]
        for i in range(rows - 2):
            for j in range(cols):
                table[i][j] = sheet.cell(i + 2, j).value
            table[i][cols] = projectID
            table[i][cols + 1] = table[i][6].split('-')[0]
        return [table, projectID]


class budgetMerge:
    def __init__(self, fileNames, detailPath, rootPath):
        self.fileNames = fileNames
        self.detailPath = detailPath
        self.rootPath = rootPath
        self.tables = []
        self.tableHead = self.__tableHead__()
        self.__merge__()

    def __merge__(self):
        for name in self.fileNames:
            path = self.detailPath + name
            sheet = open_workbook(path).sheet_by_index(0)
            [table, projectID] = self.__exchange__(sheet)
            self.tables = self.tables + table
            newPath = self.rootPath + 'budget_new/' + projectID + '_budget.xls'
            copy(path, newPath)

    def save(self):
        file = Workbook()
        table = file.add_sheet('budget', cell_overwrite_ok=True)
        tables = self.tableHead + self.tables
        for i in range(len(tables)):
            for j in range(len(tables[i])):
                try:
                    table.write(i, j, tables[i][j])
                except IndexError:
                    print(i)
                    print(j)
        file.save(self.rootPath + 'budget.xls')

    def __tableHead__(self):
        path = self.detailPath + self.fileNames[0]
        sheet = open_workbook(path).sheet_by_index(0)
        tableHead = [[]]
        cols = sheet.ncols
        for i in range(cols):
            tableHead[0].append(sheet.cell(1, i).value)
        tableHead[0].append(u"项目代码")
        return tableHead

    def __exchange__(self, sheet):
        cols = sheet.ncols
        rows = sheet.nrows
        sheetTitle = sheet.cell(0, 0).value
        projectID = sheetTitle[5:17]
        table = [["" for i in range(cols + 1)] for j in range(rows - 3)]
        for i in range(rows - 3):
            for j in range(cols):
                table[i][j] = sheet.cell(i + 2, j).value
            table[i][cols] = projectID
        return [table, projectID]


class finalMerge:
    def __init__(self, detailTable, detailHead, budgetTable, budgetHead, rootPath):
        self.table = []
        self.tableHead = [detailHead[0] + budgetHead[0]]
        self.rootPath = rootPath
        # print(self.table)
        self.__merge__(detailTable, budgetTable)

    def __merge__(self, detailTable, budgetTable):
        indexOccupied = []
        for dt in detailTable:
            detailProjectID = dt[7]
            detailItemID = dt[8]
            flag = False
            for i in range(len(budgetTable)):
                bt = budgetTable[i]
                budgetProjectID = bt[7]
                budgetItemID = bt[0]
                if detailProjectID == budgetProjectID and detailItemID == budgetItemID:
                    self.table.append(dt + bt)
                    indexOccupied.append(i)
                    flag = True
                    break
            if not flag:
                self.table.append(dt + ["-"] * (len(bt) - 1) + [detailProjectID])
        for i in range(len(budgetTable)):
            if i not in indexOccupied:
                self.table.append(["-"] * len(dt) + budgetTable[i])

    def save(self):
        file = Workbook()
        table = file.add_sheet('final', cell_overwrite_ok=True)
        tables = self.tableHead + self.table
        # print(tables)
        for i in range(len(tables)):
            for j in range(len(tables[i])):
                try:
                    table.write(i, j, tables[i][j])
                except IndexError:
                    print(i)
                    print(j)
        file.save(self.rootPath + 'final.xls')


if __name__ == "__main__":
    user = getuser()
    rootPath = 'C:/Users/' + user + '/Desktop/table/'
    detailPath = rootPath + 'detail/'
    budgetPath = rootPath + 'budget/'

    if path.exists(rootPath+'budget_new/'):
        rmtree(rootPath+'budget_new/')
    if path.exists(rootPath + 'detail_new/'):
        rmtree(rootPath+'detail_new/')
    mkdir(rootPath+'budget_new/')
    mkdir(rootPath+'detail_new/')
    detailFileNames = listdir(detailPath)
    budgetFileNames = listdir(budgetPath)
    dm = detailMerge(detailFileNames, detailPath, rootPath)
    dm.save()

    bm = budgetMerge(budgetFileNames, budgetPath, rootPath)
    bm.save()

    fm = finalMerge(dm.tables, dm.tableHead, bm.tables, bm.tableHead, rootPath)
    fm.save()
