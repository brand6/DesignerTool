from re import split
import xlwings as xl
import os.path as path
from tkinter import filedialog


class WorkBook(xl.Book):
    def __init__(self, path, name, app, defSht=0):
        if name != "":
            if name[-5:] != ".xlsx":
                name = name + ".xlsx"
            else:
                name = name
            if path[-12:] != "excel_config":
                path = path + "\\Design\\excel_config"

            WorkBook.closeByName(name)
            impl = app.books.open(path + "\\" + name, 0)
        else:
            impl = app.books.open(path, 0)
        self.impl = impl

        if defSht == -1:
            self.sheet = xl.sheets.active
        else:
            self.sheet = self.sheets[defSht]

        self.range = self.sheet.range
        self.cells = self.sheet.cells

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()

    # 删除行
    def deleteRow(self, firstRow, lastRow=None):
        if lastRow == None:
            lastRow = firstRow
        delRows = str(int(firstRow)) + ':' + str(int(lastRow))
        self.sheet[delRows].delete()

    # 删除列
    def deleteCol(self, firstCol, lastCol=None):
        if lastCol == None:
            lastCol = firstCol

        delCols = self.getColumnAddress(firstCol) + ':' + self.getColumnAddress(lastCol)
        self.sheet[delCols].delete()

    # 插入行
    def insertRow(self, firstRow, lastRow=None):
        if lastRow == None:
            lastRow = firstRow
        insertStr = str(int(firstRow)) + ':' + str(int(lastRow))
        self.sheet[insertStr].insert()

    # 插入列
    def insertCol(self, firstCol, lastCol=None):
        if lastCol == None:
            lastCol = firstCol
        insertStr = self.getColumnAddress(firstCol) + ':' + self.getColumnAddress(lastCol)
        self.sheet[insertStr].insert()

    # 指定操作的sheet
    def setSht(self, defSht):
        self.sheet = self.sheets[defSht]

    # 获取表格指定数据
    def getRangeData(self, rangeStr, expendStr=''):
        sht = self.sheet
        if expendStr == '':
            return sht.range(rangeStr)
        else:
            return sht.range(rangeStr).expand(expendStr)

    # 获取表格全部数据
    def getData(self):
        return self.sheet.used_range.value

    # 写入表格数据
    def setData(self, dataValue):
        if self.sheet.api.FilterMode == True:
            self.sheet.api.ShowAllData()
        self.sheet.range('A1').value = dataValue

    # 关闭workbook
    def close(self, isSave=False):
        if isSave == True:
            self.save()
        if self.impl != None:
            self.impl.close()

    # 获得列名
    def getColumnAddress(self, colNum):
        colAddress = self.cells(1, colNum).get_address(True, False)
        colName = colAddress.split('$')[0]
        return colName

    # 通过名字关闭workbook
    @classmethod
    def closeByName(cls, wbName):
        if wbName[-5:] != ".xlsx":
            wbName = wbName + ".xlsx"

        for app in xl.apps:
            for wb in app.books:
                if wb.name == wbName:
                    wb.close()

    # 获取表格指定数据
    @classmethod
    def getRangeByShtName(cls,shtName,rangeStr='',expendStr='',wb=''): #yapf:disable
        if wb == '':
            wb = cls.getActiveWb()
        sht = wb.sheets(shtName)
        if expendStr == '':
            if rangeStr != '':
                return sht.range(rangeStr)
            else:
                return sht.used_range
        else:
            return sht.range(rangeStr).expand(expendStr)

    # 获取当前激活的表格
    @classmethod
    def getActiveSht(cls, wb=None):
        if wb == None:
            return xl.sheets.active
        else:
            return wb.sheets.active

    # 获取当前激活的工作簿
    @classmethod
    def getActiveWb(cls, app=None):
        if app == None:
            return xl.books.active
        else:
            return app.books.active

    # 获得表格工具
    @classmethod
    def getXlToolWb(cls):
        for app in xl.apps:
            for wb in app.books:
                if wb.name[-4:] == 'xlsm':
                    return wb
        else:
            return None

    # 获取操作的配置表目录
    @classmethod
    def getActivePath(cls, wb=None, offset=0):
        if wb == None:
            wb = cls.getActiveWb()
        name = wb.name

        fullname = wb.fullname
        configPath = fullname[:len(fullname) - len(name)] + 'config.txt'
        if path.exists(configPath):
            with open(configPath, 'r', encoding='GBK') as file:
                user = file.read()
        else:
            with open(configPath, 'w', encoding='GBK') as file:
                user = input("请输入用户名：")
                file.write(user)
        sht = wb.sheets("目录")
        lastRow = sht.range("a999").end('up').row
        if lastRow % 2 == 1:
            lastRow += 1
        for i in range(lastRow):
            if user == sht.range('a' + str(i + 1)).value:
                opCell = sht.range('a' + str(i + 2)).offset(0, offset)
                projectPath = opCell.value
                if projectPath == None or projectPath == '':
                    opCell.value = filedialog.askdirectory().replace('/', '\\')
                    return opCell.value
                else:
                    return projectPath
        else:
            opCell = sht.range('a' + str(lastRow + 1))
            opCell.value = user
            opCell.offset(1, offset).value = filedialog.askdirectory().replace('/', '\\')
            return opCell.offset(1, offset).value
