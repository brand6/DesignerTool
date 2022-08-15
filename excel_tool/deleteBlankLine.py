from myClass.dataDeal import DataDeal
from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
import os

checkSize = 1 * 1024  # 大于该大小的文件
whiteList = []  #yapf:disable

printer = Printer()
toStr = DataDeal.toStr
xlToolWb = WorkBook.getXlToolWb()
path = WorkBook.getActivePath(xlToolWb)

if path[-12:] != "excel_config":
    path = path + "\\Design\\excel_config"

printer.printColor("开始处理数据", 'green')
with WorkApp() as app:
    for name in os.listdir(path):
        fileStat = os.stat(path + '\\' + name)
        if name[-4:] == 'xlsx' and fileStat.st_size > checkSize and not name in whiteList:
            with WorkBook(path, name, app) as wb:
                saveTag = False
                cells = wb.cells
                lastCell = wb.sheet.used_range.last_cell
                maxCol = lastCell.column
                maxRow = lastCell.row

                firstCol = 0
                tempCol = 0
                for column in range(1, maxCol + 1):
                    col = column - tempCol
                    if cells(maxRow, col).value == None and cells(maxRow, col).end('up').value == None:
                        if firstCol == 0:
                            firstCol = col
                    elif firstCol != 0:
                        wb.deleteCol(firstCol, col - 1)
                        printer.setStartTime(name + "：删除空列" + toStr(col - firstCol))
                        saveTag = True
                        firstCol = 0
                        tempCol += col - firstCol
                else:
                    if firstCol != 0:
                        if col == firstCol:
                            col += 1
                        wb.deleteCol(firstCol, col - 1)
                        printer.setStartTime(name + "：删除空列" + toStr(col - firstCol))
                        saveTag = True

                if cells(maxRow, maxCol).value == None and cells(maxRow, maxCol).end('left').value == None:
                    firstRow = cells(maxRow, 1).end('up').row + 1
                    wb.deleteRow(firstRow, maxRow)
                    printer.setStartTime(name + "：删除空行" + toStr(maxRow - firstRow))
                    saveTag = True

                if saveTag:
                    wb.save()
                    printer.printGapTime("处理完毕，耗时:", skipLines=1)

printer.printColor("数据处理完毕", 'green')
