from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
import os

printer = Printer()
toStr = DataDeal.toStr

printer.printColor('开始处理数据', 'green')
toolWb = WorkBook.getXlToolWb()
tarPath = WorkBook.getActivePath(toolWb)
oriPath = WorkBook.getActivePath(toolWb, 1)
if oriPath[-12:] != "excel_config":
    oriPath = oriPath + "\\Design\\excel_config"
dataSht = WorkBook.getActiveSht(toolWb)

if dataSht['A2'].value == None:
    os._exit()

flagList = []  # 操作的flag列表
for r in dataSht['A1'].expand('down')[1:]:
    flagList.append(r[0].value)

defineList = []  # 指定操作列表
for r in dataSht['B1'].expand('down')[1:]:
    name = r[0].value
    if name[-4:] != 'xlsx':
        name = name + '.xlsx'
    defineList.append(name)

blackList = []  # 黑名单列表
for r in dataSht['C1'].expand('down')[1:]:
    name = r[0].value
    if name[-4:] != 'xlsx':
        name = name + '.xlsx'
    blackList.append(name)

with WorkApp() as app:
    for name in os.listdir(oriPath):
        if name[-4:] == 'xlsx':
            opTag = False
            if len(defineList) == 0:
                if name not in blackList:
                    opTag = True
            elif name in defineList:
                opTag = True

            if opTag:
                dataRows = []
                copyMap = {}
                oriData = None
                with WorkBook(oriPath, name, app) as oriWb:
                    oriData = oriWb.getData()
                    flagCol = DataDeal.getDataColOrder(oriData, 'total_flag_id')
                    if flagCol == -1:
                        continue
                    else:
                        for row in range(len(oriData)):
                            if oriData[row][flagCol] in flagList:
                                dataRows.append(row)
                if len(dataRows) > 0:
                    printer.setStartTime("开始操作目标表格:" + name)
                    with WorkBook(tarPath, name, app) as wb:
                        dataDeal = DataDeal(wb.getData())
                        dataDeal.setOriData(oriData)
                        dataDeal.setRepeatLog()

                        checkCol = dataDeal.getMainCol(1)
                        if checkCol == None:
                            printer.printGapTime(name + "找不到比对列，耗时:", skipLines=1)
                            continue
                        colOrder = 0
                        if type(checkCol).__name__ == 'list':
                            colOrder = dataDeal.getColOrder(checkCol[0], dataType=1)
                        else:
                            colOrder = dataDeal.getColOrder(checkCol, dataType=1)

                        for row in dataRows:
                            copyId = oriData[row][colOrder]
                            if copyId not in copyMap:
                                copyMap[copyId] = 1
                                dataDeal.copyData(copyId, checkCol)

                        printer.printGapTime("数据同步完毕，耗时:", skipLines=1)

                        printer.setStartTime("开始检查单元格格式:" + name)
                        dataDeal.strCheck()
                        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

                        wb.setData(dataDeal.getData())
                        wb.save()