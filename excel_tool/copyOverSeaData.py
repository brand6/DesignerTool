from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
from myClass.caller import Caller

printer = Printer()
dataSht = WorkBook.getActiveSht()
dataList = dataSht['D2:P2'].expand('down').value
oriPath = dataSht['A2'].value
tarPath = dataSht['A5'].value
batName = dataSht['A12'].value
updateCol = dataSht['A14'].value
updateStr = dataSht['A15'].value

with WorkApp() as app:
    for rowList in dataList:
        hasData = False
        for cell in rowList[3:]:
            if cell != None:
                hasData = True
                break
        if hasData:
            wbName = rowList[0] + ".xlsx"
            oriData = None
            printer.setStartTime("开始打开来源表格:" + rowList[0], 'green')
            with WorkBook(oriPath, wbName, app) as wb:
                printer.setCompareTime(printer.printGapTime("来源表格打开完毕，耗时:"))
                oriData = wb.getData()

            printer.setStartTime("开始打开目标表格:" + rowList[0])

            with WorkBook(tarPath, rowList[0] + '.unpro.xlsx', app) as wb:
                printer.printGapTime("目标表格打开完毕，耗时:", skipLines=1, isCompare=True)
                dataDeal = DataDeal(wb.getData())
                dataDeal.setOriData(oriData)

                printer.setStartTime("开始同步数据:" + rowList[0])
                updateColOrder = dataDeal.getColOrder(updateCol)
                if updateColOrder == -1:
                    updateColOrder = dataDeal.addCol(updateCol)

                if '|' in str(rowList[1]):
                    checkCol = rowList[1].rsplit('|')
                else:
                    checkCol = rowList[1]

                for cell in rowList[3:]:
                    if cell != None:
                        if '|' in str(cell):
                            valueList = cell.rsplit('|')
                            for value in valueList:
                                dataDeal.copyData(value, checkCol)
                                dataDeal.updateData(value, checkCol, updateColOrder, updateStr)
                        else:
                            dataDeal.copyData(cell, checkCol)
                            dataDeal.updateData(cell, checkCol, updateColOrder, updateStr)
                printer.printGapTime("数据同步完毕，耗时:", skipLines=1)

                printer.setStartTime("开始检查单元格格式:" + rowList[0])
                dataDeal.strCheck()
                printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

                wb.setData(dataDeal.getData())
                wb.save()
printer.printColor("数据处理完毕：", 'yellow')