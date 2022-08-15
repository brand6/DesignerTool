from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
from myClass.caller import Caller

printer = Printer()
dataSht = WorkBook.getActiveSht()
dataList = dataSht['D2:P2'].expand('down').value
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
            wbName = rowList[0] + ".unpro.xlsx"
            printer.setStartTime("开始打开目标表格:" + rowList[0], 'green')
            with WorkBook(tarPath, wbName, app) as wb:
                printer.printGapTime("目标表格打开完毕，耗时:", skipLines=1, isCompare=True)
                updateTag = 0
                dataDeal = DataDeal(wb.getData())
                printer.setStartTime("开始隐藏[" + wbName + "]数据:")

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
                                updateTag += dataDeal.updateData(value, checkCol, updateColOrder, updateStr)
                        else:
                            updateTag += dataDeal.updateData(cell, checkCol, updateColOrder, updateStr)
                printer.printGapTime(wbName + "数据隐藏完毕，耗时:", skipLines=1)

                printer.setStartTime("开始检查单元格格式:")
                dataDeal.strCheck()
                printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

                wb.setData(dataDeal.getData())
                if updateTag > 0:
                    wb.save()