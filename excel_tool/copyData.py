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
rebuildMap = {"phone_msg_option_info_data.xlsx": 0}

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
            with WorkBook(tarPath, wbName, app) as wb:
                printer.printGapTime("目标表格打开完毕，耗时:", skipLines=1, isCompare=True)
                dataDeal = DataDeal(wb.getData())
                dataDeal.setOriData(oriData)

                printer.setStartTime("开始同步数据:" + rowList[0])
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
                        else:
                            dataDeal.copyData(cell, checkCol)
                printer.printGapTime("数据同步完毕，耗时:", skipLines=1)

                rebuildCol = rebuildMap.get(wbName)
                if rebuildCol != None:
                    printer.setStartTime("开始重新生成连续的uid:" + rowList[0])
                    dataDeal.uidRebuild(rebuildCol)
                    printer.printGapTime("uid生成完毕，耗时:", skipLines=1)

                printer.setStartTime("开始检查单元格格式:" + rowList[0])
                dataDeal.strCheck()
                printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

                wb.setData(dataDeal.getData())
                wb.save()

if batName == None:
    Caller.callBat(tarPath, "1.一键处理表格.bat")
else:
    Caller.callBat(tarPath, batName)
