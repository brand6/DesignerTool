from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
from myClass.caller import Caller

printer = Printer()

xlToolWb = WorkBook.getXlToolWb()
dataSht = WorkBook.getActiveSht(xlToolWb)
oriPath = dataSht['A5'].expand('down').value
if type(oriPath) == str:
    oriPath = [oriPath]
tarPath = dataSht['A2'].value
batName = dataSht['A12'].value
configList = WorkBook.getRangeByShtName("捞文本配置", "A2:L2", "down", xlToolWb).value

with WorkApp() as app:
    printer.setStartTime("开始打开翻译表格:", 'green')
    transList = []
    for tansFile in oriPath:
        transList.append(WorkBook(tansFile, "", app))
    printer.setCompareTime(printer.printGapTime("翻译表格打开完毕，耗时:", skipLines=1))

    for rowList in configList:
        hasData = False
        for tansFile in transList:
            if hasData == False:
                for sht in tansFile.sheets:
                    if sht.name == rowList[0]:
                        hasData = True
                        break

        if hasData == True:
            wbName = rowList[0] + ".unpro.xlsx"
            oriData = None

            printer.setStartTime("开始打开目标表格:" + rowList[0], 'green')
            with WorkBook(tarPath, wbName, app) as wb:
                printer.printGapTime("目标表格打开完毕，耗时:", skipLines=1)
                dataDeal = DataDeal(wb.getData())

                printer.setStartTime("开始同步数据:" + rowList[0])
                for tansFile in transList:
                    try:
                        oriData = tansFile.getRangeByShtName(rowList[0], wb=tansFile).value
                        dataDeal.setOriData(oriData)
                        dataDeal.setTransColumn(rowList[1:])
                    except:
                        printer.printColor("无目标表格翻译页签：" + tansFile.name, 'yellow')
                        continue

                    if '|' in str(rowList[11]):
                        checkCol = rowList[11].rsplit('|')
                    else:
                        checkCol = rowList[11]

                    idCol = rowList[1]
                    idColOrder = dataDeal.getColOrder(idCol, dataType=1)
                    for row in oriData[1:]:
                        if row[idColOrder] != dataDeal.transId and row[idColOrder] != None:
                            dataDeal.transData(row[idColOrder], checkCol)
                printer.printGapTime("数据同步完毕，耗时:", skipLines=1)

                printer.setStartTime("开始检查单元格格式:" + rowList[0])
                dataDeal.strCheck()
                printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

                wb.setData(dataDeal.getData())
                wb.save()

    for tansFile in transList:
        tansFile.close()

printer.printColor("数据处理完毕：", 'yellow')
