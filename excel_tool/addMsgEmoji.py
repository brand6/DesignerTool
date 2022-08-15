from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.msgDeal import MsgDeal
from myClass.printer import Printer
from myClass.caller import Caller

toStr = MsgDeal.toStr
isEmpty = MsgDeal.isEmpty
printer = Printer()
tarPath = WorkBook.getActivePath()
dataSht = WorkBook.getActiveSht()

textWb = 'short_msg_text.xlsx'
dataList = dataSht['E1'].expand('table').value

with WorkApp() as app:
    # -----------------------------------short_msg_text-------------------------------------------------
    printer.setStartTime("开始打开表格:" + textWb, 'green')
    with WorkBook(tarPath, textWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        textData = MsgDeal(wb.getData())
        for row in dataList[1:]:
            dataMap = {
                'emoji': row[2],
            } #yapf:disable
            for r in textData.getData():
                if row[0] == r[1] and row[1] in r[4]:
                    textData.setRowData(r[0], dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + textWb)
        textData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(textData.getData())
        wb.save()
