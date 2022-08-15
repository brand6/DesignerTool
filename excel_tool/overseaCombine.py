from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
import numpy

toStr = DataDeal.toStr
printer = Printer()
combineWb = 'card_combine.xlsx'
infoWb = 'card_info_data.xlsx'
tarPath = WorkBook.getActivePath()

with WorkApp() as app:
    cardlist = []
    printer.setStartTime("开始读取上线卡牌列表:")
    with WorkBook(tarPath, infoWb, app) as wb:
        dataDeal = DataDeal(wb.getData())
        idCol = dataDeal.getColOrder('card_id')
        rareCol = dataDeal.getColOrder('card_rare')
        skipCol = dataDeal.getColOrder('skip')

        for row in dataDeal.rngData:
            if row[rareCol] == 7:
                if skipCol != -1:
                    if row[skipCol] != 1:
                        cardlist.append(row[idCol])
                else:
                    cardlist.extend(row[idCol])
    printer.printGapTime("卡牌数据读取完毕，耗时:", skipLines=1)

    printer.setStartTime("开始修改表格:" + combineWb)
    with WorkBook(tarPath, combineWb, app) as wb:
        dataDeal = DataDeal(wb.getData())
        skipCol = dataDeal.getColOrder('skip')
        card1Col = dataDeal.getColOrder('card_id1')
        card2Col = dataDeal.getColOrder('card_id2')

        if skipCol == -1:
            skipCol = wb.sheet['A1'].end('right').column
            numpy.insert(dataDeal.rngData, skipCol, values=None, axis=1)
            dataDeal.rngData[0][skipCol] = 'skip'

        for row in dataDeal.rngData[2:]:
            if row[card1Col] in cardlist and row[card2Col] in cardlist:
                continue
            elif DataDeal.isNumber(row[0]):
                row[skipCol] = 1
        printer.printGapTime("数据生成完毕，耗时:", skipLines=1)

        printer.setStartTime("开始检查单元格格式:" + combineWb)
        dataDeal.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(dataDeal.getData())
        wb.save()
