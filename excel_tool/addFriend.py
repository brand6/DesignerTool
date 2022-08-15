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

infoWb = 'friend_msg_info_data.xlsx'
textWb = 'friend_msg_text.xlsx'
infoList = dataSht['A1'].expand('table').value
textList = dataSht['U1'].expand('table').value

# 获取短电朋参数
favorData = WorkBook.getRangeByShtName('短电朋数据', 'E1', 'table').value
weekTimeData = WorkBook.getRangeByShtName('短电朋数据', 'A1', 'table').value
roleMap = {1: '1', 2: '2', 3: '3', 4: '4', 33: '3', 34: '8'}

with WorkApp() as app:
    # -----------------------------------friend_msg_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + infoWb, 'green')
    with WorkBook(tarPath, infoWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        infoData = MsgDeal(wb.getData())
        infoData.setFavorData(favorData)
        infoData.setWeekTimeData(weekTimeData)
        for row in infoList[1:]:
            reward = infoData.getFavorValue(row[7], 2)
            saleTime = ''
            if isEmpty(row[12]) == False:
                saleTime = row[12]
            elif isEmpty(row[13]) == False:
                saleTime = infoData.getSaleTime(infoData.getWeekTime(row[13], 2))
            dataMap = {
                'friend_msg_id': row[6],
                'friend_msg_title': row[1],
                'msg_loop_cnt': 1,
                'friend_msg_type': 0 if row[2] == 0 else 1,
                'friend_msg_need': roleMap[row[2]] + ',54,' + toStr(row[8]) if isEmpty(row[8]) == False and row[2] != 0 else '',
                'friend_msg_npc': row[2],
                'msg_text': row[6] * 100,
                'msg_option_ids': row[3],
                'friend_msg_option_reply': row[4],
                'extra_reply': row[5],
                'friend_msg_background': row[6] if isEmpty(row[9]) == False else '',
                'friend_msg_rewards': row[11],
                'msg_option_rewards': roleMap[row[2]] + ',46,' + toStr(reward) if isEmpty(row[7]) ==False and row[2] != 0 else '',
                'text': row[10],
                'drop_year': infoData.getYear(),
                'background': infoData.getMsgBackGround(row[7]),
                'friend_msg_price': '119195,101,2;0,2,80' if isEmpty(row[12]) == False or isEmpty(row[13]) == False else '',
                'friend_msg_sale_time': saleTime
            } #yapf:disable
            infoData.setRowData(row[6], dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + infoWb)
        infoData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(infoData.getData())
        wb.save()

    # -----------------------------------friend_msg_text-------------------------------------------------
    printer.setStartTime("开始打开表格:" + textWb, 'green')
    with WorkBook(tarPath, textWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        textData = MsgDeal(wb.getData())
        for row in textList[1:]:
            dataMap = {
                'friend_msg_text_id': row[0],
                'reply_npc': row[3],
                'friend_msg_text': row[4],
            }
            descMap = {
                '相关id[hide]': row[1],
                '备注[hide]': row[2]
            } #yapf:disable
            textData.setRowData(row[0], dataMap)
            textData.setRowData(row[0], descMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + textWb)
        textData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(textData.getData())
        wb.save()