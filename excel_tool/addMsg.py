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
infoWb = 'short_msg_info_data.xlsx'
optionWb = 'short_msg_option_info_data.xlsx'
dataList = dataSht.used_range.value

# 获取短电朋参数
favorData = WorkBook.getRangeByShtName('短电朋数据', 'E1', 'table').value
weekTimeData = WorkBook.getRangeByShtName('短电朋数据', 'A1', 'table').value
roleMap = {1: '1', 2: '2', 3: '3', 4: '4', 33: '3', 34: '8'}

with WorkApp() as app:
    # -----------------------------------short_msg_text-------------------------------------------------
    printer.setStartTime("开始打开表格:" + textWb, 'green')
    with WorkBook(tarPath, textWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        textData = MsgDeal(wb.getData())
        for row in dataList[2:]:
            dataMap = {
                'id': row[0],
                'title': row[3],
                'content': row[4],
                'text_time': row[5],
                'send_data': row[6],
                'send_time': row[7],
                'reward': row[8],
                'send_data_typ': row[9],
            } #yapf:disable
            descMap = {'相关项[hide]': row[1],
                       '备注[hide]': row[2],
            }# yapf:disable
            textData.setRowData(row[0], dataMap)
            textData.setRowData(row[0], descMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + textWb)
        textData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(textData.getData())
        wb.save()

    # -----------------------------------short_msg_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + infoWb, 'green')
    with WorkBook(tarPath, infoWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        infoData = MsgDeal(wb.getData())
        infoData.setFavorData(favorData)
        infoData.setWeekTimeData(weekTimeData)
        for row in dataList[2:]:
            if isEmpty(row[11]) == False:
                reward = infoData.getFavorValue(row[13], 3)
                saleTime = ''
                if isEmpty(row[10]) == False:
                    saleTime = row[10]
                elif isEmpty(row[14]) == False:
                    saleTime = infoData.getSaleTime(infoData.getWeekTime(row[14], 2))
                dataMap = {
                    'short_msg_id': row[1],
                    'short_msg_type': 0 if row[11] == 0 else 1,
                    'msg_loop_cnt': 1,
                    'favor_level_need': roleMap[row[11]] + ',54,' + toStr(row[12]) if isEmpty(row[12]) == False and row[11] != 0 else '',
                    'phone_contact_id': row[11],
                    'firstmsg_ids': MsgDeal.getFirstIds(dataList,row[1]) if row[11] != 0 else '',
                    'msg_option_ids': MsgDeal.getMsgOptions(dataList,row[1],1),
                    'msg_visible_type': 1,
                    'background': infoData.getMsgBackGround(row[13]),
                    # 'msg_list_hide': '',
                    'short_msg_rewards': roleMap[row[11]] + ',46,' + toStr(reward) if isEmpty(row[12]) ==False and row[11] != 0 else '',
                    'name': row[3],
                    # 'priority': '',
                    'drop_year': infoData.getYear(),
                    'short_msg_price': '119195,101,3;0,2,120' if isEmpty(row[14]) == False else '',
                    'short_msg_sale_time': saleTime,
                }# yapf:disable
                infoData.setRowData(row[1], dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + infoWb)
        infoData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(infoData.getData())
        wb.save()

    # -----------------------------------short_msg_option_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + optionWb, 'green')
    with WorkBook(tarPath, optionWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        optionData = MsgDeal(wb.getData())
        for row in dataList[2:]:
            optionId = toStr(row[0])
            if optionId[-1] == '0' and optionId[-3] != '0':
                dataMap = {
                    'short_msg_id': row[1],
                    'msg_option_id': row[0],
                    'replies': MsgDeal.getReplyIds(dataList, row[1], optionId[-3:-1]),
                    'follow_option_ids': MsgDeal.getMsgOptions(dataList,row[1],int(optionId[-3]) + 1),
                } #yapf:disable
                optionData.setRowData(row[0], dataMap, 1)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + optionWb)
        optionData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(optionData.getData())
        wb.save()