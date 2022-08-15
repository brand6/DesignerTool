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

infoWb = 'phone_msg_info_data.xlsx'
optionWb = 'phone_msg_option_info_data.xlsx'
infoList = dataSht['A1'].expand('table').value
optionList = dataSht['Q1'].expand('table').value

# 获取短电朋参数
favorData = WorkBook.getRangeByShtName('短电朋数据', 'E1', 'table').value
weekTimeData = WorkBook.getRangeByShtName('短电朋数据', 'A1', 'table').value
roleMap = {1: '1', 2: '2', 3: '3', 4: '4', 33: '3', 34: '8'}
maleMap = {'李泽言': 1, '$u': '', '许墨': '2','周棋洛': '3','Helios': '3', '白起': '4', 'Ares': '2', '凌肖': '8'}

with WorkApp() as app:
    # -----------------------------------phone_msg_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + infoWb, 'green')
    with WorkBook(tarPath, infoWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        infoData = MsgDeal(wb.getData())
        infoData.setFavorData(favorData)
        infoData.setWeekTimeData(weekTimeData)
        for row in infoList[1:]:
            reward = infoData.getFavorValue(row[3], 1)
            dataMap = {
                'phone_msg_id': row[1],
                'msg_loop_cnt': 1,
                'favor_level_need': roleMap[row[4]] + ',54,' + toStr(row[5]) if isEmpty(row[5]) == False and row[4] != 0 else '',
                'phone_contact_id': row[4],
                'msg_option_ids': row[6],
                'phone_title': row[2],
                'phone_vo_id': 0 if row[7] == 1 else row[1],
                'phone_msg_rewards': roleMap[row[4]] + ',46,' + toStr(reward) if isEmpty(row[3]) ==False and row[2] != 0 else '',
                'phone_prompt': row[8],
                'is_video': row[7],
                'video_path': row[1] if row[7] == 1 else '',
                'drop_year': infoData.getYear(),
                'phone_price': '119195,101,10;0,2,400' if isEmpty(row[12]) == False or isEmpty(row[13]) == False else '',
                'phone_sale_time': row[9],
            } #yapf:disable
            infoData.setRowData(row[1], dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + infoWb)
        infoData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(infoData.getData())
        wb.save()

    # -----------------------------------phone_msg_option_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + optionWb, 'green')
    with WorkBook(tarPath, optionWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        optionData = MsgDeal(wb.getData())
        for row in optionList[1:]:
            dataMap = {
                'uid': row[0],
                'phone_msg_id': row[1],
                'msg_option_id': row[2],
                'next': row[3],
                'follow_option_ids': row[4],
                'msg_option_rewards': row[5],
                'renming': row[6],
                'dialog': row[7],
                'dhead': row[8],
                'demoji': row[9],
                'namebg': row[10],
                'title': row[11],
                'sentence': row[12],
                'voice': row[13],
                'male': maleMap[row[6]],
            }
            optionData.setRowData(row[1], dataMap, checkCol=1, extraId=row[2], extraCol=2)
        optionData.uidRebuild(0)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + optionWb)
        optionData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(optionData.getData())
        wb.save()