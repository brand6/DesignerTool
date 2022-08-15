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
tagWb = 'short_msg_chat_tag.xlsx'
feedbackWb = 'short_msg_feedback'
dataList = dataSht.used_range.value

# 获取短电朋参数
favorData = WorkBook.getRangeByShtName('短电朋数据', 'E1', 'table').value
weekTimeData = WorkBook.getRangeByShtName('短电朋数据', 'A1', 'table').value
roleMap = {'李泽言': '1', '许墨': '2', '周棋洛': '3', '白起': '4', '凌肖': '8'}
contactMap = {'李泽言': '1', '许墨': '2', '周棋洛': '3', '白起': '4', '凌肖': '34'}

with WorkApp() as app:
    # -----------------------------------short_msg_chat_tag-------------------------------------------------
    tagData = []
    printer.setStartTime("开始打开表格:" + tagWb, 'green')
    with WorkBook(tarPath, tagWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")
        tagData = wb.getData()

    # -----------------------------------short_msg_feedback-------------------------------------------------
    printer.setStartTime("开始打开表格:" + feedbackWb, 'green')
    with WorkBook(tarPath, feedbackWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        feedbackData = MsgDeal(wb.getData())
        feedbackData.setOriData(tagData)
        for row in dataList[1:]:
            tag1Id = toStr(feedbackData.getRowData(row[1], 0, 1, 1))
            tag3Id = toStr(feedbackData.getRowData(row[3], 0, 1, 1))
            feedbackId = roleMap[row[0]] + '4' + toStr(row[7])
            dataMap = {
                'feedback_id': feedbackId,
                'male': roleMap[row[0]],
                'feedback_type': 4,
                'feedback_msg_list': toStr(row[6]) + '|1',
                'feedback_condition_param': tag1Id + ':' + tag3Id,
            } #yapf:disable
            descMap = {'反馈库名称[hide]': '主动聊天:' + row[1] + '-' + row[2] + '-' + row[3],
                       '反馈库编号[hide]': row[7],
            }# yapf:disable
            feedbackData.setRowData(feedbackId, dataMap)
            feedbackData.setRowData(feedbackId, descMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + feedbackWb)
        feedbackData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(feedbackData.getData())
        wb.save()

    # -----------------------------------short_msg_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + infoWb, 'green')
    with WorkBook(tarPath, infoWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        infoData = MsgDeal(wb.getData())
        for row in dataList[1:]:
            dataMap = {
                'short_msg_id': row[6],
                'short_msg_type': 0,
                'msg_loop_cnt': 1,
                'favor_level_need': '',
                'phone_contact_id': contactMap[row[0]],
                'firstmsg_ids': '',
                'msg_option_ids': toStr(row[6]) + '110',
                'msg_visible_type': 0,
                'background': 1,
                'msg_list_hide': 1,
                'short_msg_rewards':'',
                'name': row[0] + row[1] + '-' +  row[3],
                #'priority': '',
                #'drop_year': infoData.getYear(),
                #'short_msg_price': '119195,101,3;0,2,120' if isEmpty(row[14]) == False else '',
                #'short_msg_sale_time': 'saleTime',
            }# yapf:disable
            infoData.setRowData(row[6], dataMap)
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
        for row in dataList[1:]:
            optionId = row[6] * 1000 + 110
            dataMap = {
                'short_msg_id': row[6],
                'msg_option_id': optionId,
                'replies': optionData.getReplyListStr(optionId,row[5]),
                'follow_option_ids': '-1',
            } #yapf:disable
            optionData.setRowData(optionId, dataMap, 1)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + optionWb)
        optionData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(optionData.getData())
        wb.save()

    # -----------------------------------short_msg_text-------------------------------------------------
    printer.setStartTime("开始打开表格:" + textWb, 'green')
    with WorkBook(tarPath, textWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        textData = MsgDeal(wb.getData())
        for row in dataList[1:]:
            optionId = row[6] * 1000 + 110
            dataMap = {
                'id': optionId,
                #'title': '',
                'content': row[4],
                'text_time': 0,
                #'send_data': row[6],
                #'send_time': row[7],
                #'reward': row[8],
                #'send_data_typ': row[9],
            } #yapf:disable
            descMap = {
                '相关项[hide]': row[6],
                '备注[hide]': '主动聊天-女主开场',
            }# yapf:disable
            textData.setRowData(optionId, dataMap)
            textData.setRowData(optionId, descMap, defRow=3)

            replyList = str.splitlines(row[5])
            for order in range(len(replyList)):
                dataMap = {
                    'id': optionId + order + 1,
                    #'title': row[3],
                    'content': replyList[order],
                    'text_time': len(replyList[order]) * 150,
                    #'send_data': row[6],
                    #'send_time': row[7],
                    #'reward': row[8],
                    #'send_data_typ': row[9],
                } #yapf:disable
                descMap = {
                    '相关项[hide]': row[6],
                    '备注[hide]': '主动聊天-男主反馈' + toStr(order + 1),
                }# yapf:disable
                textData.setRowData(optionId + order + 1, dataMap)
                textData.setRowData(optionId + order + 1, descMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + textWb)
        textData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(textData.getData())
        wb.save()