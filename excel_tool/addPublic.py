from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.msgDeal import MsgDeal
from myClass.dataDeal import DataDeal
from myClass.printer import Printer
from myClass.caller import Caller

toStr = MsgDeal.toStr
printer = Printer()
infoWb = 'public_msg_info_data.xlsx'
commentWb = 'public_msg_comment.xlsx'
tarPath = WorkBook.getActivePath()
dataSht = WorkBook.getActiveSht()
dataList = dataSht['A1'].expand('table').value

with WorkApp() as app:
    # -----------------------------------public_msg_info_data-------------------------------------------------
    printer.setStartTime("开始打开表格:" + infoWb, 'green')
    with WorkBook(tarPath, infoWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        infoData = DataDeal(wb.getData())
        for row in dataList[2:]:
            dataMap = {
                'public_msg_id': row[1],
                'msg_loop_cnt': 1,
                'public_contact_id': row[2],
                'public_msg_title': row[3],
                'public_msg_text': row[4],
                'public_msg_picture': row[1],
                'public_msg_text_id': row[1],
                'public_msg_read_num': row[5],
                'public_msg_like_num': row[6],
                'public_msg_comment_id': toStr(row[1]) + '1:' + toStr(row[1]) + '2:' + toStr(row[1]) + '3:' + toStr(row[1]) + '4'
            }
            infoData.setRowData(row[1], dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + infoWb)
        infoData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(infoData.getData())
        wb.save()

    # -----------------------------------public_msg_comment-------------------------------------------------
    printer.setStartTime("开始打开表格:" + commentWb, 'green')
    with WorkBook(tarPath, commentWb, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        commentData = MsgDeal(wb.getData())
        for row in dataList[2:]:
            for col in range(1, 5):
                commentId = row[1] * 10 + col
                commentList = str.splitlines(row[6 + col])

                dataMap = {
                    'public_msg_comment_id': commentId,
                    'public_msg_comment_name': commentList[0],
                    'public_msg_comment_icon': commentList[2],
                    'public_msg_comment_text': commentList[1],
                    'public_msg_comment_like_num': commentList[3],
                }
                descMap = {'相关id[hide]': row[1]}
                commentData.setRowData(commentId, dataMap)
                commentData.setRowData(commentId, descMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + commentWb)
        commentData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(commentData.getData())
        wb.save()