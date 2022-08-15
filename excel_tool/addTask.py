from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.taskDeal import TaskDeal
from myClass.printer import Printer
from myClass.caller import Caller
import sys

toStr = TaskDeal.toStr
isEmpty = TaskDeal.isEmpty
printer = Printer()
tarPath = WorkBook.getActivePath()
dataSht = WorkBook.getActiveSht()

groupWb = 'task_chapter_group_info.xlsx'
chapterWb = 'task_chapter_info.xlsx'
boxWb = 'task_chapter_treasurebox.xlsx'
lineWb = 'task_line_info_data.xlsx'
infoWb = 'task_info_data.xlsx'
eventWb = 'task_event_info_data.xlsx'
commentWb = 'task_comment.xlsx'
textWb = 'task_text_data.xlsx'
rewardWb = 'task_extra_reward.xlsx'
goodsWb = 'goods_info_data.xlsx'
achieveWb = 'achievement_info_data.xlsx'
daybookWb = 'quest_my_daybook.xlsx'
unlockWb = 'taskline_unlock_condition.xlsx'

tagWb = 'tag_info_data.xlsx'
placeWb = 'task_map_place_info.xlsx'

maxLine = dataSht.used_range.shape[0]
chapterList = dataSht['A1'].expand('table').value
paraList = dataSht.range(dataSht['G1'], dataSht['G' + toStr(maxLine)].end('up')).value
lineList = dataSht['I1'].expand('table').value
updateList = dataSht['C11'].expand('table').value

outPath = paraList[9]
tagData = None
placeData = None

with WorkApp() as app:
    fightMap = {}  # 章节内战斗关卡数量
    for row in lineList[1:]:
        if row[2] == '战斗':
            if row[0] in fightMap:
                fightMap[row[0]] += 1
            else:
                fightMap[row[0]] = 1

    with WorkBook(tarPath, tagWb, app) as wb:
        tagData = TaskDeal(wb.getData())

    with WorkBook(tarPath, placeWb, app) as wb:
        placeData = TaskDeal(wb.getData())
    # -----------------------------------taskline_unlock_condition-------------------------------------------------
    if TaskDeal.needUpdate(updateList, unlockWb):
        printer.setStartTime("开始打开表格:" + unlockWb, 'green')
        with WorkBook(tarPath, unlockWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            unlockData = TaskDeal(wb.getData())
            chapters = ''
            lineId = 0
            gId = unlockData.getMaxId('id')

            for row in chapterList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])
                if lineId == 0:
                    lineId = TaskDeal.getLineId(chapterId, 1)
                if chapters == '':
                    chapters = toStr(chapterId)
                else:
                    chapters = chapters + ',' + toStr(chapterId)
            if unlockData.getRow(lineId, 'taskline_id') == -1:
                gId += 1
            dataMap = {
                'id': gId,
                'taskline_id': lineId,
                'controlled_chapter_id': chapters,
                'unlock_free_start_time': paraList[7],
                'unlock_free_end_time': paraList[8],
                'unlock_condition': '10212,125,1',
            }
            unlockData.setRowData(lineId, dataMap, 'taskline_id')
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + unlockWb)
            unlockData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(unlockData.getData())
            wb.save()

    # -----------------------------------task_chapter_group_info-------------------------------------------------
    if TaskDeal.needUpdate(updateList, groupWb):
        printer.setStartTime("开始打开表格:" + groupWb, 'green')
        with WorkBook(tarPath, groupWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            groupData = TaskDeal(wb.getData())
            chapters = ''

            for row in chapterList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])
                if chapters == '':
                    chapters = toStr(chapterId)
                else:
                    chapters = chapters + ':' + toStr(chapterId)

            for i in range(2):
                gId = 200 + paraList[10] + i
                dataMap = {
                    'chapter_group_id': gId,
                    'chapter_include': chapters if i == 0 else 'max',
                    'group_name': paraList[11] if i == 0 else '未完待续',
                    'group_pv': 'mainline_' + toStr(paraList[12]) if i == 0 else '',
                    'tag_desc': '' if i == 0 else '未完待续，敬请期待……',
                    'season': 2,
                    'title_img': 'back_pv_s2_v' + toStr(paraList[10]) if i == 0 else '',
                    'PV_desc': paraList[13] if i == 0 else '',
                }
                groupData.setRowData(gId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + groupWb)
            groupData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(groupData.getData())
            wb.save()

    # -----------------------------------task_chapter_info-------------------------------------------------
    if TaskDeal.needUpdate(updateList, chapterWb):
        printer.setStartTime("开始打开表格:" + chapterWb, 'green')
        with WorkBook(tarPath, chapterWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            chapterData = TaskDeal(wb.getData())
            for row in chapterList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])

                dataMap = {
                    'chapter_id': chapterId,
                    'chapter_type': 3,
                    'first_task_id': TaskDeal.getFirstId(chapterId),
                    'chapter_title': row[1],
                    'hard_task_hide':1,
                    'chapter_task_unlock_coin': '' if isEmpty(paraList[2]) else 1,
                    'chapter_desc': row[2],
                    'chapter_img': chapterId,
                } #yapf:disable
                chapterData.setRowData(chapterId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + chapterWb)
            chapterData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(chapterData.getData())
            wb.save()

    # -----------------------------------task_chapter_treasurebox-------------------------------------------------
    if TaskDeal.needUpdate(updateList, boxWb):
        printer.setStartTime("开始打开表格:" + boxWb, 'green')
        with WorkBook(tarPath, boxWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            boxData = TaskDeal(wb.getData())
            for row in chapterList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])
                for i in range(3):
                    boxId = TaskDeal.getBoxId(chapterId, i)
                    dataMap = {
                        'box_id': boxId,
                        'box_need_star': TaskDeal.getBoxStar(row[3],i),
                        'box_rewards':TaskDeal.getBoxReward(i),
                        'box_import_itemid': TaskDeal.getBoxItem(i),
                    } #yapf:disable
                    boxData.setRowData(boxId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + boxWb)
            boxData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(boxData.getData())
            wb.save()

    # -----------------------------------task_line_info_data-------------------------------------------------
    if TaskDeal.needUpdate(updateList, lineWb):
        printer.setStartTime("开始打开表格:" + lineWb, 'green')
        with WorkBook(tarPath, lineWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            lineData = TaskDeal(wb.getData())
            unLockId = paraList[1]
            unLockItem = paraList[2]
            fightId = paraList[3] - 1
            storyOrder = 0
            lastChapter = 0
            placeId = 0
            for row in lineList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])
                lineId = TaskDeal.getLineId(chapterId, row[1])

                if row[2] == '剧情':
                    if lastChapter == chapterId:
                        storyOrder += 1
                    else:
                        lastChapter = chapterId
                        storyOrder = 1
                else:
                    fightId += 1

                if row[1] > 1:
                    unLockId = TaskDeal.getUnlockId(chapterId, row[1])

                if not isEmpty(row[9]):
                    placeId = placeData.getRowData(row[9], 'place_id', 'place_name')

                dataMap = {
                    'task_line_id': lineId,
                    'task_line_unlock_id': unLockId if row[1] == 1 else TaskDeal.getUnlockId(chapterId,row[1]-1),
                    'chapter_id':chapterId,
                    'section_id': row[1],
                    'task_line_type': 0,
                    'male_role_id': -1,
                    'section_type': TaskDeal.getSecType(row[2],row[5]),
                    'section_content_id': fightId if row[2]=='战斗' else TaskDeal.getSecId(row[0],storyOrder),
                    'task_icon': -1,
                    'task_line_rewards':'' if isEmpty(row[6]) or row[2]=='战斗' else row[6],
                    'task_line_unlock_cost': '' if row[2]=='战斗' else TaskDeal.getUnlockItem(unLockItem,row[5]),
                    'task_map_place':placeId,
                    'task_map_weather':TaskDeal.getWeather(placeId,row[10]),
                    'story_title': TaskDeal.getTitleId(lineId),
                    'task_desc': TaskDeal.getDescId(lineId),
                } #yapf:disable
                lineData.setRowData(lineId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + lineWb)
            lineData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(lineData.getData())
            wb.save()

    # -----------------------------------task_info_data-------------------------------------------------
    if TaskDeal.needUpdate(updateList, infoWb):
        printer.setStartTime("开始打开表格:" + infoWb, 'green')
        with WorkBook(tarPath, infoWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            infoData = TaskDeal(wb.getData())
            unLockItem = paraList[2]
            fightId = paraList[3] - 1
            score = paraList[4] - paraList[5]

            for row in lineList[1:]:
                if row[2] == '战斗':
                    fightId += 1
                    score += paraList[5]
                    chapterId = TaskDeal.getChapterId(row[0])
                    lineId = TaskDeal.getLineId(chapterId, row[1])
                    special = 0  # 判断是否三波战斗
                    if row[8] == 1:
                        special = 1

                    dataMap = {
                        'task_id': fightId,
                        'task_type':0,
                        'male_role_id': -1,
                        'task_line_id': lineId,
                        'daily_challenge_times': 0,
                        'heart_need': '0,41,5',
                        'task_buff_group': TaskDeal.getBuff(fightId,special) if special == 0 else '',
                        'cards_groups_cnt': 3 if special == 1 else '',
                        'cards_groups_ratio': '100:80:60' if special == 1 else '',
                        'cards_groups_buff': TaskDeal.getBuff(fightId,special) if special == 1 else '',
                        'task_score_standard': TaskDeal.getScore(score,special),
                        'task_virtual_score_prop': TaskDeal.getProp(fightId) if special == 0 else 0,
                        'task_rewards_money': '100001,1,100',
                        'task_rewards_firm_exp': '0,42,300',
                        'task_rewards_card_exp': '0,44,100'if special == 0 else '0,44,0',
                        'sweep_rewards_card_exp': '120001,101,1'if special == 0 else '120001,101,0',
                        'task_item_drop_a': TaskDeal.getDrop(row[6],special),
                        'task_item_drop_s': TaskDeal.getDrop(row[6],special),
                        'task_rewards_first_pass': '120002,101,1',
                        'task_item_icon': TaskDeal.getDropIcon(paraList[2],row[6]),
                        'task_background': row[7],
                        'tv_desc': TaskDeal.getFightDesc(fightId),
                        'task_start_desc': TaskDeal.getFightStartDesc(fightId),
                        'task_end_desc': TaskDeal.getFightEndDesc(fightId),
                        'task_comment_f': TaskDeal.getCommentsF(fightId,special),
                        'task_comment_b': TaskDeal.getCommentsB(fightId,special),
                        'task_comment_a': TaskDeal.getCommentsA(fightId,special),
                        'task_comment_s': TaskDeal.getCommentsS(fightId,special),
                        'share_num': TaskDeal.getShareNum(fightId,special),
                        'comment_num': TaskDeal.getCommentNum(fightId,special),
                        'task_extra_reward': TaskDeal.getExtraReward(paraList[6],special),
                        'task_packing': 3 if special == 1 else '',
                        'task_rank_condition': toStr(fightId) + ',328,3',
                        'task_season': 2,
                        'cards_groups_title': TaskDeal.getGroupTitleIds(fightId) if special == 1 else '',
                        'card_group_desc': TaskDeal.getGroupDescIds(fightId) if special == 1 else '',
                        'task_BGM': '50_Task_S2' if special == 0 else '51_Battle_2',
                    } #yapf:disable
                    infoData.setRowData(fightId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + infoWb)
            infoData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(infoData.getData())
            wb.save()

    # -----------------------------------goods_info_data-------------------------------------------------
    if TaskDeal.needUpdate(updateList, goodsWb):
        printer.setStartTime("开始打开表格:" + goodsWb, 'green')
        with WorkBook(tarPath, goodsWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            fightId = paraList[3] - 1
            goodsData = TaskDeal(wb.getData())
            for row in lineList[1:]:
                if row[2] == '战斗':
                    fightId += 1
                    goodsId = row[6]
                    oriSet = goodsData.getRowData(goodsId, 'goods_set', 'goods_id')
                    dataMap = {
                        'goods_id': goodsId,
                        'goods_set': TaskDeal.getGoodsSet(oriSet, fightId),
                    }
                    goodsData.setRowData(goodsId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + goodsWb)
            goodsData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(goodsData.getData())
            wb.save()

    # -----------------------------------task_extra_reward-------------------------------------------------
    if TaskDeal.needUpdate(updateList, rewardWb):
        printer.setStartTime("开始打开表格:" + rewardWb, 'green')
        with WorkBook(tarPath, rewardWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            gradeList = [0, 10000, 20000, 25000, 30000, 35000, 40000, 45000, 50000, 60000, 70000, 80000, 90000, 1000000]
            rewardData = TaskDeal(wb.getData())
            for j in range(2):
                for i in range(len(gradeList) - 1):
                    chapterId = TaskDeal.getChapterId(row[0])

                    dataMap = {
                        'id': paraList[6] + j,
                        'grade_id': i + 1,
                        'task_grade_down': gradeList[i] if j == 0 else gradeList[i] * 2.4,
                        'task_grade_up': gradeList[i + 1] - 1 if j == 0 else gradeList[i + 1] * 2.4 - 1,
                        'extra_reward': toStr(paraList[2]) + ',101,' + toStr(i + 8),
                    }
                    rewardData.setRowData(paraList[6] + j, dataMap, extraId=i + 1)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + rewardWb)
            rewardData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(rewardData.getData())
            wb.save()

    # -----------------------------------task_comment-------------------------------------------------
    if TaskDeal.needUpdate(updateList, commentWb):
        printer.setStartTime("开始打开表格:" + commentWb, 'green')
        with WorkBook(tarPath, commentWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            commentData = TaskDeal(wb.getData())
            fightId = paraList[3] - 1

            for row in lineList[1:]:
                if row[2] == '战斗':
                    fightId += 1
                    commentNum = 16
                    special = 0  # 判断是否三波战斗
                    if row[8] == 1:
                        special = 1
                        commentNum = 7
                    for i in range(commentNum):
                        commentId = TaskDeal.getCommentId(fightId, i + 1, special)
                        dataMap = {
                            'id': commentId,
                            'name':commentData.getCommentName(commentId),
                            'head_icon': TaskDeal.getCommentHead(commentId),
                            'comment': TaskDeal.getCommentText(commentId),
                            'approve_num': TaskDeal.getCommentApprove(fightId,i+1,special),
                            'comment_num': TaskDeal.getCommentComment(fightId,i+1) if special == 1 else '',
                            'forward_num': TaskDeal.getCommentForward(fightId,i+1) if special == 1 else '',
                        } #yapf:disable
                        commentData.setRowData(commentId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + commentWb)
            commentData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(commentData.getData())
            wb.save()

    # -----------------------------------achievement_info_data-------------------------------------------------
    if TaskDeal.needUpdate(updateList, achieveWb):
        printer.setStartTime("开始打开表格:" + achieveWb, 'green')
        with WorkBook(tarPath, achieveWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            achieveData = TaskDeal(wb.getData())
            for row in chapterList[1:]:
                chapterId = TaskDeal.getChapterId(row[0])
                chapterChs = TaskDeal.getChineseNum(row[0])
                achId = 25000 + row[0]
                dataMap = {
                    'achv_id': achId,
                    'achv_type': 2,
                    'achv_theme': 50,
                    'achv_rank': row[0],
                    'achv_need': toStr(chapterId) + ',1048,' + toStr(fightMap[row[0]] * 3),
                    'achv_rewards': '100001,2,50',
                    'achv_name': '新的开始·阶段' + chapterChs,
                    'achv_desc': '第二季第' + chapterChs + '章满星通关',
                    'achv_id_old': 0,
                }
                achieveData.setRowData(achId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + achieveWb)
            achieveData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(achieveData.getData())
            wb.save()

    # -----------------------------------quest_my_daybook-------------------------------------------------
    if TaskDeal.needUpdate(updateList, daybookWb):
        printer.setStartTime("开始打开表格:" + daybookWb, 'green')
        with WorkBook(tarPath, daybookWb, app) as wb:
            printer.printGapTime("表格打开完毕，耗时:")

            printer.setStartTime("开始处理数据数据...")
            daybookData = TaskDeal(wb.getData())
            for i in range(1, 6):
                for row in chapterList[1:]:
                    chapterId = TaskDeal.getChapterId(row[0])
                    male = i
                    if male == 5:
                        male = 8
                    questId = 5005600 + male * 100000 + row[0] + 1
                    dataMap = {
                        'quest_id': questId,
                        'activity_id': -1,
                        'group_id': male,
                        'sub_group_id': 56,
                        'quest_type': 5,
                        'pre_quest_id': -1,
                        'target_type': 1,
                        'target_param': 310000 + row[0] * 100 + fightMap[row[0]] * 3,
                        'rewards': toStr(male) + ',347,100',
                        'desc': '通关第二季主线剧情第[c][A96C67FF]' + toStr(row[0]) + '[-][/c]章',
                        'show_order': row[0] + 1,
                        'title': '1#主线故事',
                        'source': '73,' + toStr(chapterId),
                        'is_achievement_data': 1,
                    }
                    daybookData.setRowData(questId, dataMap)
            printer.printGapTime("数据处理完毕，耗时:")

            printer.setStartTime("开始检查单元格格式:" + daybookWb)
            daybookData.strCheck()
            printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

            wb.setData(daybookData.getData())
            wb.save()

    #------------------------------------关卡文案-----------------------------------------------------
    if TaskDeal.needUpdate(updateList, '关卡文案'):
        printer.setStartTime("开始打开文案及相关表格:", 'green')
        outWb = WorkBook(outPath, '', app)
        taskLineWb = WorkBook(tarPath, lineWb, app)
        taskInfoWb = WorkBook(tarPath, infoWb, app)
        taskTextWb = WorkBook(tarPath, textWb, app)
        taskEventWb = WorkBook(tarPath, eventWb, app)
        printer.printGapTime("表格打开完毕，耗时:")

        lineData = TaskDeal(taskLineWb.getData())
        infoData = TaskDeal(taskInfoWb.getData())
        textData = TaskDeal(taskTextWb.getData())
        eventData = TaskDeal(taskEventWb.getData())

        printer.setStartTime("开始处理剧情关文本:", 'green')
        for row in lineList[1:]:
            chapterId = TaskDeal.getChapterId(row[0])
            lineId = TaskDeal.getLineId(chapterId, row[1])
            for i in range(2):
                textId = 0
                if i == 0:
                    textId = TaskDeal.getTitleId(lineId)
                else:
                    textId = TaskDeal.getDescId(lineId)
                textMap = {
                    'id': textId,
                    'desc': row[3] if i == 0 else row[4],
                }
                hideMap = {
                    '相关id[hide]': lineId,
                    'id名称[hide]': chapterId,
                    '文本类型[hide]': '节目名' if i == 0 else '节目简介',
                }
                textData.setRowData(textId, textMap)
                textData.setRowData(textId, hideMap, defRow=3)
        printer.printGapTime("剧情关文本处理完毕，耗时:")

        for sht in outWb.sheets:
            if sht.name[0] != '#' and sht.used_range[
                    0, 0].value == '章节' and sht.used_range[0, 1].value > 0 and sht.used_range[1, 1].value > 0:

                printer.setStartTime("开始处理" + sht.name, 'green')
                shtData = TaskDeal(sht.used_range.value)
                shtData.setOriData(tagData.getData())
                chapterId = TaskDeal.getChapterId(shtData.getRowValue('章节'))
                lineId = TaskDeal.getLineId(chapterId, shtData.getRowValue('编号'))
                fightId = lineData.getRowData(lineId, 'section_content_id')
                event3 = shtData.getRowValue('知彼事件描述')
                special = 0
                if shtData.getRowValue('第一轮目标人群') != -1:
                    special = 1
                infoMap = {
                        'task_id': fightId,
                        'task_name': shtData.getRowValue('节目名'),
                        'task_event_ids': TaskDeal.getEventIds(fightId,2) if isEmpty(event3) else TaskDeal.getEventIds(fightId,3),
                        'type': tagData.getRowData(shtData.getRowValue('节目名',2),'staff_tag_id','tag_desc'),
                        'card_group_img': TaskDeal.getGroupImg(shtData.getRowValue('第一轮目标人群',3),shtData.getRowValue('第二轮目标人群',3),
                                                            shtData.getRowValue('第三轮目标人群',3)) if special == 1 else '',
                    } #yapf:disable
                infoData.setRowData(fightId, infoMap)

                eventNum = 2
                rightAnswer = 0
                if not isEmpty(event3):
                    eventNum += 1
                    rightAnswer = shtData.getRightAnswer()
                for i in range(eventNum):
                    eventId = TaskDeal.getEventId(fightId, i + 1)
                    eventMap = {
                        'task_event_id':TaskDeal.getEventId(fightId,i+1),
                        'task_event_type':1 if i<2 else 2,
                        'task_event_cond':shtData.getEventTags(i+1) if i < 2 else rightAnswer-1,
                        'task_event_effect':'0:0:50:50' if i < 2 else TaskDeal.getAnswerEffect(rightAnswer),
                        'event_option_cnt':1,
                        'task_event_weight':shtData.getEventWeight(eventNum,i+1),
                        'task_event_desc':TaskDeal.getEventDesc(eventId),
                        'task_event_choose':TaskDeal.getEventChooses(eventId) if i == 2 else '',
                        'task_event_feedback':TaskDeal.getEventFeedbacks(eventNum,eventId),
                        'task_event_null':TaskDeal.getEventNull(eventId) if i < 2 else '',
                        'role_desc':TaskDeal.getEventRoleDesc(eventId) if i < 2 else '',
                        'event_sub_type':1 if i ==2 else '',
                        'sub_type_params':shtData.getEventMale() if i == 2 else '',
                        'event_img':shtData.getEventImg() if i == 2 else '',
                    }#yapf:disable
                    eventData.setRowData(eventId, eventMap)

                for row in sht.used_range.value[2:]:
                    if isEmpty(row[1]):
                        continue

                    textId = TaskDeal.getTextId(row[0], lineId, fightId)
                    if textId == -1:
                        continue
                    textMap = {
                        'id': textId,
                        'desc': shtData.getRowValue(row[0]),
                    }
                    hideMap = {
                        '相关id[hide]': fightId,
                        'id名称[hide]': shtData.getRowValue('节目名'),
                        '文本类型[hide]': row[0],
                    }
                    textData.setRowData(textId, textMap)
                    textData.setRowData(textId, hideMap, defRow=3)

                printer.printGapTime(sht.name + "处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:", 'green')
        infoData.strCheck()
        textData.strCheck()
        eventData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)
        printer.printColor("开始保存数据。。。")

        taskInfoWb.setData(infoData.getData())
        taskTextWb.setData(textData.getData())
        taskEventWb.setData(eventData.getData())

        outWb.close()
        taskLineWb.close(False)
        taskInfoWb.close(True)
        taskTextWb.close(True)
        taskEventWb.close(True)
        lineData = None
        infoData = None
        textData = None
    printer.printColor("数据保存完毕，本次程序运行结束", 'green')
