from os import truncate
from myClass.dataDeal import DataDeal
import datetime

toStr = DataDeal.toStr
isEmpty = DataDeal.isEmpty


class TaskDeal(DataDeal):
    fightTextMap = {
        '节目描述': lambda x: TaskDeal.getFightDesc(x),
        '开始描述': lambda x: TaskDeal.getFightStartDesc(x),
        '结束描述': lambda x: TaskDeal.getFightEndDesc(x),
        '失败评论1': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 1, 0)),
        '失败评论2': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 2, 0)),
        '失败评论3': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 3, 0)),
        '失败评论4': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 4, 0)),
        '一星评论1': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 5, 0)),
        '一星评论2': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 6, 0)),
        '一星评论3': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 7, 0)),
        '一星评论4': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 8, 0)),
        '二星评论1': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 9, 0)),
        '二星评论2': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 10, 0)),
        '二星评论3': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 11, 0)),
        '二星评论4': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 12, 0)),
        '三星评论1': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 13, 0)),
        '三星评论2': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 14, 0)),
        '三星评论3': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 15, 0)),
        '三星评论4': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 16, 0)),
        '第一轮目标人群': lambda x: TaskDeal.getGroupTitleId(x, 1),
        '第一轮描述': lambda x: TaskDeal.getGroupDescId(x, 1),
        '第二轮目标人群': lambda x: TaskDeal.getGroupTitleId(x, 2),
        '第二轮描述': lambda x: TaskDeal.getGroupDescId(x, 2),
        '第三轮目标人群': lambda x: TaskDeal.getGroupTitleId(x, 3),
        '第三轮描述': lambda x: TaskDeal.getGroupDescId(x, 3),
        '失败反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 1, 1)),
        '目标人群1的负反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 2, 1)),
        '目标人群2的负反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 3, 1)),
        '目标人群3的负反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 4, 1)),
        '目标人群1的正反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 5, 1)),
        '目标人群2的正反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 6, 1)),
        '目标人群3的正反馈': lambda x: TaskDeal.getCommentText(TaskDeal.getCommentId(x, 7, 1)),
        '事件1描述': lambda x: TaskDeal.getEventDesc(TaskDeal.getEventId(x, 1)),
        '事件1正反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 1), 1),
        '事件1标签负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 1), 2),
        '事件1专长负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 1), 3),
        '事件1负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 1), 4),
        '事件1无选项反馈': lambda x: TaskDeal.getEventNull(TaskDeal.getEventId(x, 1)),
        '事件1女主对白': lambda x: TaskDeal.getEventRoleDesc(TaskDeal.getEventId(x, 1)),
        '事件2描述': lambda x: TaskDeal.getEventDesc(TaskDeal.getEventId(x, 2)),
        '事件2正反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 2), 1),
        '事件2标签负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 2), 2),
        '事件2专长负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 2), 3),
        '事件2负反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 2), 4),
        '事件2无选项反馈': lambda x: TaskDeal.getEventNull(TaskDeal.getEventId(x, 2)),
        '事件2女主对白': lambda x: TaskDeal.getEventRoleDesc(TaskDeal.getEventId(x, 2)),
        '知彼事件描述': lambda x: TaskDeal.getEventDesc(TaskDeal.getEventId(x, 3)),
        '知彼事件选项A': lambda x: TaskDeal.getEventChoose(TaskDeal.getEventId(x, 3), 1),
        '知彼事件选项B': lambda x: TaskDeal.getEventChoose(TaskDeal.getEventId(x, 3), 2),
        '知彼事件选项C': lambda x: TaskDeal.getEventChoose(TaskDeal.getEventId(x, 3), 3),
        '知彼事件选项A反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 3), 1),
        '知彼事件选项B反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 3), 2),
        '知彼事件选项C反馈': lambda x: TaskDeal.getEventFeedback(TaskDeal.getEventId(x, 3), 3),
    }

    def __init__(self, rngData, colRow=0):
        DataDeal.__init__(self, rngData, colRow)

    # 获取对应文本
    def getRowValue(self, rowStr, returnCol=1):
        return self.getRowData(rowStr, returnCol)

    # 获取解锁id
    def getUnlockId2(self, chapterId, sId):
        if sId == 1:
            findRows = self.getRows(chapterId - 1, "chapter_id")
            secCol = self.getColOrder('section_id')
            maxSecId = 1
            for row in findRows:
                if row[secCol] > maxSecId:
                    maxSecId = row[secCol]
            findRow = self.getRow2(chapterId - 1, "chapter_id", maxSecId, 'section_id')
            lineId = self.getColOrder('task_line_id')
            return findRow[lineId]
        else:
            return chapterId * 100 + sId - 1

    # 获取知彼事件正确答案
    def getRightAnswer(self):
        if not isEmpty(self.getRowValue('知彼事件选项A', 2)):
            return 1
        elif not isEmpty(self.getRowValue('知彼事件选项B', 2)):
            return 2
        elif not isEmpty(self.getRowValue('知彼事件选项C', 2)):
            return 3
        else:
            return -1

    # 获取知彼男主
    def getEventMale(self):
        male = self.getRowValue('知彼事件描述', 2)
        if male == '李泽言':
            return 1
        elif male == '许墨':
            return 2
        elif male == '周棋洛':
            return 3
        elif male == '白起':
            return 4
        elif male == '凌肖':
            return 8

    # 获取事件图片
    def getEventImg(self):
        img1 = toStr(self.getRowValue('知彼事件选项A', 3))
        img2 = toStr(self.getRowValue('知彼事件选项B', 3))
        img3 = toStr(self.getRowValue('知彼事件选项C', 3))
        return img1 + ':' + img2 + ':' + img3

    # 获取专家tag
    def getEventTags(self, order):
        tag1Str = ''
        tag2Str = ''
        if order == 1:
            tag1Str = self.getRowValue('事件1描述', 2)
            tag2Str = self.getRowValue('事件1描述', 3)
        else:
            tag1Str = self.getRowValue('事件2描述', 2)
            tag2Str = self.getRowValue('事件2描述', 3)

        tag1 = toStr(self.getRowData(tag1Str, 'staff_tag_id', 'tag_desc', 1))
        tag2 = toStr(self.getRowData(tag2Str, 'staff_tag_id', 'tag_desc', 1))
        return tag1 + ':' + tag2

    # 获取事件权重
    def getEventWeight(self, eventNum, order):
        if eventNum == 2:
            return 50
        elif eventNum == 3:
            if order == 3:
                return 40
            else:
                return 30

    # 获取评论者
    def getCommentName(self, commentId):
        row = int((commentId * (commentId % 10000)) % 1000)
        return self.getData()[row][1]

    # 获取章节id
    @classmethod
    def getChapterId(cls, cId):
        return 3100 + cId

    # 获取解锁id
    @classmethod
    def getUnlockId(cls, chapterId, sId):
        return chapterId * 100 + sId

    # 获取关卡线id
    @classmethod
    def getLineId(cls, chapterId, sId):
        return chapterId * 100 + sId

    # 获取首关id
    @classmethod
    def getFirstId(cls, chapterId):
        return chapterId * 100 + 1

    # 获取星星宝箱id
    @classmethod
    def getBoxId(cls, chapterId, orderId):
        return chapterId * 100 + orderId + 1

    # 获取星星宝箱条件
    @classmethod
    def getBoxStar(cls, fightNum, orderId):
        return fightNum * (orderId + 1)

    # 获取星星宝箱奖励
    @classmethod
    def getBoxReward(cls, orderId):
        if orderId == 0:
            return '0,1,12000'
        elif orderId == 1:
            return '120003,101,7'
        else:
            return '100061,101,1'

    # 获取星星宝箱重要物品展示
    @classmethod
    def getBoxItem(cls, orderId):
        if orderId == 0:
            return '0,1'
        elif orderId == 1:
            return '120003,101'
        else:
            return '100061,101'

    # 获取节点类型
    @classmethod
    def getSecType(cls, secType, unlockNum):
        if secType == '剧情':
            if unlockNum == 0 or isEmpty(unlockNum):
                return 0
            else:
                return 2
        else:
            return 1

    # 获取节点id
    @classmethod
    def getSecId(cls, cId, storyOrder):
        return 120000 + cId * 100 + storyOrder

    # 获取剧情解锁道具
    @classmethod
    def getUnlockItem(cls, unlockItem, unlockNum):
        if unlockNum == 0 or isEmpty(unlockNum):
            return ''
        else:
            return toStr(unlockItem) + ',101,' + toStr(unlockNum)

    # 获取标题id
    @classmethod
    def getTitleId(cls, lineId):
        return 40000000 + lineId * 10 + 1

    # 获取简介id
    @classmethod
    def getDescId(cls, lineId):
        return 40000000 + lineId * 10 + 2

    # 获取事件id
    @classmethod
    def getEventId(cls, fightId, eventOrder):
        return fightId * 10 + eventOrder

    # 获取事件组
    @classmethod
    def getEventIds(cls, fightId, eventNum):
        eventStr = toStr(cls.getEventId(fightId, 1))
        for i in range(1, eventNum):
            eventStr = eventStr + ':' + toStr(cls.getEventId(fightId, i + 1))
        return eventStr

    # 获取buff
    @classmethod
    def getBuff(cls, fightId, special):
        buffStr = 'BUFF_NT5_'

        if special == 0:
            order = fightId * 7 % 24 + 1
            buff1 = toStr(order)
            return buffStr + buff1
        else:
            order1 = fightId * 7 % 24 + 1
            order2 = (order1 + 8) % 24 + 1
            order3 = (order1 + 16) % 24 + 1
            buff1 = toStr(order1)
            buff2 = toStr(order2)
            buff3 = toStr(order3)
            return buffStr + buff1 + ':' + buffStr + buff2 + ':' + buffStr + buff3

    # 获取关卡分数
    @classmethod
    def getScore(cls, score, special):
        score1 = ''
        score2 = ''
        score3 = ''
        if special == 0:
            score1 = toStr(score)
            score2 = toStr(int(score * 0.8))
            score3 = toStr(int(score * 0.1))
        else:
            score1 = toStr(int(score * 2.4))
            score2 = toStr(int(score * 1.92))
            score3 = toStr(int(score * 0.24))
        return score1 + ':' + score2 + ':' + score3

    # 获取热播系数
    @classmethod
    def getProp(cls, fightId):
        return 1400 + (fightId * fightId * 7) % 500

    # 获取掉落
    @classmethod
    def getDrop(cls, item, special):
        str1 = '3333:120002,101,2'
        str2 = '3333:' + toStr(item) + ',101,1'
        str3 = '10000:120001,101,1'
        reStr = str1 + '\n' + str2 + '\n' + str3
        if special == 1:
            reStr = reStr + '\n' + '10000:120002,101,1'
        return reStr

    # 获取掉落icon
    @classmethod
    def getDropIcon(cls, item1, item2):
        returnStr = ''
        if not isEmpty(item1):
            returnStr = toStr(item1) + ',101,1'

        if not isEmpty(item2):
            str2 = toStr(item2) + ',101,1'
            if returnStr == '':
                returnStr = str2
            else:
                returnStr = returnStr + ':' + str2
        return returnStr

    # 节目描述
    @classmethod
    def getFightDesc(cls, fightId):
        return 10000000 + fightId * 1000

    # 节目开始描述
    @classmethod
    def getFightStartDesc(cls, fightId):
        return 10000000 + fightId * 1000 + 1

    # 节目结束描述
    @classmethod
    def getFightEndDesc(cls, fightId):
        return 10000000 + fightId * 1000 + 2

    # 获取评论文本
    @classmethod
    def getCommentText(cls, commentId):
        return commentId + 20000000

    # 获取评论id
    @classmethod
    def getCommentId(cls, fight, order, special):
        if special == 0:
            if order < 5:
                return cls.getCommentF(fight, order)
            elif order < 9:
                return cls.getCommentB(fight, order - 4)
            elif order < 13:
                return cls.getCommentA(fight, order - 8)
            else:
                return cls.getCommentS(fight, order - 12)
        else:
            if order < 2:
                return cls.getCommentF(fight, order)
            elif order < 5:
                return cls.getCommentB(fight, order - 1)
            else:
                return cls.getCommentS(fight, order - 4)

    # 节目失败评论
    @classmethod
    def getCommentF(cls, fightId, order):
        return fightId * 1000 + 100 + order

    # 节目失败评论
    @classmethod
    def getCommentsF(cls, fightId, special):
        if special == 0:
            str1 = toStr(cls.getCommentF(fightId, 1))
            str2 = toStr(cls.getCommentF(fightId, 2))
            str3 = toStr(cls.getCommentF(fightId, 3))
            str4 = toStr(cls.getCommentF(fightId, 4))
            return str1 + ':' + str2 + ':' + str3 + ':' + str4
        else:
            return cls.getCommentF(fightId, 1)

    # 节目1星评论
    @classmethod
    def getCommentB(cls, fightId, order):
        return fightId * 1000 + 200 + order

    # 节目1星评论
    @classmethod
    def getCommentsB(cls, fightId, special):
        if special == 0:
            str1 = toStr(cls.getCommentB(fightId, 1))
            str2 = toStr(cls.getCommentB(fightId, 2))
            str3 = toStr(cls.getCommentB(fightId, 3))
            str4 = toStr(cls.getCommentB(fightId, 4))
            return str1 + ':' + str2 + ':' + str3 + ':' + str4
        else:
            str1 = toStr(cls.getCommentB(fightId, 1))
            str2 = toStr(cls.getCommentB(fightId, 2))
            str3 = toStr(cls.getCommentB(fightId, 3))
            return str1 + ':' + str2 + ':' + str3

    # 节目2星评论
    @classmethod
    def getCommentA(cls, fightId, order):
        return fightId * 1000 + 300 + order

    # 节目2星评论
    @classmethod
    def getCommentsA(cls, fightId, special):
        if special == 0:
            str1 = toStr(cls.getCommentA(fightId, 1))
            str2 = toStr(cls.getCommentA(fightId, 2))
            str3 = toStr(cls.getCommentA(fightId, 3))
            str4 = toStr(cls.getCommentA(fightId, 4))
            return str1 + ':' + str2 + ':' + str3 + ':' + str4
        else:
            str1 = toStr(cls.getCommentB(fightId, 1))
            str2 = toStr(cls.getCommentB(fightId, 2))
            str3 = toStr(cls.getCommentB(fightId, 3))
            return str1 + ':' + str2 + ':' + str3

    # 节目23星评论
    @classmethod
    def getCommentS(cls, fightId, order):
        return fightId * 1000 + 400 + order

    # 节目3星评论
    @classmethod
    def getCommentsS(cls, fightId, special):
        if special == 0:
            str1 = toStr(cls.getCommentS(fightId, 1))
            str2 = toStr(cls.getCommentS(fightId, 2))
            str3 = toStr(cls.getCommentS(fightId, 3))
            str4 = toStr(cls.getCommentS(fightId, 4))
            return str1 + ':' + str2 + ':' + str3 + ':' + str4
        else:
            str1 = toStr(cls.getCommentS(fightId, 1))
            str2 = toStr(cls.getCommentS(fightId, 2))
            str3 = toStr(cls.getCommentS(fightId, 3))
            return str1 + ':' + str2 + ':' + str3

    # 转发数
    @classmethod
    def getShareNum(cls, fightId, special):
        if special == 0:
            return 50000 + (fightId * fightId) % 30000
        else:
            return 9999

    # 评论数
    @classmethod
    def getCommentNum(cls, fightId, special):
        if special == 0:
            return 20000 + (fightId * fightId * 7) % 20000
        else:
            return 9999

    # 评论头像
    @classmethod
    def getCommentHead(cls, commontId):
        return 101 + (commontId * (commontId % 1000)) % 77

    @classmethod
    def getGroupTitleId(cls, fightId, order):
        return fightId * 10000 + 5000 + order

    # 组名称
    @classmethod
    def getGroupTitleIds(cls, fightId):
        str1 = toStr(cls.getGroupTitleId(fightId, 1))
        str2 = toStr(cls.getGroupTitleId(fightId, 2))
        str3 = toStr(cls.getGroupTitleId(fightId, 3))
        return str1 + ':' + str2 + ':' + str3

    @classmethod
    def getGroupDescId(cls, fightId, order):
        return fightId * 10000 + 5003 + order

    # 组描述
    @classmethod
    def getGroupDescIds(cls, fightId):
        str1 = toStr(cls.getGroupDescId(fightId, 1))
        str2 = toStr(cls.getGroupDescId(fightId, 2))
        str3 = toStr(cls.getGroupDescId(fightId, 3))
        return str1 + ':' + str2 + ':' + str3

    # 组描述
    @classmethod
    def getGroupImg(cls, img1, img2, img3):
        str1 = toStr(img1)
        str2 = toStr(img2)
        str3 = toStr(img3)
        return str1 + ':' + str2 + ':' + str3

    # 获取事件描述
    @classmethod
    def getEventDesc(cls, eventId):
        return 10000000 + eventId * 100

    # 获取事件评论
    @classmethod
    def getEventFeedback(cls, eventId, order):
        return 10000000 + eventId * 100 + 10 + order

    # 获取事件评论
    @classmethod
    def getEventFeedbacks(cls, eventNum, eventId):
        returnStr = toStr(cls.getEventFeedback(eventId, 1))
        returnNum = 0
        if eventNum == 2:
            returnNum = 4
        elif eventNum == 3:
            returnNum = 3
        for i in range(1, returnNum):
            returnStr = returnStr + ':' + toStr(cls.getEventFeedback(eventId, i + 1))
        return returnStr

    # 获取未选择反馈
    @classmethod
    def getEventNull(cls, eventId):
        return 10000000 + eventId * 100 + 19

    # 获取女主评论
    @classmethod
    def getEventRoleDesc(cls, eventId):
        return 10000000 + eventId * 100 + 20

    # 获取事件选项
    @classmethod
    def getEventChoose(cls, eventId, order):
        return 10000000 + eventId * 100 + order

    # 获取事件选项
    @classmethod
    def getEventChooses(cls, eventId):
        returnStr = toStr(cls.getEventChoose(eventId, 1))
        for i in range(1, 3):
            returnStr = returnStr + ':' + toStr(cls.getEventChoose(eventId, i + 1))
        return returnStr

    # 获取文本id
    @classmethod
    def getTextId(cls, textType, lineId, fightId):
        if textType == '节目名':
            return cls.getTitleId(lineId)
        elif textType == '节目简介':
            return cls.getDescId(lineId)
        else:
            if textType in cls.fightTextMap:
                return cls.fightTextMap[textType](fightId)
            else:
                return -1

    @classmethod
    def getAnswerEffect(cls, rightAnswer):
        if rightAnswer == 1:
            return '100:20:20'
        elif rightAnswer == 2:
            return '20:100:20'
        elif rightAnswer == 3:
            return '20:20:100'

    @classmethod
    def getCommentApprove(cls, fightId, order, special):
        num = cls.getRandomCommentNum(fightId)
        offset = 1
        if special == 0:
            offset = 1 + (order - 1) % 4
        else:
            if order < 2:
                pass
            elif order < 5:
                offset = order - 1
            else:
                offset = order - 4
        num = round(num / (1 + offset * 0.3) * (1 + order * 0.1), 1)
        return str(num) + '万'

    @classmethod
    def getCommentForward(cls, fightId, order):
        num = cls.getRandomCommentNum(fightId)
        offset = 1
        if order < 2:
            pass
        elif order < 5:
            offset = order - 1
        else:
            offset = order - 4
        num = num / (1 + offset * 0.35) * (1 + order * 0.12) / (10 + fightId % 10)
        if num < 1:
            return int(num * 9000)
        else:
            return str(int(num * (1 + (fightId % 10) * 0.3))) + '万+'

    @classmethod
    def getCommentComment(cls, fightId, order):
        num = cls.getRandomCommentNum(fightId)
        offset = 1
        if order < 2:
            pass
        elif order < 5:
            offset = order - 1
        else:
            offset = order - 4
        num = num / (1 + offset * 0.25) * (1 + order * 0.1) / (10 + fightId % 10)
        if num < 1:
            return int(num * 9000)
        else:
            return str(int(num * (1 + (fightId % 10) * 0.3))) + '万+'

    @classmethod
    def getRandomCommentNum(cls, id):
        return (id * (id % 10) * 0.1) % 10 + 15

    @classmethod
    def getGoodsSet(cls, oriStr, fightId):
        fightStr = toStr(fightId)
        if oriStr == '':
            return '3,' + fightStr
        elif fightStr in oriStr:
            return oriStr
        else:
            return oriStr + ':3,' + fightStr

    @classmethod
    def needUpdate(cls, updateList, shtName):
        if updateList[1][0] == '全部更新':
            if updateList[1][1] == 1:
                return True

        shtStr = shtName.split('.')[0]
        for row in updateList[2:]:
            if shtStr == row[0].split('.')[0]:
                if isEmpty(row[1]) or row[1] == 0:
                    return False
                else:
                    return True

    @classmethod
    def getWeather(cls, placeId, weatherStr):
        if isEmpty(weatherStr) or placeId > 300:
            return ''
        elif '晚' in weatherStr:
            return 'night'
        elif '雾' in weatherStr:
            return 'fog'
        else:
            return ''

    @classmethod
    def getExtraReward(cls, id, special):
        if isEmpty(id):
            return ''
        elif special == 0:
            return id
        else:
            return id + 1
