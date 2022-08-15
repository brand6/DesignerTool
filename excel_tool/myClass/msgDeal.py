from myClass.dataDeal import DataDeal
import datetime

toStr = DataDeal.toStr


class MsgDeal(DataDeal):
    def __init__(self, rngData, colRow=0):
        DataDeal.__init__(self, rngData, colRow)
        self.favorData = []
        self.weekTimeData = []

    # 设置好感数据源
    def setFavorData(self, favorData):
        self.favorData = favorData

    # 设置活动数据源
    def setWeekTimeData(self, weekTimeData):
        self.weekTimeData = weekTimeData

    # 获取短电朋好感,returnType:1 电话   2 朋友圈   3 短信
    def getFavorValue(self, msgType, returnType):
        for row in self.favorData:
            if row[0] == msgType:
                return row[returnType]
        else:
            return 0

    # 获取短信背景
    def getMsgBackGround(self, msgType):
        for row in self.favorData:
            if row[0] == msgType:
                return row[4]
        else:
            return 1

    # 获取周活跃数据 returnType 1 库id   2 时间
    def getWeekTime(self, weekStr, returnType):
        for row in self.weekTimeData:
            if row[0] == weekStr:
                return row[returnType]
        else:
            return 1

    # 设置数据的值
    def setRowData(self, checkId, setMap={}, checkCol=0, dataType=0, extraId='', extraCol=1, defRow=-1, insertType=1):
        DataDeal.setRowData(self, checkId, setMap, checkCol, dataType, extraId, extraCol, defRow, insertType)

    # 获取年份
    @classmethod
    def getYear(cls):
        return datetime.datetime.now().year

    # 根据当前时间获取上架时间
    @classmethod
    def getSaleTime(cls, curTime):
        returnTime = datetime.datetime.strptime(curTime, r'%Y/%m/%d %H:%M') + datetime.timedelta(days=365)
        return returnTime.strftime(r'%Y/%m/%d %H:%M:%S')

    # 获取男主回复
    @classmethod
    def getReplyIds(cls, dataList, phoneId, optionOrder):
        returnStr = ''
        optionOrder = toStr(optionOrder)
        for row in dataList[2:]:
            optionId = toStr(row[0])
            if row[1] == phoneId and optionId[-3:-1] == optionOrder and optionId[-1] != '0':
                if returnStr == '':
                    returnStr = optionId
                else:
                    returnStr = returnStr + ';' + optionId
        return returnStr

    # 获取短信的选项 optionOrder:第几轮选项
    @classmethod
    def getMsgOptions(cls, dataList, phoneId, optionOrder):
        returnStr = '-1'
        optionOrder = toStr(optionOrder)
        for row in dataList[2:]:
            optionId = toStr(row[0])
            if row[1] == phoneId and optionId[-1] == '0' and optionId[-3] == optionOrder:
                if returnStr == '-1':
                    returnStr = optionId
                else:
                    returnStr = returnStr + ':' + optionId
        return returnStr

    # 获取开场文本
    @classmethod
    def getFirstIds(cls, dataList, phoneId):
        returnStr = ''
        for row in dataList[2:]:
            optionId = toStr(row[0])
            if row[1] == phoneId and optionId[-3] == '0':
                if returnStr == '':
                    returnStr = optionId
                else:
                    returnStr = returnStr + ',' + optionId
        return returnStr

    # 获取回复
    @classmethod
    def getReplyListStr(cls, optionId, content):
        content = toStr(content)
        listStr = ''
        optionId = float(optionId)
        strList = str.splitlines(content)
        for order in range(len(strList)):
            if order == 0:
                listStr = toStr(optionId + order + 1)
            else:
                listStr = listStr + ';' + toStr(optionId + order + 1)
        return listStr
