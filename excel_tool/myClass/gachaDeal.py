from myClass.dataDeal import DataDeal
import datetime
toStr = DataDeal.toStr
toNum = DataDeal.toNum
toInt = DataDeal.toInt


class GachaDeal(DataDeal):
    furniturePartMap = {1: '大玩具', 2: '饭盆', 3: '厕所', 4: '小玩具', 5: '床垫'}
    decoratioPartMap = {1: '头部', 2: '颈部', 3: '背部', 4: '面部'}

    def __init__(self, paraList, itemList, ballList, propertyList, colRow=0):
        DataDeal.__init__(self, None, colRow)
        self.poolStrMap = {}  # 奖池类型标记
        self.paraList = paraList
        self.itemList = itemList
        self.ballList = ballList
        self.propertyList = propertyList
        self.initCol()
        self.poolUpdateMap = self.getUpdatePool()  # 奖池更新标记
        self.priceMap = {}  # 用于商店排序

    # 初始化列的位置
    def initCol(self):
        for r in self.paraList:
            if r[0] == '奖池类型':
                for c in range(1, len(r)):
                    self.poolStrMap[r[c]] = c
                break

        self.ballPoolCol = GachaDeal.getDataColOrder(self.ballList, '奖池')
        self.ballTypeCol = GachaDeal.getDataColOrder(self.ballList, '蛋类型')
        self.ballQualityCol = GachaDeal.getDataColOrder(self.ballList, '品质')
        self.ballGroupCol = GachaDeal.getDataColOrder(self.ballList, '组权重')
        self.ballRandomCol = GachaDeal.getDataColOrder(self.ballList, '随机权重')
        self.ballProtectCol = GachaDeal.getDataColOrder(self.ballList, '刷新保底')

        self.dropUniqueCol = GachaDeal.getDataColOrder(self.ballList, '去重')
        self.dropProtectQualityCol = GachaDeal.getDataColOrder(self.ballList, '保底品质')
        self.dropProtectCol = GachaDeal.getDataColOrder(self.ballList, '掉落保底')
        self.dropStar0Col = GachaDeal.getDataColOrder(self.ballList, '0星')
        self.dropStar1Col = GachaDeal.getDataColOrder(self.ballList, '1星')
        self.dropStar2Col = GachaDeal.getDataColOrder(self.ballList, '2星')
        self.dropStar3Col = GachaDeal.getDataColOrder(self.ballList, '3星')
        self.dropTicketCol = GachaDeal.getDataColOrder(self.ballList, '券数')

        self.itemTypeCol = GachaDeal.getDataColOrder(self.itemList, 'type')
        self.itemQualityCol = GachaDeal.getDataColOrder(self.itemList, '品质')
        self.itemCatCol = GachaDeal.getDataColOrder(self.itemList, 'cat_id')
        self.itemDogCol = GachaDeal.getDataColOrder(self.itemList, 'dog_id')
        self.itemPoolCol = GachaDeal.getDataColOrder(self.itemList, '扭蛋池')
        self.itemShopCol = GachaDeal.getDataColOrder(self.itemList, '商店')
        self.itemActivityCol = GachaDeal.getDataColOrder(self.itemList, '奖励活动')
        self.itemActivityTimeCol = GachaDeal.getDataColOrder(self.itemList, '活动时间')
        self.itemActivityEndTimeCol = GachaDeal.getDataColOrder(self.itemList, '结束时间')
        self.itemMailCol = GachaDeal.getDataColOrder(self.itemList, '补发邮件')

        self.propertyTypeCol = GachaDeal.getDataColOrder(self.propertyList, '类型')
        self.propertyQualityCol = GachaDeal.getDataColOrder(self.propertyList, '品质')
        self.propertyNumCol = GachaDeal.getDataColOrder(self.propertyList, '属性')
        self.propertyParaCol = GachaDeal.getDataColOrder(self.propertyList, '属性参数')
        self.propertyPriceCol = GachaDeal.getDataColOrder(self.propertyList, '售价')
        self.propertyConvertCol = GachaDeal.getDataColOrder(self.propertyList, '分解')

    # 家具表数据
    def setFurnitureData(self, dataDeal: DataDeal):
        self.furnitureData = dataDeal
        self.furnitureQualityCol = dataDeal.getColOrder('quality')
        self.furniturePackCol = dataDeal.getColOrder('backpack_type')
        self.furnitureTypeCol = dataDeal.getColOrder('type')
        self.furniturePetCol = dataDeal.getColOrder('pet')
        self.furnitureAddCol = dataDeal.getColOrder('property_add')
        self.furniturePeriodCol = dataDeal.getColOrder('property_add_period')
        self.furnitureTimesCol = dataDeal.getColOrder('property_add_times')
        self.furnitureVolumeCol = dataDeal.getColOrder('volume')
        self.furnitureConvertCol = dataDeal.getColOrder('convert_items')
        self.furnitureSourceCol = dataDeal.getColOrder('source')

    # 饰品表数据
    def setDecorationData(self, dataDeal: DataDeal):
        self.decorationData = dataDeal
        self.decorationTypeCol = dataDeal.getColOrder('decor_part')
        self.decorationPetCol = dataDeal.getColOrder('species_label_id')
        self.decorationQualityCol = dataDeal.getColOrder('decor_rare')
        self.decorationPropertyCol = dataDeal.getColOrder('get_buff')
        self.decorationDressCol = dataDeal.getColOrder('dress_buff')
        self.decorationConvertCol = dataDeal.getColOrder('convert_items')
        self.decorationSourceCol = dataDeal.getColOrder('decor_set')
        self.decorationClearCol = dataDeal.getColOrder('tournament_time_clear')

    # 道具表数据
    def setGoodsData(self, dataDeal: DataDeal):
        self.goodData = dataDeal
        self.goodPropertyCol = dataDeal.getColOrder('property_change')
        self.goodUseCol = dataDeal.getColOrder('use_limit')

    # 获取参数名对应的值
    def getParaValue(self, poolStr, paraStr):
        paraCol = self.poolStrMap[poolStr]
        for p in self.paraList:
            if p[0] == paraStr:
                return p[paraCol]
        else:
            return None

    # 获取更新的池子
    def getUpdatePool(self):
        updatePool = {}
        for p in self.poolStrMap:
            flag = False
            if self.getParaValue(p, '是否更新') == 1:
                flag = True
            updatePool[p] = flag
        return updatePool

    # 获取道具品质
    def getItemQuality(self, item, itemType):
        if itemType == '道具':
            return None
        elif itemType == '服饰':
            return self.decorationData.getRowData(item, self.decorationQualityCol) / 2
        else:
            return self.furnitureData.getRowData(item, self.furnitureQualityCol) / 2

    # 获取道具类型的id
    def getItemTypeId(self, itemType, itemId):
        if itemType == '道具':
            return 375
        elif itemType == '服饰':
            return 374
        else:
            itemPack = self.furnitureData.getRowData(itemId, self.furniturePackCol)
            if itemPack == 1:
                return 373
            else:
                return 380

    # 获取奖励
    def getWeightReward(self, rewardDetail, petType, refreshId, ballType, isUnique, ticketNum, weight0, weight1, weight2,weight3): # yapf:disable
        weight0 = toNum(weight0)
        weight1 = toNum(weight1)
        weight2 = toNum(weight2)
        weight3 = toNum(weight3)
        if self.isNumberValid(ticketNum):
            maxWeight = 1.0
            if '神秘' in ballType:
                maxWeight = 0.25
            ticketWeight = maxWeight - weight0 - weight1 - weight2 - weight3
            rewardDetail = self.getWeightTicket(rewardDetail, ballType, ticketWeight, ticketNum)
        if self.isNumberValid(weight0):
            rewardDetail = self.getWeightRewardStr(rewardDetail, petType, refreshId, ballType, isUnique, 0, weight0)
        if self.isNumberValid(weight1):
            rewardDetail = self.getWeightRewardStr(rewardDetail, petType, refreshId, ballType, isUnique, 1, weight1)
        if self.isNumberValid(weight2):
            rewardDetail = self.getWeightRewardStr(rewardDetail, petType, refreshId, ballType, isUnique, 2, weight2)
        if self.isNumberValid(weight3):
            rewardDetail = self.getWeightRewardStr(rewardDetail, petType, refreshId, ballType, isUnique, 3, weight3)
        return rewardDetail

    # 获取保底奖励
    def getProtectWeightReward(self, rewardDetail, petType, refreshId, ballType, isUnique, qualityStr):
        quality = toInt(qualityStr[0])
        rewardDetail = self.getWeightRewardStr(rewardDetail, petType, refreshId, ballType, isUnique, quality, 1)
        return rewardDetail

    # 获取奖励详情
    def getWeightRewardStr(self, rewardStr, petType, refreshId, ballType, isUnique, quality, weight):
        tempStr = ''
        if isUnique == 1:
            tempStr = toStr(self.getUniqueRewardId(refreshId, ballType, quality)) + ',107,1|' + toStr(toInt(weight * 10000))
        else:
            tempStr = self.getWeightItem(petType, refreshId, ballType, quality, weight)

        if rewardStr == '':
            rewardStr = tempStr
        elif tempStr != '':
            rewardStr = rewardStr + ';' + tempStr
        return rewardStr

    # 获取非去重奖励
    def getWeightItem(self, petType, refreshId, ballType, quality, weight):
        returnStr = ''
        itemArray = []
        typeArray = []
        refreshId = toStr(refreshId)
        for row in self.itemList[1:]:
            itemType = row[self.itemTypeCol]
            itemPools = self.split(row[self.itemPoolCol], '|')
            if itemType in ballType and refreshId in itemPools:
                itemId = row[self.itemCatCol]
                if petType == 1:
                    itemId = row[self.itemDogCol]

                itemQuality = row[self.itemQualityCol]
                if itemQuality == None:
                    itemQuality = self.getItemQuality(itemId, itemType)
                if itemQuality == quality:
                    itemArray.append(itemId)
                    typeArray.append(self.getItemTypeId(itemType, itemId))

        if len(itemArray) > 0:
            itemWeight = toInt(weight * 10000 / len(itemArray))
            lastWeight = toInt(weight * 10000 - itemWeight * (len(itemArray) - 1))
            for i in range(len(itemArray)):
                itemId = itemArray[i]
                itemTypeId = typeArray[i]
                if i == len(itemArray) - 1:
                    itemWeight = lastWeight
                if returnStr == '':
                    returnStr = toStr(itemId) + ',' + toStr(itemTypeId) + ',1|' + toStr(itemWeight)
                else:
                    returnStr = returnStr + ';' + toStr(itemId) + ',' + toStr(itemTypeId) + ',1|' + toStr(itemWeight)
        return returnStr

    # 获取券奖励
    def getWeightTicket(self, rewardDetail, ballType, ticketWeight, ticketNum):
        ticketWeight1 = toInt(ticketWeight * 5000)
        ticketWeight2 = toInt(ticketWeight * 10000 - ticketWeight1)
        ticket = '6004'
        if '服饰' in ballType:
            ticket = '6005'
        tempStr = ticket + ',375,' + toStr(ticketNum) + '|' + toStr(ticketWeight1)
        tempStr = tempStr + ';' + ticket + ',375,' + toStr(toInt(ticketNum * 1.5)) + '|' + toStr(ticketWeight2)  # yapf:disable

        if rewardDetail == '':
            rewardDetail = tempStr
        else:
            rewardDetail = rewardDetail + ';' + tempStr
        return rewardDetail

    # 获取球是否掉落
    def isBallDrop(self, ballType, groupWeight, randomWeight, protect):
        if '神秘' in ballType:
            return True
        else:
            if toNum(groupWeight) + toNum(randomWeight) + toNum(protect) > 0:
                return True
            else:
                return False

    # 获取goods_drop_pet_group的uid
    def getGroupUid(self, configId, itemStr, tag, itemOrder):
        row = self.getRow3(configId, 'config_id', itemStr, 'rewards', tag, 'true_or_false')
        if row != -1:
            uidCol = self.getColOrder('uid')
            return self.rngData[row][uidCol]
        else:
            id = int(str(configId)[1:5])
            return (id * 10 + tag) * 1000 + itemOrder

    # 获取preview表数据
    def getPreviewIndex(self, refreshId, itemId, itemTypeId):
        itemStr = toStr(itemId) + ',' + toStr(itemTypeId) + ',1'
        if itemId == 6004 or itemId == 6005:
            itemStr = toStr(itemId) + ',375,1'
        findRow = self.getRow2(itemStr, 'reward_value', refreshId, 'refresh_id')
        if findRow == -1:
            return -1
        else:
            return self.rngData[findRow][0]

    # 获取preview表uid
    def getPreviewUid(self, refreshId, itemOrder, itemQuality):
        return refreshId * 10000 + (3 - itemQuality) * 1000 + itemOrder

    # 获取扭蛋池组随机数据
    def getPoolRandomData(self, refreshId, poolStr, randomCol):
        groupData = ''
        for row in self.ballList[1:]:
            if row[self.ballPoolCol] == poolStr:
                ballId = self.getBallId(refreshId, row[self.ballTypeCol], row[self.ballQualityCol])
                if ballId != -1 and toNum(row[randomCol]) > 0:
                    ballStr = toStr(ballId) + ':' + toStr(row[randomCol])
                    if groupData == '':
                        groupData = ballStr
                    else:
                        groupData = groupData + ';' + ballStr
        return groupData

    # 获取扭蛋池保底数据
    def getPoolProtectData(self, refreshId, poolStr):
        protectData = ''
        for i in range(3, -1, -1):
            for row in self.ballList[1:]:
                if row[self.ballPoolCol] == poolStr and toInt(row[self.ballQualityCol]) == i:
                    ballId = self.getBallId(refreshId, row[self.ballTypeCol], row[self.ballQualityCol])
                    if ballId != -1 and toNum(row[self.ballProtectCol]) > 0:
                        ballStr = toStr(ballId) + ':1|' + toStr(row[self.ballProtectCol])
                        if protectData == '':
                            protectData = ballStr
                        else:
                            protectData = protectData + ';' + ballStr
        return protectData

    # 获取唯一openId
    def getPoolOpenId(self, refreshId, openNum):
        findRow = self.getRow2(refreshId, 'refresh_id', openNum, 'open_gacha_num')
        if findRow == -1:
            return self.getMaxId('id') + 1
        else:
            return self.rngData[findRow][0]

    # 获取物品三元组
    def getItemInfo(self, itemId, itemType, itemNum):
        itemTypeId = self.getItemTypeId(itemType, itemId)
        return toStr(itemId) + ',' + toStr(itemTypeId) + ',' + toStr(itemNum)

    # 获取物品部位
    def getItemPart(self, itemId, itemType):
        if itemType == '道具':
            return None
        elif itemType == '服饰':
            part = self.decorationData.getRowData(itemId, self.decorationTypeCol)
            return self.decoratioPartMap[part]
        else:
            part = self.furnitureData.getRowData(itemId, self.furnitureTypeCol)
            return self.furniturePartMap[part]

    # 获取邮件奖励id
    def getMailItemInfo(self, male, itemId, itemType, itemNum):
        itemTypeId = self.getItemTypeId(itemType, itemId)
        return toStr((male * 1000 + itemTypeId) * 100000 + itemId) + ',376,' + toStr(itemNum)

    # 获取物品商店排序
    def getGoodsRank(self, itemID, itemType, male):
        price = self.getItemPropertyData(itemID, itemType, self.propertyPriceCol)
        key = male * 1000 + price
        if key in self.priceMap:
            self.priceMap[key] += 1
        else:
            self.priceMap[key] = 1

        return (99 - price / 5) * 100 + self.priceMap[key]

    # 获取物品价格
    def getItemPrice(self, itemID, itemType):
        price = self.getItemPropertyData(itemID, itemType, self.propertyPriceCol)
        if itemType == '服饰':
            return '6005,375,' + toStr(price)
        else:
            return '6004,375,' + toStr(price)

    # 获取物品属性
    def getItemPropertyData(self, itemId, itemType, returnCol):
        itemQuality = self.getItemQuality(itemId, itemType)
        itemPart = self.getItemPart(itemId, itemType)
        return self.getItemPartProperty(itemPart, itemQuality, returnCol)

    # 获取物品累抽奖池参数
    def getQuestPara(self, petType, itemPool):
        pools = self.split(itemPool, '|')
        return '68:' + pools[petType]

    # 获得任务序号
    def getQuestOrder(self, activityId, questType, role, group):
        order = 0
        typeCol = self.getColOrder('quest_type')
        roleCol = self.getColOrder('role_id')
        groupCol = self.getColOrder('group_id')
        activityCol = self.getColOrder('activity_id')

        for row in self.rngData:
            if row[activityCol] == activityId:
                return order + 1
            elif row[typeCol] == questType and row[roleCol] == role and row[groupCol] == group:
                order += 1
        return order + 1

    # 获取物品属性
    def getItemPartProperty(self, part, quality, returnCol):
        for row in self.propertyList[1:]:
            if row[self.propertyTypeCol] == part and row[self.propertyQualityCol] == quality:
                return row[returnCol]

    # 获取家具属性间隔
    def getFurniturePeriod(self, itemPart, itemQuality):
        partStr = self.furniturePartMap[itemPart]
        if partStr == '小玩具' or partStr == '大玩具':
            return self.getItemPartProperty(partStr, itemQuality, self.propertyParaCol)
        else:
            return None

    # 获取家具属性次数
    def getFurnitureTimes(self, itemPart, itemQuality):
        partStr = self.furniturePartMap[itemPart]
        if partStr == '小玩具' or partStr == '大玩具':
            return int(self.getItemPartProperty(partStr, itemQuality, self.propertyNumCol) / 10)
        else:
            return None

    # 获取饭盆容量
    def getFurnitureVolume(self, itemPart, itemQuality):
        partStr = self.furniturePartMap[itemPart]
        if partStr == '饭盆':
            return self.getItemPartProperty(partStr, itemQuality, self.propertyParaCol)
        else:
            return None

    # 获取物品分解数量
    def getFurnitureConvert(self, itemPart, itemQuality):
        partStr = self.furniturePartMap[itemPart]
        num = self.getItemPartProperty(partStr, itemQuality, self.propertyConvertCol)
        if num == None:
            print(itemPart, itemQuality)
        return '6006,375,' + toStr(num)

    # 获取物品分解数量
    def getDecorationConvert(self, itemPart, itemQuality):
        partStr = self.decoratioPartMap[itemPart]
        num = self.getItemPartProperty(partStr, itemQuality, self.propertyConvertCol)
        if num == None:
            print(itemPart, itemQuality)
        return '6006,375,' + toStr(num)

    # 获取装饰属性
    def getDecorationProperty(self, itemPart, itemQuality):
        partStr = self.decoratioPartMap[itemPart]
        return '0,372,' + toStr(self.getItemPartProperty(partStr, itemQuality, self.propertyNumCol))

    # 获取装饰穿戴属性
    def getDecorationDressProperty(self, itemPart, itemQuality):
        partStr = self.decoratioPartMap[itemPart]
        return int(self.getItemPartProperty(partStr, itemQuality, self.propertyParaCol))

    # 获取蛋id
    @classmethod
    def getBallId(cls, refreshId, ballType, quality):
        ballTypeId = 3
        if ballType == '服饰':
            ballTypeId = 0
        elif ballType == '家具':
            ballTypeId = 1
        elif ballType == '道具':
            ballTypeId = 2
        if quality != None:
            return refreshId * 100 + ballTypeId * 10 + quality + 1
        else:
            return -1

    # 获取奖励组
    @classmethod
    def getBallReward(cls, ballId, special=0):
        id = (2 + special) * 1000000
        return id + ballId * 100

    # 获取去重奖励id
    @classmethod
    def getUniqueRewardId(cls, refreshId, itemType, quality):
        ballId = cls.getBallId(refreshId, itemType, quality)
        rewardId = 3000000 + ballId * 100
        return rewardId

    # 获取物品是否掉落
    @classmethod
    def isItemDrop(cls, refreshId, itemPool):
        itemPools = cls.split(itemPool, '|')
        if toStr(refreshId) in itemPools:
            return True
        else:
            return False

    # 获取priview类型
    @classmethod
    def getPreviewType(cls, itemType):
        if itemType == '服饰':
            return 1
        elif itemType == '家具':
            return 2
        else:
            return 3

    # 获取蛋皮肤
    @classmethod
    def getBallSkin(cls, ballType, quality):
        returnStr = 'icon_pets_gacha_'
        if ballType == '道具':
            returnStr = returnStr + '1'
        elif ballType == '家具':
            returnStr = returnStr + '2'
        elif ballType == '服饰':
            returnStr = returnStr + '3'
        else:
            returnStr = returnStr + '4'
        if '神秘' not in ballType:
            if quality > 0:
                quality = quality * 2 - 1
            returnStr = returnStr + '_' + toStr(quality)
        return returnStr

    # 获取蛋特效
    @classmethod
    def getBallEffect(cls, ballType, quality):
        if '神秘' in ballType:
            return 'effect' + toStr(quality + 1)
        elif ballType == '道具' or quality < 2:
            return ''
        else:
            return 'effect' + toStr(quality)

    # 获取扭蛋额外奖励
    @classmethod
    def getPoolExtraReward(cls, refreshId, openNum, rewardStr):
        qualityMap = {'高': 3, '低': 2}
        if openNum == 10 and rewardStr != None:
            if cls.isNumber(rewardStr):
                rewardQuality = cls.toInt(rewardStr)
            elif rewardStr in qualityMap:
                rewardQuality = qualityMap[rewardStr]
            else:
                rewardQuality = 3
            return refreshId * 100 + 30 + rewardQuality
        else:
            return ''

    # 获取扭蛋额外奖励规则
    @classmethod
    def getPoolExtraRewardRule(cls, openNum, rewardStr, ruleStr):
        ruleMap = {'天': '1', '日': '1', '周': '2', '月': '3', '活动': '4'}
        if openNum == 10 and rewardStr != None:
            if cls.isNumber(ruleStr):
                ruleStr = toStr(ruleStr)
                return ruleStr + ',283,1'
            elif ruleStr in ruleMap:
                return ruleMap[ruleStr] + ',283,1'
            else:
                return ''
        else:
            return ''

    # 获取刷新提示
    @classmethod
    def getRefreshHint(cls, poolStr, ballQuality):
        if poolStr == '普通' and ballQuality > 1:
            return 1
        elif ballQuality > 2:
            return 1
        else:
            return ''

    # 获取累抽活动提示文本
    @classmethod
    def getQuestDesc(cls, beginTime, endTime):
        beginTime = cls.getDateTime(beginTime)
        endTime = cls.getDateTime(endTime)
        if beginTime.day == 1 and endTime.day == 1 and (endTime - beginTime).days > 1:
            return '本月累计扭蛋68:次'
        else:
            return '累计扭蛋68:次'

    @classmethod
    def getMailBeginTime(cls, activityEndTime):
        time = cls.getDateTime(activityEndTime) + datetime.timedelta(seconds=1)
        return cls.getDateString(time)

    @classmethod
    def getMailEndTime(cls, activityEndTime):
        time = cls.getDateTime(activityEndTime) + datetime.timedelta(days=13, hours=19, seconds=-59)
        return cls.getDateString(time)

    @classmethod
    def getDateString(cls, time):
        return datetime.datetime.strftime(time, '%Y/%m/%d %H:%M:%S')

    @classmethod
    def getDateTime(cls, timeStr):
        try:
            return datetime.datetime.strptime(timeStr, '%Y/%m/%d %H:%M:%S')
        except:
            return datetime.datetime.strptime(timeStr, '%Y/%m/%d %H:%M')