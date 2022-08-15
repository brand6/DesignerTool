from myClass.workApp import WorkApp
from myClass.workBook import WorkBook
from myClass.dataDeal import DataDeal
from myClass.gachaDeal import GachaDeal
from myClass.printer import Printer
import sys

toStr = GachaDeal.toStr
printer = Printer()

openWb = 'gacha_open_info.xlsx'
refreshWb = 'gacha_refresh_info.xlsx'
rewardWb = 'gacha_reward_info.xlsx'
previewWb = 'gacha_preview.xlsx'
weightsetWb = 'goods_weightset_data.xlsx'
groupWb = 'goods_drop_pet_group.xlsx'

exchangeWb = 'goods_exchange_info_data.xlsx'
activityWb = 'activity_info_data.xlsx'
questWb = 'pet_quest.xlsx'
mailWb = 'mail_info_data.xlsx'

goodsWb = 'pet_goods_info.xlsx'
furnitureWb = 'pet_furniture.xlsx'
decorationWb = 'pet_decoration.xlsx'

printer.setStartTime("开始获取表格工具的扭蛋数据:", 'green')
xlToolWb = WorkBook.getXlToolWb()
tarPath = WorkBook.getActivePath(xlToolWb)
dataSht = WorkBook.getActiveSht(xlToolWb)
paraList = dataSht['A1'].expand('table').value
ballList = dataSht['F1'].expand('table').value
itemList = dataSht['U1'].expand('table').value
propertyList = dataSht['AG1'].expand('table').value
printer.printGapTime("表格工具数据获取完毕，耗时:")

with WorkApp() as app:
    wbData = GachaDeal(paraList, itemList, ballList, propertyList)
    # -----------------------------------获取家具&装饰&道具品质-----------------------------------------------
    printer.setStartTime("开始从宠物家具/服饰/道具表获取数据:")
    wbName = furnitureWb
    with WorkBook(tarPath, furnitureWb, app) as wb:
        wbData.setData(wb.getData())
        wbData.setFurnitureData(DataDeal(wb.getData()))
        for row in wbData.rngData[4:]:
            if row[0] != None and '领养' not in row[wbData.furnitureSourceCol]:
                itemId = row[0]
                itemPart = row[wbData.furnitureTypeCol]
                itemQuality = row[wbData.furnitureQualityCol] / 2
                indexId = itemId
                dataMap = {
                    'id':indexId,
                    'property_add_period':wbData.getFurniturePeriod(itemPart,itemQuality),
                    'property_add_times':wbData.getFurnitureTimes(itemPart,itemQuality),
                    'volume':wbData.getFurnitureVolume(itemPart,itemQuality),
                    'convert_items':wbData.getFurnitureConvert(itemPart,itemQuality),
                    # 'total_flag_id': ''
                }  #yapf:disable
                wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()
        wbData.setFurnitureData(DataDeal(wb.getData()))

    wbName = decorationWb
    with WorkBook(tarPath, decorationWb, app) as wb:
        wbData.setData(wb.getData())
        wbData.setDecorationData(DataDeal(wb.getData()))
        for row in wbData.rngData[4:]:
            if row[0] != None and row[wbData.decorationClearCol] != 1:
                itemId = row[0]
                itemPart = row[wbData.decorationTypeCol]
                itemQuality = row[wbData.decorationQualityCol] / 2
                indexId = itemId
                dataMap = {
                    'uid':indexId,
                    'get_buff':wbData.getDecorationProperty(itemPart,itemQuality),
                    'dress_buff':wbData.getDecorationDressProperty(itemPart,itemQuality),
                    'convert_items':wbData.getDecorationConvert(itemPart,itemQuality),
                    # 'total_flag_id': ''
                }  #yapf:disable
                wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()
        wbData.setDecorationData(DataDeal(wb.getData()))
    with WorkBook(tarPath, goodsWb, app) as wb:
        wbData.setGoodsData(DataDeal(wb.getData()))
    printer.printGapTime("宠物相关数据获取完毕，耗时:")

    sys.exit()
    # -----------------------------------mail_info_data.xlsx-------------------------------------------------
    wbName = mailWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for row in itemList[1:]:
            if row[wbData.itemActivityCol] != None:
                itemType = row[wbData.itemTypeCol]
                for petType in range(2):
                    itemId = row[wbData.itemCatCol]
                    if petType == 1:
                        itemId = row[wbData.itemDogCol]
                    for i in range(1, 6):
                        male = i
                        if male == 5:
                            male = 8
                        activityId = row[wbData.itemActivityCol] + (i - 1) + petType * 5
                        reward = wbData.getMailItemInfo(male, itemId, itemType, 1)
                        indexId = row[wbData.itemMailCol] + (i - 1) + petType * 5
                        dataMap = {
                            'mail_id':indexId,
                            'mail_platforms':-1,
                            'mail_sender':'恋语市政府',
                            'mail_from_npc':'',
                            'mail_title':'扭蛋机奖励补发',
                            'mail_desc':'扭蛋机累计奖励活动已结束，请制作人查收尚未领取的奖励',
                            'mail_rewards':reward,
                            'mail_effective':wbData.getMailBeginTime(row[wbData.itemActivityEndTimeCol]),
                            'mail_deadline':wbData.getMailEndTime(row[wbData.itemActivityEndTimeCol]),
                            'mail_condition':activityId,
                            'mail_condition_type':2,
                            'mail_mydaybook':'',
                            # 'total_flag_id': ''
                        }  #yapf:disable
                        wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------activity_info_data.xlsx-------------------------------------------------
    wbName = activityWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for row in itemList[1:]:
            if row[wbData.itemActivityCol] != None:
                for petType in range(2):
                    for i in range(1, 6):
                        male = i
                        if male == 5:
                            male = 8
                        indexId = row[wbData.itemActivityCol] + (i - 1) + petType * 5
                        dataMap = {
                            'activity_id':indexId,
                            'activity_start_time':row[wbData.itemActivityTimeCol],
                            'activity_end_time':row[wbData.itemActivityEndTimeCol],
                            'activity_title':'宠物扭蛋活动',
                            'activity_recommend':2,
                            'activity_weight':9999,
                            'activity_repeat_cnt':1,
                            # 'total_flag_id': ''
                        }  #yapf:disable
                        wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------pet_quest.xlsx-------------------------------------------------
    wbName = questWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for row in itemList[1:]:
            if row[wbData.itemActivityCol] != None:
                itemType = row[wbData.itemTypeCol]
                for petType in range(2):
                    itemId = row[wbData.itemCatCol]
                    if petType == 1:
                        itemId = row[wbData.itemDogCol]
                    for i in range(1, 6):
                        male = i
                        if male == 5:
                            male = 8
                        activityId = row[wbData.itemActivityCol] + (i - 1) + petType * 5
                        reward = wbData.getItemInfo(itemId, itemType, 1)
                        targetPara = wbData.getQuestPara(petType, row[wbData.itemPoolCol])
                        desc = wbData.getQuestDesc(row[wbData.itemActivityTimeCol], row[wbData.itemActivityEndTimeCol])
                        order = wbData.getQuestOrder(activityId, 105, male, 1)
                        indexId = 1050000 + male * 1000 + 100 + order
                        dataMap = {
                            'quest_id':indexId,
                            'activity_id':activityId,
                            'quest_type':105,
                            'role_id':male,
                            'group_id':1,
                            'target_type':119,
                            'target_param':targetPara,
                            'unit_desc':'',
                            'rewards':reward,
                            'desc':desc,
                            'show_order':order,
                            'title':'累计扭蛋',
                            # 'total_flag_id': ''
                        }  #yapf:disable
                        wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------goods_exchange_info_data.xlsx-------------------------------------------------
    wbName = exchangeWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        typeCol = wbData.getColOrder('exchange_type')
        rewardCol = wbData.getColOrder('exchange_rewards')
        for row in itemList[1:]:
            if row[wbData.itemShopCol] != None and row[wbData.itemTypeCol] != '道具':
                itemType = row[wbData.itemTypeCol]
                for petType in range(2):
                    itemId = row[wbData.itemCatCol]
                    if petType == 1:
                        itemId = row[wbData.itemDogCol]
                    for i in range(1, 6):
                        male = i
                        if male == 5:
                            male = 8
                        exchangeType = 90 + male
                        exchangeReward = wbData.getItemInfo(itemId, itemType, 1)
                        price = wbData.getItemPrice(itemId, itemType)
                        label = petType + 1
                        if wbData.getItemPart(itemId, itemType) == '床垫':
                            label = 0
                        goodsType = 3
                        if itemType == '服饰':
                            goodsType = 4
                        goodsRank = wbData.getGoodsRank(itemId, itemType, male)

                        indexId = 0
                        goodsRow = wbData.getRow2(exchangeType, typeCol, exchangeReward, rewardCol)
                        if goodsRow == -1:
                            indexId = wbData.getMaxId('exchange_id', 0, exchangeType, rewardCol) + 1
                        else:
                            indexId = wbData.rngData[goodsRow][0]
                        dataMap = {
                            'exchange_id':indexId,
                            'exchange_type':exchangeType,
                            'exchange_price':price,
                            'exchange_rewards':exchangeReward,
                            'tag':2,
                            'label':label,
                            'exchange_need':'',
                            'exchange_new_flag':0,
                            'goods_rank':goodsRank,
                            'start_time':'',
                            'end_time':'',
                            'exchange_rules':'4,281,1',
                            'goods_type': goodsType,
                            'hide_on_soldout':0
                        }  #yapf:disable
                        wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------gacha_open_info.xlsx-------------------------------------------------
    wbName = openWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')

                startTime = wbData.getParaValue(poolType, '奖池开始时间')
                endTime = wbData.getParaValue(poolType, '奖池结束时间')
                extraReward = wbData.getParaValue(poolType, '十连额外奖励')
                extraRule = wbData.getParaValue(poolType, '额外奖励周期')

                for petType in range(2):
                    refreshId = 0
                    animalStr = ''
                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                        animalStr = '-猫'
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')
                        animalStr = '-狗'
                    for openType in range(2):
                        openNum = 0
                        price = ''
                        remarkStr = ''
                        if openType == 0:
                            openNum = 1
                            remarkStr = wbData.getParaValue(poolType, '奖池类型') + '单抽' + animalStr
                            price = wbData.getParaValue(poolType, '单抽价格')
                        else:
                            openNum = 10
                            remarkStr = wbData.getParaValue(poolType, '奖池类型') + '十连' + animalStr
                            price = wbData.getParaValue(poolType, '十连价格')
                        indexId = wbData.getPoolOpenId(refreshId, openNum)
                        dataMap = {
                            'id':indexId,
                            'refresh_id':refreshId,
                            'open_gacha_num':openNum,
                            'price':price,
                            'extra_reward':GachaDeal.getPoolExtraReward(refreshId,openNum,extraReward),
                            'extra_reward_rule':GachaDeal.getPoolExtraRewardRule(openNum,extraReward,extraRule),
                            'start_time':startTime,
                            'end_time':endTime,
                            'total_flag_id':updateTag
                        }  #yapf:disable
                        remarkMap = {'备注': remarkStr}
                        wbData.setRowData(indexId, dataMap)
                        wbData.setRowData(indexId, remarkMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------gacha_refresh_info.xlsx-------------------------------------------------
    wbName = refreshWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")
        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')

                gachaNum = wbData.getParaValue(poolType, '刷新出蛋个数')
                freeTime = wbData.getParaValue(poolType, '每日免费次数')
                freshPrice = wbData.getParaValue(poolType, '刷新代币')
                increasePrice = wbData.getParaValue(poolType, '刷新递增价格')
                skin = wbData.getParaValue(poolType, '扭蛋机皮肤')
                token = wbData.getParaValue(poolType, '界面显示代币')
                for petType in range(2):
                    refreshId = 0
                    poolStr = wbData.getParaValue(poolType, '奖池类型')
                    animalStr = ''
                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                        animalStr = poolStr + '-猫'
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')
                        animalStr = poolStr + '-狗'

                    indexId = refreshId
                    dataMap = {
                        'id':refreshId,
                        'type':petType + 1,
                        'gacha_num':gachaNum,
                        'refresh_free_time':freeTime,
                        'refresh_price':freshPrice,
                        'refresh_increase_price':increasePrice,
                        'base_reward_id':wbData.getPoolRandomData(refreshId,poolStr,wbData.ballGroupCol),
                        'reward_id_protect_num':wbData.getPoolProtectData(refreshId,poolStr),
                        'random_reward_id':wbData.getPoolRandomData(refreshId,poolStr,wbData.ballRandomCol),
                        'token_item':token,
                        'effect':'effect' if poolStr != '普通' else '',
                        'skin':skin,
                        'total_flag_id':updateTag
                    }  #yapf:disable
                    remarkMap = {'备注': animalStr}
                    wbData.setRowData(indexId, dataMap)
                    wbData.setRowData(indexId, remarkMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------gacha_reward_info.xlsx-------------------------------------------------
    wbName = rewardWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")
        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')
                for petType in range(2):
                    refreshId = 0
                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')

                    for row in ballList[1:]:
                        ballType = row[wbData.ballTypeCol]
                        ballQuality = row[wbData.ballQualityCol]
                        dropTag = wbData.isBallDrop(ballType, row[wbData.ballGroupCol], row[wbData.ballRandomCol],
                                                    row[wbData.ballProtectCol])
                        poolStr = wbData.getParaValue(poolType, '奖池类型')
                        if dropTag == True and row[wbData.ballPoolCol] == poolStr: # yapf:disable
                            ballId = GachaDeal.getBallId(refreshId,ballType,ballQuality) # yapf:disable
                            indexId = ballId * 100
                            dataMap = {
                                'uid':indexId,
                                'id':ballId,
                                'begin_time':'2021/7/1 5:00:00',
                                'end_time':'2038/1/1 4:59:59',
                                'reward':toStr(GachaDeal.getBallReward(ballId))+ ',103,1',
                                'type':GachaDeal.getBallSkin(ballType,ballQuality),
                                'special_reward_item':toStr(GachaDeal.getBallReward(ballId,1))+ ',103,1' if row[wbData.dropProtectCol] != None else '',
                                'special_reward_protect_num':row[wbData.dropProtectCol],
                                'refresh_hint':GachaDeal.getRefreshHint(poolStr,ballQuality),
                                'delay_open':1 if '神秘' in ballType else '',
                                'effect':GachaDeal.getBallEffect(ballType,ballQuality),
                                'total_flag_id':updateTag
                            }  #yapf:disable
                            remarkMap = {'序号': 0}
                            wbData.setRowData(indexId, dataMap)
                            wbData.setRowData(indexId, remarkMap, defRow=3)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------gacha_preview.xlsx-------------------------------------------------
    wbName = previewWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")
        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())

        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')
                for petType in range(2):
                    refreshId = 0
                    itemCol = 0
                    itemOrder = 0

                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                        itemCol = wbData.itemCatCol
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')
                        itemCol = wbData.itemDogCol

                    for row in itemList[1:]:
                        itemOrder += 1
                        itemId = row[itemCol]
                        itemType = row[wbData.itemTypeCol]
                        itemTypeId = wbData.getItemTypeId(itemType, itemId)
                        itemQuality = row[wbData.itemQualityCol]
                        if itemQuality == None:
                            itemQuality = wbData.getItemQuality(itemId, itemType)
                        if wbData.isItemDrop(refreshId, row[wbData.itemPoolCol]):
                            uId = wbData.getPreviewUid(refreshId, itemOrder, itemQuality)
                            indexId = wbData.getPreviewIndex(refreshId, itemId, itemTypeId)  # 查询是否存在数据
                            if indexId == -1:
                                indexId = uId
                            dataMap = {
                                'uid':uId,
                                'refresh_id':refreshId,
                                'reward_type':GachaDeal.getPreviewType(itemType),
                                'beging_time':'2021/7/1 5:00:00',
                                'end_time':'2038/1/1 4:59:59',
                                'reward_value':toStr(itemId) + ',' + toStr(itemTypeId) + ',1',
                                'up_type':'',
                                'up_beging_time':'',
                                'up_end_time':'',
                                'total_flag_id':updateTag
                            }  #yapf:disable
                            wbData.setRowData(indexId, dataMap)

                    for ticket in range(2):
                        itemOrder += 1
                        itemQuality = -1
                        itemId = 6005
                        itemType = '服饰'
                        itemTypeId = 374
                        if ticket == 1:
                            itemId = 6004
                            itemTypeId = 373
                            itemType = '家具'

                        uId = wbData.getPreviewUid(refreshId, itemOrder, itemQuality)
                        indexId = wbData.getPreviewIndex(refreshId, itemId, itemTypeId)  # 查询是否存在数据
                        if indexId == -1:
                            indexId = uId
                        dataMap = {
                            'uid':uId,
                            'refresh_id':refreshId,
                            'reward_type':GachaDeal.getPreviewType(itemType),
                            'beging_time':'2021/7/1 5:00:00',
                            'end_time':'2038/1/1 4:59:59',
                            'reward_value':toStr(itemId) + ',375,1',
                            'up_type':'',
                            'up_beging_time':'',
                            'up_end_time':'',
                            'total_flag_id':updateTag
                        }  #yapf:disable
                        wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------goods_drop_pet_group.xlsx-------------------------------------------------
    wbName = groupWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())
        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')
                isPoolUnique = wbData.getParaValue(poolType, '是否去重')
                if isPoolUnique != 1:
                    continue
                for petType in range(2):
                    refreshId = 0
                    ballStr = ''
                    idCol = 0
                    itemOrder = 1
                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                        ballStr = '-cat' + wbData.getParaValue(poolType, '奖池类型') + '池'
                        idCol = wbData.itemCatCol
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')
                        ballStr = '-dog' + wbData.getParaValue(poolType, '奖池类型') + '池'
                        idCol = wbData.itemDogCol

                    # 统计数量计算权重
                    configMap = {}
                    for row in itemList[1:]:
                        itemInPool = GachaDeal.isItemDrop(refreshId, row[wbData.itemPoolCol])
                        itemType = row[wbData.itemTypeCol]
                        if itemInPool and itemType != '道具':
                            itemId = row[idCol]
                            itemType = row[wbData.itemTypeCol]
                            itemStr = toStr(itemId) + ',' + toStr(wbData.getItemTypeId(itemType, itemId)) + ',1'
                            quality = row[wbData.itemQualityCol]
                            if quality == None:
                                quality = wbData.getItemQuality(itemId, itemType)
                            configId = GachaDeal.getUniqueRewardId(refreshId, itemType, quality)
                            if configId in configMap:
                                configMap[configId] += 1
                            else:
                                configMap[configId] = 1

                    countMap = {}
                    for row in itemList[1:]:
                        itemInPool = GachaDeal.isItemDrop(refreshId, row[wbData.itemPoolCol])
                        itemType = row[wbData.itemTypeCol]
                        if itemInPool and itemType != '道具':
                            itemId = row[idCol]
                            itemStr = toStr(itemId) + ',' + toStr(wbData.getItemTypeId(itemType, itemId)) + ',1'
                            quality = row[wbData.itemQualityCol]
                            if quality == None:
                                quality = wbData.getItemQuality(itemId, itemType)
                            configId = GachaDeal.getUniqueRewardId(refreshId, itemType, quality)
                            weight = int(10000 / configMap[configId])
                            if configId in countMap:
                                countMap[configId] += 1
                                itemOrder = countMap[configId]
                                if countMap[configId] == configMap[configId]:
                                    weight = int(10000 - weight * (configMap[configId] - 1))
                            else:
                                countMap[configId] = 1
                                itemOrder = 1

                            for tag in range(2):
                                indexId = wbData.getGroupUid(configId, itemStr, tag, itemOrder)
                                dataMap = {
                                    'uid':indexId,
                                    'config_id':configId,
                                    'name':'去重掉落' + toStr(quality) + ballStr,
                                    'desc':'',
                                    'icon':'',
                                    'drop_rule':itemStr if tag == 0 else '',
                                    'true_or_false':tag,
                                    'random_weight':weight,
                                    'rewards':itemStr,
                                    'skip':'',
                                    'total_flag_id':updateTag
                                }  #yapf:disable
                                wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()

    # -----------------------------------goods_weightset_data.xlsx-------------------------------------------------
    wbName = weightsetWb
    printer.setStartTime("开始打开表格:" + wbName, 'green')
    with WorkBook(tarPath, wbName, app) as wb:
        printer.printGapTime("表格打开完毕，耗时:")

        printer.setStartTime("开始处理数据数据...")
        wbData.setData(wb.getData())
        for poolType in wbData.poolUpdateMap:
            if wbData.poolUpdateMap[poolType]:
                updateTag = wbData.getParaValue(poolType, '更新标记')
                for petType in range(2):
                    refreshId = 0
                    ballStr = ''
                    rewardDetail = ''
                    protectDetail = ''
                    extendTag = False
                    if petType == 0:
                        refreshId = wbData.getParaValue(poolType, '猫refresh奖池')
                        ballStr = '-cat' + wbData.getParaValue(poolType, '奖池类型')
                    else:
                        refreshId = wbData.getParaValue(poolType, '狗refresh奖池')
                        ballStr = '-dog' + wbData.getParaValue(poolType, '奖池类型')

                    for row in ballList[1:]:
                        ballDrop = wbData.isBallDrop(row[wbData.ballTypeCol],row[wbData.ballGroupCol],row[wbData.ballRandomCol],row[wbData.ballProtectCol]) #yapf:disable
                        if row[wbData.ballPoolCol] == poolType and ballDrop:
                            ballType = row[wbData.ballTypeCol]
                            ballQuality = row[wbData.ballQualityCol]
                            ballId = GachaDeal.getBallId(refreshId,ballType,ballQuality) # yapf:disable
                            # 常规掉落
                            if extendTag and '神秘' in ballType:
                                pass
                            else:
                                rewardDetail = ''
                                protectDetail = ''

                            if '神秘' in ballType:
                                extendTag = True
                            else:
                                extendTag = False

                            indexId = GachaDeal.getBallReward(ballId)
                            rewardDetail = wbData.getWeightReward(rewardDetail, petType, refreshId, ballType,
                                                                  row[wbData.dropUniqueCol], row[wbData.dropTicketCol],
                                                                  row[wbData.dropStar0Col], row[wbData.dropStar1Col],
                                                                  row[wbData.dropStar2Col], row[wbData.dropStar3Col])
                            dataMap = {
                                'goods_weightset_id':indexId,
                                'goods_weightset_detail':rewardDetail,
                                'goods_weightset_name':ballType + '蛋' + toStr(ballQuality) + ballStr,
                                'goods_weightset_desc':'',
                                'goods_weightset_icon':'',
                                'total_flag_id':updateTag
                            }  #yapf:disable
                            wbData.setRowData(indexId, dataMap)

                            if row[wbData.dropProtectCol] != None:
                                # 保底掉落
                                indexId = GachaDeal.getBallReward(ballId, 1)
                                protectDetail = wbData.getProtectWeightReward(protectDetail,petType,refreshId,ballType,
                                                                              row[wbData.dropUniqueCol], row[wbData.dropProtectQualityCol]) # yapf:disable
                                dataMap = {
                                    'goods_weightset_id':indexId,
                                    'goods_weightset_detail':protectDetail,
                                    'goods_weightset_name':ballType + '蛋' + toStr(ballQuality) + ballStr + '-保底',
                                    'goods_weightset_desc':'',
                                    'goods_weightset_icon':'',
                                    'total_flag_id':updateTag
                                }  #yapf:disable
                                wbData.setRowData(indexId, dataMap)
        printer.printGapTime("数据处理完毕，耗时:")

        printer.setStartTime("开始检查单元格格式:" + wbName)
        wbData.strCheck()
        printer.printGapTime("单元格格式检查完毕，耗时:", skipLines=1)

        wb.setData(wbData.getData())
        wb.save()
