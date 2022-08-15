import csv

filePath = 'C:\\Users\\PC\\Downloads\\sql_data\\'
chargeFile = 'charge'
giftFile = 'gifts'
experimentsFile = 'experiments'
goodsFile = 'goods'
cardsFile = 'cards'


# 获取有效充值用户
def getChargeUser(chargeMoney):
    with open(filePath + chargeFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my_' + chargeFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
            csvReader = csv.reader(file)
            csvWriter = csv.writer(outFile)
            outList = []
            for line in csvReader:
                if line[1] != '' and toInt(line[1]) > chargeMoney:
                    outList.append(line)
            outList = sorted(
                outList,
                key=(lambda x: toInt(x[1])),
                reverse=True,
            )
            csvWriter.writerows(outList)


# 将充值表和心灵试炼进度组合
def combineUserExperiments(chargeFile, dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my_' + chargeFile + '.csv', 'r', encoding='UTF-8') as inFile:
            with open(filePath + chargeFile + '&' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                expList = csv.reader(file)
                userList = csv.reader(inFile)
                csvWriter = csv.writer(outFile)
                resultList = []
                expMap = {}

                replaceMap = {
                    '11': '李泽言-决策','12': '李泽言-创造','13': '李泽言-亲和','14': '李泽言-行动',
                    '21': '许墨-决策','22': '许墨-创造','23': '许墨-亲和','24': '许墨-行动',
                    '31': '周棋洛-决策','32': '周棋洛-创造','33': '周棋洛-亲和','34': '周棋洛-行动',
                    '41': '白起-决策','42': '白起-创造','43': '白起-亲和','44': '白起-行动'
                } # yapf:disable

                for expLine in expList:
                    if expMap.get(expLine[0]):
                        expMap[expLine[0]].extend(expLine[1:])
                    else:
                        expMap[expLine[0]] = expLine[1:]
                for userLine in userList:
                    if expMap.get(userLine[0]):
                        expRow = expMap.get(userLine[0])
                        for l in expRow:
                            if toInt(l) > 90000:
                                l = l[-2:]
                            else:
                                l = replaceMap.get(l)
                            userLine.append(l)
                        userLine = sortRow(userLine, 4, 2, sortNum=1, reverse=True)

                        resultList.append(userLine)
                csvWriter.writerows(resultList)


# 分析心灵试炼数据
def analyseUserExperiments(dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'analyse_' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
            expList = csv.reader(file)
            csvWriter = csv.writer(outFile)
            expMap = {}
            countMap = {}
            outList = []
            title = True
            for line in expList:
                if title:
                    title = False
                    outList.append([line[0]])
                    outList.append(['心灵试炼最高关卡', '玩家数', '玩家总数', '占比'])
                else:
                    for i in range(len(line[4:])):
                        if i % 2 == 0:
                            mapName = line[4 + i] + line[5 + i]
                            if expMap.get(mapName):
                                expMap[mapName] += 1
                            else:
                                expMap[mapName] = 1

                            if countMap.get(line[4 + i]):
                                countMap[line[4 + i]] += 1
                            else:
                                countMap[line[4 + i]] = 1
            for k in expMap:
                if countMap.get(k[:-2]):
                    totalNum = countMap.get(k[:-2])
                else:
                    totalNum = countMap.get(k[:-1])
                if totalNum == None:
                    totalNum = 1
                outList.append([k, expMap.get(k), totalNum, str(round(expMap.get(k) / totalNum, 5) * 100) + '%'])
            outList[2:] = sorted(
                outList[2:],
                key=(lambda x: x[0]),
                reverse=True,
            )
            csvWriter.writerows(outList)


# 将充值表和卡牌表组合
def combineUserCard2(chargeFile, dataFile, spotNum=0):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my2_' + chargeFile + '.csv', 'r', encoding='UTF-8') as inFile:
            with open(filePath + chargeFile + '&' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                expList = csv.reader(file)
                userList = csv.reader(inFile)
                csvWriter = csv.writer(outFile)
                resultList = []
                countMap = {}
                cardMap = {}
                title = True
                roleMap = {'1': '李泽言', '2': '许墨', '3': '周棋洛', '4': '白起', '8': '凌肖'}
                typeMap = {'5': 'SSR', '6': 'SP', '7': 'ER', '8': 'SER'}
                orderMap = {
                    '61': 0,'62': 1,'63': 2,'64': 3,
                    '51': 4,'52': 5,'53': 6,'54': 7,'58': 8,
                    '81': 9,'82': 10,'83': 11,'84': 12,'88': 13,
                    '71': 14,'72': 15,'73': 16,'74': 17,'78': 18,
                } # yapf:disable

                for expLine in expList:
                    if title:
                        title = False
                        countMap[expLine[0]] = [
                            '李泽言SP数量', '许墨SP数量', '周棋洛SP数量', '白起SP数量',
                            '李泽言SSR数量', '许墨SSR数量', '周棋洛SSR数量', '白起SSR数量', '凌肖SSR数量',
                            '李泽言SER数量', '许墨SER数量', '周棋洛SER数量', '白起SER数量', '凌肖SER数量',
                            '李泽言ER数量', '许墨ER数量', '周棋洛ER数量', '白起ER数量','凌肖ER数量'
                        ] # yapf:disable
                    elif expLine[1] != '':
                        if countMap.get(expLine[0]):
                            countMap[expLine[0]][orderMap.get(expLine[1][:2])] += 1
                        else:
                            countMap[expLine[0]] = [0] * 19
                            countMap[expLine[0]][orderMap.get(expLine[1][:2])] = 1

                        if expLine[4] == '' or expLine[4] == '37':
                            expLine[4] = '0'
                        cardType = roleMap.get(expLine[1][1]) + typeMap.get(expLine[1][0]) + expLine[4]
                        if cardMap.get(expLine[0]):
                            if cardMap[expLine[0]].get(cardType):
                                cardMap[expLine[0]][cardType] += 1
                            else:
                                cardMap[expLine[0]][cardType] = 1
                        else:
                            cardMap[expLine[0]] = {cardType: 1}

                for userLine in userList:
                    if countMap.get(userLine[0]):
                        userLine.extend(countMap[userLine[0]])
                    if cardMap.get(userLine[0]):
                        lineDict = cardMap.get(userLine[0])
                        lineDict = sorted(
                            lineDict,
                            key=(lambda x: toInt(x[-2:], 1)),
                            reverse=True,
                        )
                        for k in lineDict:
                            userLine.extend([k, cardMap.get(userLine[0]).get(k)])

                    resultList.append(userLine)
                csvWriter.writerows(resultList)


# 将充值表和卡牌表组合
def combineUserCard(chargeFile, dataFile, spotNum=0):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my_' + chargeFile + '.csv', 'r', encoding='UTF-8') as inFile:
            with open(filePath + chargeFile + '&' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                expList = csv.reader(file)
                userList = csv.reader(inFile)
                csvWriter = csv.writer(outFile)
                resultList = []
                countMap = {}
                title = True
                roleMap = {'1': 0, '2': 1, '3': 2, '4': 3, '8': 4}
                spotList = [36, 31, 21, 11, 1, 0]

                for expLine in expList:
                    if title:
                        title = False
                        countMap[expLine[0]] = ['李泽言36','许墨36','周棋洛36','白起36','凌肖36',
                                                '李泽言31','许墨31','周棋洛31','白起31','凌肖31',
                                                '李泽言21','许墨21','周棋洛21','白起21','凌肖21',
                                                '李泽言11','许墨11','周棋洛11','白起11','凌肖11',
                                                '李泽言1','许墨1','周棋洛1','白起1','凌肖1',
                                                '李泽言0','许墨0','周棋洛0','白起0','凌肖0',] # yapf:disable
                    elif expLine[1] != '' and expLine[1][0] != '7':
                        if not countMap.get(expLine[0]):
                            countMap[expLine[0]] = [0] * 30
                        role = roleMap.get(expLine[1][1])
                        spotNum = toInt(expLine[4])
                        if expLine[1][0] == '8':
                            spotNum += 10
                        for i in range(len(spotList)):
                            if spotNum >= spotList[i]:
                                spotNum = i
                                break
                        countMap[expLine[0]][spotNum * 5 + role] += 1
                for userLine in userList:
                    if countMap.get(userLine[0]):
                        userLine.extend(countMap[userLine[0]])
                    resultList.append(userLine)
                csvWriter.writerows(resultList)


# 统计不同升华进度的卡牌数量
def analyseUserCard(dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'analyse_' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
            dataList = csv.reader(file)
            csvWriter = csv.writer(outFile)
            resultList = []
            numList = [40, 30, 20, 15, 12, 9, 6, 4, 3, 2, 1, 0]
            spotList = [36, 31, 21, 11, 1, 0]
            strList = ['[36]', '[31,35]', '[21,30]', '[11,20]', '[1,10]', '[0]']
            lineMap = {}
            for n in numList:
                for s in range(len(spotList)):
                    s = strList[s]
                    lineMap[s + str(n)] = [0, 0, 0, 0, 0]
            title = True
            for lineList in dataList:
                if title:
                    title = False
                    resultList.append([lineList[0]])
                else:
                    for i in range(len(lineList[4:])):
                        cardNum = toInt(lineList[4 + i])
                        for n in numList:
                            if cardNum >= n:
                                k = str(strList[int(i / 5)]) + str(n)
                                lineMap[k][i % 5] += 1
                                break
            for l in lineMap:
                temp = [l]
                for s in lineMap[l]:
                    temp.append(s)
                resultList.append(temp)
            csvWriter.writerows(resultList)


# 将充值表和物品表组合
def combineUserGoods(chargeFile, dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my_' + chargeFile + '.csv', 'r', encoding='UTF-8') as inFile:
            with open(filePath + chargeFile + '&' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                expList = csv.reader(file)
                userList = csv.reader(inFile)
                csvWriter = csv.writer(outFile)
                title = True
                resultList = []
                expMap = {}
                nameList = [
                    '心绪之石','心绪之花','特级助力','高级助力','中级助力','初级助力','回忆星辰','心乐之章',
                    '心绪之石比例','心绪之花比例','特级助力比例','高级助力比例','中级助力比例','初级助力比例','回忆星辰比例',
                ] # yapf:disable
                costList = [104235, 1003, 350, 470, 290, 140, 6]
                for expLine in expList:
                    if title:
                        title = False
                        expMap[expLine[0]] = nameList
                    else:
                        stoneNum = toInt(expLine[1])
                        flowerNum = toInt(expLine[2])
                        item1Num = toInt(expLine[6]) + toInt(expLine[10]) + toInt(expLine[14]) + toInt(expLine[18])
                        item2Num = toInt(expLine[5]) + toInt(expLine[9]) + toInt(expLine[13]) + toInt(expLine[17])
                        item3Num = toInt(expLine[4]) + toInt(expLine[8]) + toInt(expLine[12]) + toInt(expLine[16])
                        item4Num = toInt(expLine[3]) + toInt(expLine[7]) + toInt(expLine[11]) + toInt(expLine[15])
                        starNum = toInt(expLine[19])
                        if flowerNum < 60180:
                            expMap[expLine[0]] = [stoneNum]
                            expMap[expLine[0]].append(flowerNum)
                            expMap[expLine[0]].append(item1Num)
                            expMap[expLine[0]].append(item2Num)
                            expMap[expLine[0]].append(item3Num)
                            expMap[expLine[0]].append(item4Num)
                            expMap[expLine[0]].append(starNum)
                            expMap[expLine[0]].append(expLine[22])
                            expMap[expLine[0]].append(round(stoneNum / costList[0], 2))
                            expMap[expLine[0]].append(round(flowerNum / costList[1], 2))
                            expMap[expLine[0]].append(round(item1Num / costList[2], 2))
                            expMap[expLine[0]].append(round(item2Num / costList[3], 2))
                            expMap[expLine[0]].append(round(item3Num / costList[4], 2))
                            expMap[expLine[0]].append(round(item4Num / costList[5], 2))
                            expMap[expLine[0]].append(round(starNum / costList[6], 2))
                for userLine in userList:
                    if expMap.get(userLine[0]):
                        userLine.extend(expMap[userLine[0]])
                        resultList.append(userLine)
                csvWriter.writerows(resultList)


# 统计物品数量分布
def analyseUserGoods(dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'analyse_' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
            goodsList = csv.reader(file)
            csvWriter = csv.writer(outFile)
            outList = []
            numList = [60, 40, 20, 10, 8, 6, 4, 3, 2, 1, 0.8, 0.6, 0.4, 0.3, 0.2, 0.1, 0]
            mapList = []
            nameList = []
            title = True
            for goodsLine in goodsList:
                if title:
                    title = False
                    outList.append([goodsLine[0]])
                    outList.append(['数量区间', '玩家数量'])
                    nameList = goodsLine[12:]
                    while len(mapList) < len(nameList):
                        mapList.append({})
                else:
                    for i in range(len(goodsLine[12:])):
                        goodsNum = float(goodsLine[12 + i])
                        for n in range(len(numList)):
                            if goodsNum > numList[n]:
                                if mapList[i].get(numList[n - 1]):
                                    mapList[i][numList[n - 1]] += 1
                                else:
                                    mapList[i][numList[n - 1]] = 1
                                break
                        else:
                            if mapList[i].get(numList[n]):
                                mapList[i][numList[n]] += 1
                            else:
                                mapList[i][numList[n]] = 1
            for m in range(len(mapList)):
                d = sorted(mapList[m], reverse=True)
                for k in range(len(d)):
                    if k == len(d) - 1:
                        outList.append([nameList[m] + '0', mapList[m][d[k]]])
                    else:
                        outList.append([nameList[m] + '(' + str(d[k + 1]) + ',' + str(d[k]) + ']', mapList[m][d[k]]])

            csvWriter.writerows(outList)


# 将充值表和礼包表组合
def combineUserGift(chargeFile, dataFile):
    with open(filePath + dataFile + '.csv', 'r', encoding='UTF-8') as file:
        with open(filePath + 'my_' + chargeFile + '.csv', 'r', encoding='UTF-8') as inFile:
            with open(filePath + chargeFile + '&' + dataFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                expList = csv.reader(file)
                userList = csv.reader(inFile)
                csvWriter = csv.writer(outFile)
                resultList = []
                expMap = {}

                for expLine in expList:
                    if expMap.get(expLine[0]):
                        expMap[expLine[0]].extend(expLine[1:])
                    else:
                        expMap[expLine[0]] = expLine[1:]
                for userLine in userList:
                    if expMap.get(userLine[0]):
                        userLine.extend(expMap[userLine[0]])
                        userLine = sortRow(userLine, 4, 2, sortNum=1, reverse=True)
                        resultList.append(userLine)
                csvWriter.writerows(resultList)


def sortRow(listData, startNum, splitNum, sortNum=0, reverse=False):
    returnList = listData[:startNum]
    for l in range(len(listData[startNum:])):
        if l % splitNum == sortNum:
            for r in range(len(returnList[startNum:])):
                if r % splitNum == sortNum:
                    if reverse == True:
                        if toInt(listData[startNum + l]) > toInt(returnList[startNum + r]):
                            for s in range(splitNum):
                                returnList.insert(startNum + r - sortNum + s, listData[startNum + l - sortNum + s])
                            break
                    else:
                        if toInt(listData[startNum + l]) < toInt(returnList[startNum + r]):
                            for s in range(splitNum):
                                returnList.insert(startNum + r - sortNum + s, listData[startNum + l - sortNum + s])
                            break
            else:
                for s in range(splitNum):
                    returnList.append(listData[startNum + l - sortNum + s])
    return returnList


def toInt(value, check=0):
    try:
        return int(value)
    except:
        if value == '':
            return 0
        else:
            if check != 0:
                return int(value[-check])
            else:
                return 9999999999


#getChargeUser(999)

#combineUserExperiments(chargeFile, experimentsFile)
#analyseUserExperiments(chargeFile + '&' + experimentsFile)

#combineUserGift(chargeFile, giftFile)

#combineUserGoods(chargeFile, goodsFile)
#analyseUserGoods(chargeFile + '&' + goodsFile)

#combineUserCard2(chargeFile, cardsFile, 0)
#combineUserCard(chargeFile, cardsFile, 0)
analyseUserCard(chargeFile + '&' + cardsFile)
