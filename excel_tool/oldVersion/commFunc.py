import xlwings as xl
import copy
import re
import os
import colorama


# 关闭workbook
def closeWb(wbName):
    if wbName[-5:] != ".xlsx":
        wbName = wbName + ".xlsx"

    wbList = xl.books
    for wb in wbList:
        if wb.name == wbName:
            wb.close()


# 获取表格数据
def getShtData(foldPath, wbName):
    if wbName[-5:] != ".xlsx":
        wbName = wbName + ".xlsx"

    closeWb(wbName)
    app = xl.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(foldPath + "\\" + wbName, 0)
    shtData = wb.sheets[0].used_range.value
    wb.close()
    app.quit()
    return shtData


# 获取表格对象
def getApp(foldPath, wbName):
    if wbName[-5:] != ".xlsx":
        wbName = wbName + ".xlsx"

    closeWb(wbName)
    app = xl.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    return app


# 数据写入表格
def writeData(tarWb, tarRng):
    if tarWb.sheets[0].api.AutoFilterMode == True:
        tarWb.sheets[0].api.ShowAllData()
    tarWb.sheets[0].range('A1').value = tarRng
    tarWb.save()
    tarWb.close()


# 处理单元格格式
def strCheck(dataRng):
    for col in range(len(dataRng[0])):
        if dataRng[0][col] != None:
            for row in range(len(dataRng)):
                if dataRng[row][col] != None:
                    if isNumber(dataRng[row][col]) == True:
                        pass
                    # 处理带逗号和冒号的数字
                    elif re.match(r"[^']\d*[:|,]", dataRng[row][col]) != None:
                        dataRng[row][col] = "'" + dataRng[row][col][0:]
                    # 处理'开头的文本
                    elif re.match(r"'[^']", dataRng[row][col]) != None:
                        dataRng[row][col] = "'" + dataRng[row][col][0:]


# 获得复制对应的列
def getCopyColumn(oriRng, tarRng):
    copyList = []
    for i in range(len(tarRng[0])):
        if tarRng[0][i] != None:
            for j in range(len(oriRng[0])):
                if tarRng[0][i] == oriRng[0][j]:
                    copyList.append(j)
                    break
            else:
                copyList.append(-1)
        elif tarRng[1][i] != None:
            for j in range(len(oriRng[1])):
                if tarRng[1][i] == oriRng[1][j]:
                    copyList.append(j)
                    break
            else:
                copyList.append(-1)
    return copyList


# 复制数据
def copyData(copyId, oriRng, tarRng, copyColList, checkCol, checkRow=0):
    if type(checkCol).__name__ == 'list':
        oriCol = []
        for col in checkCol:
            oriCol.append(getColOrder(oriRng, col, checkRow))

        oriRows = getRowsInRange(copyId, oriRng, oriCol[0])
        if len(checkCol) == 2:
            for r in oriRows:
                tarRow = getRowInRange_2(tarRng, copyId, checkCol[0], oriRng[r][oriCol[1]], checkCol[1])
                insertData(copyId, oriRng, tarRng, r, tarRow, copyColList, checkCol[0], checkRow)

        elif len(checkCol) == 3:
            for r in oriRows:
                tarRow = getRowInRange_3(tarRng, copyId, checkCol[0], oriRng[r][oriCol[1]], checkCol[1], oriRng[r][oriCol[2]], checkCol[2])
                insertData(copyId, oriRng, tarRng, r, tarRow, copyColList, checkCol[0], checkRow)
    else:
        oriRow = getRowInRange(copyId, oriRng, checkCol, checkRow)
        tarRow = getRowInRange(copyId, tarRng, checkCol, checkRow)
        insertData(copyId, oriRng, tarRng, oriRow, tarRow, copyColList, checkCol, checkRow)


# 数据插入处理
def insertData(copyId, oriRng, tarRng, oriRow, tarRow, copyColList, checkCol, checkRow=0):
    if tarRow == -1:
        tarRow = getInsertRowInRange(copyId, tarRng, checkCol, checkRow)
        tarRng.insert(tarRow, copy.deepcopy(tarRng[-1]))

    for i in range(len(copyColList)):
        if copyColList[i] != -1:
            tarRng[tarRow][i] = oriRng[oriRow][copyColList[i]]


# 获得插入的位置行
def getInsertRowInRange(checkId, checkRng, checkCol=0, checkRow=0):
    checkCol = getColOrder(checkRng, checkCol, checkRow)
    checkId = float(checkId)

    for row in range(len(checkRng) - 1, 0, -1):
        if checkRng[row][checkCol] != None:
            if isNumber(checkRng[row][checkCol]) == False or float(checkRng[row][checkCol]) <= checkId:
                return row + 1
    else:
        return -1


# 转为字符串
def toStr(content):
    if isNumber(content) == True:
        if float(content) - int(content) == 0:
            return str(int(content))
        else:
            return str(content)
    else:
        return str(content)


# 获得数据对应行
def getRowInRange(checkId, checkRng, checkCol=0, checkRow=0):
    checkCol = getColOrder(checkRng, checkCol, checkRow)
    checkId = toStr(checkId)

    for row in range(len(checkRng)):
        if toStr(checkRng[row][checkCol]) == checkId:
            return row
    else:
        return -1


# 获得数据对应的所有行
def getRowsInRange(checkId, checkRng, checkCol=0, checkRow=0):
    checkCol = getColOrder(checkRng, checkCol, checkRow)
    returnList = []
    checkId = toStr(checkId)

    for row in range(len(checkRng)):
        if toStr(checkRng[row][checkCol]) == checkId:
            returnList.append(row)

    return returnList


# 根据两个参数获得行
def getRowInRange_2(
    checkRng,
    checkId,
    checkCol,
    checkId2,
    checkCol2,
    checkRow=0,
):
    checkCol = getColOrder(checkRng, checkCol, checkRow)
    checkCol2 = getColOrder(checkRng, checkCol2, checkRow)
    checkId = toStr(checkId)
    checkId2 = toStr(checkId2)

    for row in range(len(checkRng)):
        if toStr(checkRng[row][checkCol]) == checkId and toStr(checkRng[row][checkCol2]) == checkId2:
            return row
    else:
        return -1


# 根据三个参数获得行
def getRowInRange_3(
    checkRng,
    checkId,
    checkCol,
    checkId2,
    checkCol2,
    checkId3,
    checkCol3,
    checkRow=0,
):
    checkCol = getColOrder(checkRng, checkCol, checkRow)
    checkCol2 = getColOrder(checkRng, checkCol2, checkRow)
    checkCol3 = getColOrder(checkRng, checkCol3, checkRow)
    checkId = toStr(checkId)
    checkId2 = toStr(checkId2)
    checkId3 = toStr(checkId3)

    for row in range(len(checkRng)):
        if toStr(checkRng[row][checkCol]) == checkId and toStr(checkRng[row][checkCol2]) == checkId2 and toStr(checkRng[row][checkCol3]) == checkId3:
            return row
    else:
        return -1


# 根据列名获得列的位置
def getColOrder(checkRng, checkCol, checkRow=0):
    if isNumber(checkCol) == True:
        return int(checkCol)
    else:
        for v in range(len(checkRng[checkRow])):
            if checkCol == checkRng[checkRow][v]:
                return v
        else:
            return -1


# 判断是否数字
def isNumber(checkValue):
    try:
        int(checkValue)
        return True
    except:
        return False


# 调用bat文件
def callBat(filePath, fileName):
    os.chdir(filePath)
    if fileName[-3:] != "bat":
        fileName = fileName + ".bat"

    os.system(r'start cmd.exe /k' + filePath + "\\" + fileName)


# 打印红色
def printColor(content, colorStr='white', wrap=False):
    colorama.init(autoreset=True)
    colorDict = {'red': '31m', 'green': '32m', 'yellow': '33m', 'blue': '34m', 'white': '37m'}

    if colorStr in colorDict:
        print("\033[1;" + colorDict[colorStr] + content + "\033[0m ")
    else:
        print(content)

    if wrap == True:
        print("")


# 打印时间
def printTime(desc, time, precision=2, colorStr="white", wrap=False):
    time = round(time, precision)
    printColor(desc + str(time) + '秒', colorStr, wrap)
