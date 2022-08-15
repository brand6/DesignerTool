import xlwings as xl
from time import time
from commFunc import getShtData, getApp, getCopyColumn, copyData, strCheck, writeData, callBat, printColor, printTime

dataSht = xl.sheets.active
oriPath = dataSht['A2'].value
tarPath = dataSht['A5'].value
batName = dataSht['A12'].value
dataRng = dataSht['D2:P2'].expand('down').value

for rowList in dataRng:
    hasData = False
    wbName = rowList[0] + ".xlsx"
    for cell in rowList[3:]:
        if cell != None:
            hasData = True
            break
    if hasData:
        printColor("开始打开来源表格:" + rowList[0], 'green')
        startTime = time()
        oriRng = getShtData(oriPath, wbName)
        oriTime = time() - startTime
        printTime("来源表格打开完毕，耗时:", oriTime, 2, wrap=True)

        printColor("开始打开目标表格:" + rowList[0])
        startTime = time()
        app = getApp(tarPath, wbName)
        tarWb = app.books.open(tarPath + "\\" + wbName, 0)
        tarRng = tarWb.sheets[0].used_range.value
        tarTime = time() - startTime
        if abs(tarTime - oriTime) < 1:
            printTime("目标表格打开完毕，耗时:", tarTime, 2, wrap=True)
        else:
            printTime("目标表格打开完毕，耗时:", tarTime, 2, 'yellow', True)

        printColor("开始同步数据:" + rowList[0])
        startTime = time()
        copyColList = getCopyColumn(oriRng, tarRng)
        if '|' in str(rowList[1]):
            checkCol = rowList[1].rsplit('|')
        else:
            checkCol = rowList[1]

        for cell in rowList[3:]:
            if cell != None:
                if '|' in str(cell):
                    valueList = cell.rsplit('|')
                    for value in valueList:
                        copyData(value, oriRng, tarRng, copyColList, checkCol)
                else:
                    copyData(cell, oriRng, tarRng, copyColList, checkCol)
        printTime("数据同步完毕，耗时:", time() - startTime, wrap=True)

        printColor("开始检查单元格格式:" + rowList[0])
        startTime = time()
        strCheck(tarRng)
        printTime("单元格格式检查完毕，耗时:", time() - startTime, wrap=True)

        writeData(tarWb, tarRng)
        app.display_alerts = True
        app.screen_updating = True
        app.quit()

if batName == None:
    callBat(tarPath, "1.一键处理表格.bat")
else:
    callBat(tarPath, batName)
