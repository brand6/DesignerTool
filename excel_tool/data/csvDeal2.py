import csv

filePath = 'C:\\Users\\PC\\Downloads\\'
qrFile = 'qrcode'
seatFile = '音乐会名单-1123终版'
combileFile = 'qrSeat'


# 合表座位和二维码数据
def combimeQrSeat(qrFile, seatFile, combileFile):
    with open(filePath + qrFile + '.csv', 'r', encoding='UTF-8') as file1:
        with open(filePath + seatFile + '.csv', 'r', encoding='UTF-8') as file2:
            with open(filePath + combileFile + '.csv', 'w', encoding='UTF-8', newline='') as outFile:
                csvReader1 = csv.reader(file1)
                csvReader2 = csv.reader(file2)
                csvWriter = csv.writer(outFile)
                outList = []
                for line1 in csvReader1:
                    for line2 in csvReader2:
                        if line1[0] == line2[0]:
                            line1.append(line2[1])
                            outList.append(line1)
                            break
                csvWriter.writerows(outList)


combimeQrSeat(qrFile, seatFile, combileFile)
