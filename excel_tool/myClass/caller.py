import os


class Caller:

    # 调用bat文件
    @classmethod
    def callBat(cls, filePath, fileName):
        os.chdir(filePath)
        if fileName[-3:] != "bat":
            fileName = fileName + ".bat"

        os.system(r'start cmd.exe /k' + filePath + "\\" + fileName)