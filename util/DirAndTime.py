import time ,os
from datetime import datetime
from config.VarConfig import screenPicturesDir,errorexcelDir

def getCurrentDate():
    timeTup = time.localtime()
    currentDate = str(timeTup .tm_year) + "-" +str(timeTup .tm_mon) + "-" + str(timeTup .tm_mday)
    return currentDate

def getCurrentTime():
    timeStr = datetime.now()
    nowtime = timeStr.strftime('%H-%M-%S-%f')
    return nowtime

def createCurrentDateDir():
    dirName = os.path.join(screenPicturesDir ,getCurrentDate())
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName

def createErroeExcelDir():
    dirName = os.path.join(errorexcelDir ,getCurrentDate())
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName

if __name__ == '__main__':
    #print(getCurrentDate())
    print(createCurrentDateDir())
   # print(getCurrentTime())
