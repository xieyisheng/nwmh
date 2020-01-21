from action.PageAction import *

from util.ParseExcel import PaseExcel
from config.VarConfig import *
import time
import traceback
from util.Log import *
import datetime
from util.DirAndTime import createErroeExcelDir
import shutil

import importlib,sys
importlib.reload(sys)

now = datetime.datetime.now()
now2 = datetime.datetime.strftime(now,'%Y-%m-%d %H:%M:%S')
now3 = datetime.datetime.strftime(now,'%Y-%m-%d-%H-%M-%S')

excelObj = PaseExcel()
excelObj.loadWorkBook(dataFilePath )

def writeTestResult(sheetObject,rowNo,colsNo,testResult,errorInfo=None,picPath=None):
    colorDict = {"pass":"green","faild":"red"}
    colsDict = {"testCase":[testCase_runTime,testCase_testResult],"caseStep":[testStep_runTime,testStep_testResult]}
    try:
        excelObj.writeCellCurrentTime(sheetObject,None,rowNo,colsNo = colsDict[colsNo][0])
        excelObj.writeCell(sheetObject,content=testResult,rowNo=rowNo,colsNo = colsDict[colsNo][1],style=colorDict[testResult])

        if errorInfo and picPath:
            excelObj.writeCell(sheetObject,content=errorInfo,rowNo=rowNo,colsNo=testStep_errorInfo)
            excelObj.writeCell(sheetObject,content=picPath,rowNo=rowNo,colsNo=testStep_errorPic)
        else:
            excelObj.writeCell(sheetObject, content="", rowNo=rowNo, colsNo=testStep_errorInfo)
            excelObj.writeCell(sheetObject, content="", rowNo=rowNo, colsNo=testStep_errorPic)
    except Exception as e:
        logging.debug("写入excel出错，%s" %traceback.format_exc())


def OAtest():
    try:
        caseSheet = excelObj.getSheetByName("测试用例")
        isExecuteColumn = excelObj.getColumn(caseSheet,testCase_isExecute)
        successfulCase=0
        requiredCase = 0
        for idx,i in enumerate(isExecuteColumn[1:]):
            if i.value == 'y':
                requiredCase += 1
                caseRow = excelObj.getRow(caseSheet,idx + 2)
                caseStepSheetName = caseRow[testCase_testStepSheetName - 1].value
                stepSheet = excelObj.getSheetByName(caseStepSheetName)
                stepNum = excelObj.getRowsNumber(stepSheet)
                successfulSteps = 0
                logging.info("开始执行用例:%s" %caseRow[testCase_testCaseName - 1].value)
                for step in range(2,stepNum + 1):
                    stepRow = excelObj.getRow(stepSheet,step)
                    keyWord = stepRow[testStep_keyWords -1].value
                    locationType = stepRow[testStep_locationType - 1].value
                    locatorExpression = stepRow[testStep_locatorExpression - 1].value
                    operateValue = stepRow[testStep_operateValue - 1].value
                    if isinstance(operateValue,int):
                        operateValue = str(operateValue)
                    expressionStr = ""
                    if stepRow[testStep_special - 1].value == "n":   #待验证自添加的代码
                        if keyWord and operateValue and locationType is None and locatorExpression is None:
                            expressionStr = keyWord.strip() + "('" +operateValue+"')"
                        elif keyWord and operateValue is None and locationType is None and locatorExpression is None:
                            expressionStr = keyWord.strip() + "()"
                        elif keyWord and locationType and operateValue and locatorExpression is None:
                            expressionStr = keyWord.strip() +"('" +locationType.strip()+"','"+operateValue+"')"
                        elif keyWord and locationType and locatorExpression and operateValue:
                            expressionStr = keyWord.strip() +"('" +locationType.strip()+"','"+locatorExpression.replace("'",'"').strip() + "','" +operateValue+"')"
                        elif keyWord and locationType and locatorExpression and operateValue is None:
                            expressionStr = keyWord.strip() +"('" +locationType.strip()+"','"+locatorExpression.replace("'",'"').strip() +"')"
                    elif stepRow[testStep_special - 1].value == "u":   # testStep_special 特殊处理标志为u时，执行以下拼接
                        expressionStr = keyWord.strip() + '("' + operateValue + '")'
                    elif stepRow[testStep_special - 1].value == "i":     #表示需要输入带时间戳的标题
                        operateValue = operateValue + now2
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locatorExpression.replace("'", '"').strip() \
                                        + "','" + operateValue + "')"
                    elif stepRow[testStep_special - 1].value == "cc":    #表示需要点击带时间戳的收文
                        title = operateValue+ now2
                        text = '收文管理【' + title + '】'
                        locatorExpression = '//a[text()="' + text + '"]'
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locatorExpression + "')"
                    elif stepRow[testStep_special - 1].value == "cn":    #表示需要点击带时间戳的内部要事
                        title = operateValue+ now2
                        text = '内部要事【' + title + '】'
                        locatorExpression = '//a[text()="' + text + '"]'
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locatorExpression + "')"
                    elif stepRow[testStep_special - 1].value == "cf":    #表示需要点击带时间戳的发文
                        title = operateValue+ now2
                        text = '发文管理【' + title + '】'
                        locatorExpression = '//a[text()="' + text + '"]'
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locatorExpression + "')"
                    elif stepRow[testStep_special - 1].value == "ck":    #表示需要断言check指定字符串出现或未出现在页面中
                        operateValue = operateValue+ now2
                        expressionStr = keyWord.strip() + "('" + operateValue + "')"
                    elif stepRow[testStep_special - 1].value == "cg":    #表示需要点击带时间戳的公文
                        text = operateValue + now2
                        locatorExpression = '//a[text()="' + text + '"]'
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "','" + locatorExpression.replace("'",'"').strip() + "')"
                    elif stepRow[testStep_special - 1].value == "cgt":
                        text = operateValue + now2
                        locatorExpression = 'str(getidfromxpath(\'//a[text()="' + text + '"]/../..\'))'
                        expressionStr = keyWord.strip() + "('" + locationType.strip() + "'," + locatorExpression + ")"


                    try:
                        print(expressionStr)
                        eval(expressionStr)
                      #  excelObj.writeCellCurrentTime(stepSheet,rowNo=step,colsNo=testStep_runTime)
                    except Exception as e:
                        capturePic=capture_screen()
                        errorInfo=traceback.format_exc()

                        writeTestResult(stepSheet,step,"caseStep","faild",errorInfo,capturePic)
                        logging.info("步骤 %s 执行失败，错误信息：%s" %(stepRow[testStep_testStepDescribe -1].value,errorInfo))
                    else:
                        writeTestResult(stepSheet,step,"caseStep","pass")
                        successfulSteps += 1
                        logging.info("步骤 %s 执行通过" %stepRow[testStep_testStepDescribe -1].value)
                if  successfulSteps == stepNum - 1:
                    writeTestResult(caseSheet,idx+2,"testCase","pass")
                    successfulCase += 1
                else:
                    writeTestResult(caseSheet,idx+2,"testCase","faild")
        logging.info("共%d条用例，%d条需要被执行，本次执行通过%d条。" %(len(isExecuteColumn)-1,requiredCase,successfulCase))
        if requiredCase != successfulCase:
            #newpath = 'C:\\Users\\Administrator\\PycharmProjects\\zhihuishequ\\errorexcel\\内网OA流程测试'+ now3 + '.xlsx'
            #shutil.copy("C:\\Users\\Administrator\\PycharmProjects\\zhihuishequ\\testData\\内网OA流程测试.xlsx",newpath)
            errorstepexcelpath = os.path.join(str(createErroeExcelDir()), '内网OA流程测试'+ now3 + '.xlsx')
            shutil.copy("C:\\Users\\Administrator\\PycharmProjects\\zhihuishequ\\testData\\内网OA流程测试.xlsx", errorstepexcelpath)

    except Exception as e:
        print(traceback.print_exc())







