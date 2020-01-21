import os

ieDriverFilePath="c:\IEDriverServer"
chromeDriverFilePath="c:\chromedriver"

#获取当前文件所在目录的父目录的绝对路径
parentDirPath=os.path.dirname(os.path.dirname(os.path.abspath(__file__ )))
#异常图片目录
screenPicturesDir = parentDirPath + "\\exceptionpictures\\"
errorexcelDir = parentDirPath + "\\errorexcel\\"

#测试数据文件存放路径
dataFilePath = parentDirPath + "\\testData\\内网OA流程测试.xlsx"

#测试目录文件各列数字序号
testCase_testCaseName = 2
# testCase_frameWorkName = 3
testCase_testStepSheetName = 4
# testCase_dataSoureSheetName = 5
testCase_isExecute = 5
testCase_runTime = 6
testCase_testResult = 7

#用例步骤表，各列序号
testStep_testStepDescribe = 2
testStep_keyWords = 3
testStep_locationType = 4
testStep_locatorExpression = 5
testStep_operateValue = 6
testStep_runTime = 7
testStep_testResult = 8
testStep_errorInfo = 9
testStep_errorPic = 10
testStep_special = 11

#数据源表各列序号,需要仔细思考一下这个表的结构
#dataSource_isExecute =
#datasource_ =
#datasource_ =
#dataSource_runTime =
#dataSource_result =