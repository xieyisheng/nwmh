import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from testScripts.OAtest import now2
from config.VarConfig import errorexcelDir,screenPicturesDir
from util.DirAndTime import getCurrentDate
import zipfile

#以下开始执行循环测试
looppath = os.path.dirname(os.path.realpath("__file__" ))
runpath = os.path.join(looppath,'RunTest.py')

#print(runpath)

testnum=0
for i in range(1):
    os.system("python %s" %runpath)
    testnum += 1
print(testnum)

#封装一个压缩文件夹的方法，给邮件发送模块调用
def zip_ya(start_dir):
    start_dir = start_dir  # 要压缩的文件夹路径
    file_news = start_dir + '.zip'  # 压缩后文件夹的名字

    z = zipfile.ZipFile(file_news, 'w', zipfile.ZIP_DEFLATED)
    for dir_path, dir_names, file_names in os.walk(start_dir):
        f_path = dir_path.replace(start_dir, '')  # 这一句很重要，不replace的话，就从根目录开始复制
        f_path = f_path and f_path + os.sep or ''  # 实现当前文件夹以及包含的所有文件的压缩
        for filename in file_names:
            z.write(os.path.join(dir_path, filename), f_path + filename)
    z.close()


#以下部分执行发送附带测试结果附件的邮件

#发送邮件登录、发送人、收件人、主题、邮件正文、附件等参数准备
smtpserver = 'smtp.qq.com'
user = '413266269@qq.com'
password = 'DX0610750'
sender = '413266269@qq.com'
receivers = ['xieys@zjshenyue.cn','wujm@zjshenyue.cn']
subject = ' 内网OA自动化测试报告'+ now2
mailbody = ' 大家好，今日内网OA自动化测试共执行'+ str(testnum)+ '次,测试结果汇总请下载附件，' \
            '附件只有单个excel结果文件表示测试结果全部通过，附件包含多个压缩文件表示有未通过的测试步骤'
msg = MIMEText(mailbody,'plain', 'utf-8')

#开始组装邮件收件人、发件人、正文、附件
message = MIMEMultipart()
message['From'] = sender
message['To'] = ','.join(receivers)
message['Subject'] = subject
message.attach(msg)
exceldirName = os.path.join(errorexcelDir ,getCurrentDate())
picdirName = os.path.join(screenPicturesDir ,getCurrentDate())
if not os.path.exists(exceldirName):
    stepexcel = MIMEText(open('C:\\Users\\Administrator\\PycharmProjects\\zhihuishequ\\testData\\内网OA流程测试.xlsx', 'rb').read(), 'base64', 'utf-8')#正确用例步骤附件
    stepexcel["Content-Type"] = 'application/octet-stream'
    stepexcel["Content-Disposition"] = 'attachment; filename="correctstepandresult.xlsx"'
    message.attach(stepexcel)
else:
    zip_ya(exceldirName)
    zip_ya(picdirName)
    errorstepexcel = MIMEText(open(exceldirName + '.zip', 'rb').read(), 'base64', 'utf-8')  # 错误用例步骤excel附件
    errorstepexcel["Content-Type"] = 'application/octet-stream'
    errorstepexcel["Content-Disposition"] = 'attachment; filename="errorstepandresult' + getCurrentDate() + '.zip"'
    message.attach(errorstepexcel)
    capturepic = MIMEText(open(picdirName + '.zip', 'rb').read(), 'base64', 'utf-8')  # 错误截图附件
    capturepic["Content-Type"] = 'application/octet-stream'
    capturepic["Content-Disposition"] = 'attachment; filename="errorpic' + getCurrentDate() + '.zip"'
    message.attach(capturepic)
#执行发送邮件
try:
    #smtpObj = smtplib.SMTP()
    #smtpObj.connect(smtpserver, 25)
    smtpObj = smtplib.SMTP_SSL(smtpserver,465)
    smtpObj.login(user,'bkievlqenpilbjfj')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print ("邮件发送成功")
except smtplib.SMTPException as e:
    raise e

