from selenium import webdriver
from config .VarConfig import  ieDriverFilePath
from config.VarConfig import  chromeDriverFilePath
from util.ObjectMap import  getElement
from util.ClipboardUtil  import Clipboard
from util.KeyBoardUtil import keyBoardKeys
from util.DirAndTime import *
from util .WaitUtil import WaitUtil
from selenium.webdriver .chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from unittest import TestCase
import time
import os
import datetime

#定义全局变量driver
driver = None
# 全局等待实例对象waitUtil
waitUtil = None

now = datetime.datetime.now()
now2 = datetime.datetime.strftime(now,'%Y-%m-%d %H:%M:%S')
print(now2)

#打开浏览器
def open_browser(browserName,*arg):
    global driver,waitUtil
    try:
        if browserName .lower() =='ie':
            driver = webdriver.Ie(executable_path= ieDriverFilePath )
        elif browserName.lower() == 'chrome':
            chrome_options = Options()
            chrome_options.add_argument('disable-infobars') #去除“chrome正受到自动化测试。。。”的弹出提示框
            chrome_options.add_argument('disable-popup-blocking') #禁用弹窗拦截
            '''
            可能会有用：禁用下载时的安全提示
            download_dir = "/pathToDownloadDir"
            preferences = {"download.default_directory": download_dir,
                           "directory_upgrade": True,
                           "safebrowsing.enabled": True}
            chrome_options.add_experimental_option("prefs", preferences)
            '''
            driver = webdriver.Chrome(executable_path= chromeDriverFilePath ,chrome_options= chrome_options )
        waitUtil = WaitUtil(driver)
    except Exception as e:
        raise e
#打开网址
def visit_url(url,*arg):
    global driver
    try:
        driver.get(url)
    except Exception as e:
        raise e

def close_browser(*arg):
    try:
        driver.quit()
    except Exception as e:
        raise e

def sleep(sleepSeconds,*arg):
    try:
        time.sleep(int(sleepSeconds))
    except Exception as e:
        raise e

#清除输入框内容
def clear(locationType,locatorExpression,*arg):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).clear()
    except Exception as e:
        raise e

#输入框中输入数据
def input_string(locationType,locatorExpression,inputContent):
    global driver
    try:
        getElement(driver,locationType,locatorExpression).send_keys(inputContent)
    except Exception as e:
        raise e

def click(locationType,locatorExpression,*arg):
    global driver
    try:
        getElement(driver, locationType, locatorExpression).click()
    except Exception as e:
        raise e

#断言页面源码是否存在关键字符串
def assert_string_in_pagesource(assertString,*arg):
    global driver
    try:
        assert assertString in driver.page_source, " %s not found in page source" %assertString
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e :
        raise e

def assert_string_notin_pagesource(assertString):
    global driver
    try:
        PageSource = driver.page_source
        TestCase.assertNotIn(None,assertString,PageSource,'错误：文本出现在页面中')
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e

#断言页面标题是否存在关键字符串
def assert_title(titleStr,*args):
    global driver
    try:
        assert titleStr in driver.title,"%s not found in title" %titleStr
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e

def get_title(*arg):
    global driver
    try:
        return driver.title
    except Exception as e:
        raise e

def getPageSource(*arg):
    try:
        return driver.page_source
    except Exception as e:
        raise e

def switch_to_frame(locationType,frameLocatorExpression,*arg):
    global driver
    try:
        driver.switch_to.frame(getElement(driver,locationType,frameLocatorExpression) )
    except Exception as e:
        print("frame error")
        raise e

def switch_to_default_content(*arg):
    global driver
    try:
        driver.switch_to.default_content()
    except Exception as e:
        raise e

def switch_to_window(winnum):
    global driver
    try:
       driver.switch_to .window(driver.window_handles[int(winnum)])
    except Exception as e:
        print("switch to window faild")
        raise e

def paste_string(pasteString,*arg):
    try:
       Clipboard .setText(pasteString)
       time.sleep(2)
       keyBoardKeys.twoKeys("ctrl","v")
    except Exception as e:
        raise e

def press_enter_key(*arg):
    try:
        keyBoardKeys .oneKey("enter")
    except Exception as e:
        raise e

def maximize_browser():
    global driver
    try:
        driver.maximize_window()
    except Exception as e:
        raise e

def capture_screen(*arg):
    global driver
    currTime = getCurrentTime()
    picNameAndPath = os.path.join(str(createCurrentDateDir()) ,str(currTime) + '.png')
    #print(picNameAndPath.replace('\\',r'\\'))
    print(picNameAndPath)
    #picname= '"'+picNameAndPath+'"'
    try:
        driver.get_screenshot_as_file(picNameAndPath)
        #webdriver.Chrome(executable_path= chromeDriverFilePath ).get_screenshot_as_file(picNameAndPath)
    except Exception as e:
        raise e
    else:
        return picNameAndPath



#显示等待页面元素出现在DOM中
def waitPresenceOfElementLocated(locationType, locatorExpression,*arg):
    global waitUtil
    try:
        waitUtil.presenceofElementlocated(locationType, locatorExpression)
    except Exception as e:
        raise e

def waitVisibilityOfElementLocated(locationType, locatorExpression,*arg):
    global waitUtil
    try:
        waitUtil.visibilityOfElementLocated(locationType, locatorExpression)
    except Exception as e:
        raise e

def waitElementToBeClickable(locationType, locatorExpression):
    global waitUtil
    try:
        waitUtil.elementToBeClickable(locationType, locatorExpression)
    except Exception as e:
        raise e

def waitFrameToBeAvailableAndSwitchToIt(locationType, locatorExpression,*arg):
    global waitUtil
    try:
        waitUtil.frameToBeAvailableAndSwtitchToIt(locationType, locatorExpression)
    except Exception as e:
        raise e

def uploadfile(dir):
    dir=str(dir)
    os.system(dir)

def switch_to_active_element():
    global driver
    try:
        driver.switch_to.active_element.click()
    except Exception as e:
        raise e

def move_by_offset(x,y):
    global driver
    try:
        ActionChains.move_by_offset(driver,x,y)
    except Exception as e:
        raise e

def move_to_element_with_offset(locationType, locatorExpression,x,y):
    global driver
    action=ActionChains(driver)
    try:
        x=int(x)
        y=int(y)
        action.move_to_element_with_offset(getElement(driver,locationType, locatorExpression),x,y)
    except Exception as e:
        raise

def context_click():
    global driver
    action = ActionChains(driver)
    action.context_click()

def mclick():
    global driver
    action = ActionChains(driver)
    action.click()

def perform():
    global driver
    action = ActionChains(driver)
    action.perform()

def doubleclick(locationType, locatorExpression):
    global driver
    try:
        action = ActionChains(driver)
        action.double_click(getElement(driver, locationType, locatorExpression)).perform()
    except Exception as e:
        raise e

def checkwindowsnum(windowsnum):
    global driver
    try:
        windowsnum = int(windowsnum)
        winlist = len(driver.window_handles)
        assert winlist == windowsnum
    except AssertionError as e:
        print('窗口数量检查出错')
        raise e
    finally:
        print(winlist)

def closewindow():
    global driver
    try:
        driver.close()
    except Exception as e:
        raise e

def getidfromxpath(locatorExpression):
    global driver
    try:
        #elementid = driver.find_element_by_xpath(locatorExpression).id
        elementid = getElement(driver, 'xpath', locatorExpression).get_attribute('id')
        return elementid
    except Exception as e:
        raise e






if __name__ == '__main__':


   open_browser('chrome')
   visit_url('http://192.168.0.203:8090/Login/Index')
   maximize_browser()
   input_string('xpath', '//*[@id="txt_account"]', 'admin')
   click('xpath', '//*[@id="txt_password"]')
   sleep(2)
   clear('xpath', '//*[@id="txt_password"]')
   input_string('xpath', '//*[@id="txt_password"]', '123456')
   click('xpath', '//*[@id="login_button"]')
   sleep(5)
   #等待个性化桌面frame
   waitVisibilityOfElementLocated('xpath','//*[@id="nw_left_menubox"]/div[1]/div/ul/li[2]')
   click('xpath', '//*[@id="nw_left_menubox"]/div[1]/div/ul/li[2]')
   sleep('2')
   click('xpath', '//*[@id="nw_left_menubox"]/div[2]/div/div[2]/ul/li[2]/a')
   sleep('1')
   click('xpath', '//*[@id="nw_left_menubox"]/div[2]/div/div[2]/ul/li[2]/ul/li[6]/a')
   sleep('1')
   click('xpath', '//*[@id="nw_left_menubox"]/div[2]/div/div[2]/ul/li[2]/ul/li[6]/ul/li[1]/a')
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="content-main"]/iframe[2]')
   #waitElementToBeClickable('xpath','//a[text()="浙江申跃信息科技自动化测试—发文2019-08-27 10:26:09"]')
   #click('xpath', '//a[text()="浙江申跃信息科技自动化测试—发文2019-08-28 23:13:57"]/../..')
   print(str(getidfromxpath('//a[text()="浙江申跃信息科技自动化测试—发文2019-08-28 23:13:57"]/../..')))
   click('id',str(getidfromxpath('//a[text()="浙江申跃信息科技自动化测试—发文2019-08-28 23:13:57"]/../..')))

   #click('css_selector','[title="浙江申跃信息科技自动化测试—发文2019-08-27 09:13:40"]')
   # \39 b96922d-7764-4b80-bb6a-ee45a1ef4386 > td:nth-child(3)
   # \39 b96922d-7764-4b80-bb6a-ee45a1ef4386 > td:nth-child(3)

   #查找符合标题的代办任务
  # assert_string_notin_pagesource('浙江申跃信息科技自动化测试—收文2019-08-23 10:20:55')

   '''
   text = '收文管理【'+ title + '】'
   print(text)
   waitVisibilityOfElementLocated('xpath','//a[text()="'+text+'"]')
   click('xpath','//a[text()="'+text+'"]')
   sleep(2)
   switch_to_default_content()
   #定位表单页面frame，使用上级元素往下查询第二个frame确保任何人打开的第一个代办表单都能准确定位
   waitFrameToBeAvailableAndSwitchToIt('xpath','//*[@id="content-main"]/iframe[2]')
   click('xpath','//*[@id="mainform"]/div[1]/div/a[2]/span')  #点击办理按钮
   switch_to_default_content()   #切换到其他frame前需要先切换回默认窗口
   switch_to_frame('xpath','/html/body/div[15]/div[2]/iframe')  #切换到选择步骤frame窗口
   click('xpath','//*[@id="step_acecd3e1-0922-4c98-bf47-1361d65a8bd6"]') #选择具体处理步骤：请局领导阅示或处室阅处
   sleep(2)
   click('xpath','//*[@id="mainform"]/table/tbody/tr[2]/td/input[3]') #点击选择步骤后出现的人员输入框后面的 选择 按钮
   sleep(2)
   switch_to_default_content()
   switch_to_frame('xpath','/html/body/div[17]/div[2]/iframe') #切换到人员选择frame窗口
   click('xpath','//*[@id="roadui_tree_3"]/ul/li[1]/span') # 点击展开办公室人员树
   sleep(2)
   click('xpath', '//span[text()= "裘继海"]')  #选择具体人员
   click('xpath', '//span[text()= "王维娜"]')
   click('xpath', '//*[@id="roadui_tree_18"]/ul/li[1]/span')  # 点击展开局领导人员树
   sleep(2)
   click('xpath', '//span[text()= "陈寿旦"]')  # 选择具体人员
   click('xpath', '//span[text()= "夏海明"]')
   click('xpath','//*[@id="selectMemberTable"]/tbody/tr/td[2]/div[4]/button')  #确认所选人员
   switch_to_default_content()
   switch_to_frame('xpath', '/html/body/div[15]/div[2]/iframe')  #确认人员后 回到步骤窗口
   click('xpath','//*[@id="mainform"]/div[4]/input[1]')  #点击确定按钮
   switch_to_default_content()
   click('xpath','//*[@id="layui-layer1"]/div[3]/a[1]') # 在弹出对话框中，确认发送
   sleep(5)
   click('xpath','/html/body/div[9]/div[2]/div[1]/div[2]/a[4]/img') #退出登录

#####################################以下是陈寿旦办理收文，选择退回给许志平####################################
   input_string('xpath','//*[@id="txt_account"]','chensd')
   input_string('xpath','//*[@id="txt_password"]','123456')
   click('xpath','//*[@id="login_button"]')
   #等待个性化桌面frame
   waitFrameToBeAvailableAndSwitchToIt('xpath','//*[@id="iframe578a7740-5681-4939-ba39-95bf1bce5d89"]')
   #查找符合标题的代办任务
   text = '收文管理【'+ title + '】'
   print(text)
   waitVisibilityOfElementLocated('xpath', "//a[text()='" + text + "']")
   click('xpath',"//a[text()='"+text+"']")
   sleep(2)
   switch_to_default_content()
   #定位表单页面frame，使用上级元素往下查询第二个frame确保任何人打开的第一个代办表单都能准确定位
   waitFrameToBeAvailableAndSwitchToIt('xpath','//*[@id="content-main"]/iframe[2]')
   click('xpath','//*[@id="mainform"]/div[1]/div/a[3]/span')  #点击回退按钮
   switch_to_default_content()   #切换到其他frame前需要先切换回默认窗口
   switch_to_frame('xpath','/html/body/div[15]/div[2]/iframe')  #切换到选择步骤frame窗口
   click('xpath','/html/body/div[3]/input[1]') #点击确认按钮，确认退回
   switch_to_default_content()
   click('xpath', '/html/body/div[9]/div[2]/div[1]/div[2]/a[4]/img')  # 退出登录

###########################以下是许志平重新发送给四个人的过程#############################################
   input_string('xpath', '//*[@id="txt_account"]', 'qqb002')
   input_string('xpath', '//*[@id="txt_password"]', '123456')
   click('xpath', '//*[@id="login_button"]')
   # 等待个性化桌面frame
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="iframe578a7740-5681-4939-ba39-95bf1bce5d89"]')
   # 查找符合标题的代办任务
   # click('css_selector', '[class="menuItem"][data-id="80ed0beb-5d3b-4052-b078-867794e445ae"]')
   text = '收文管理【' + title + '】'
   print(text)
   waitVisibilityOfElementLocated('xpath', "//a[text()='" + text + "']")
   click('xpath', "//a[text()='" + text + "']")
   sleep(2)
   switch_to_default_content()
   # 定位表单页面frame，使用上级元素往下查询第二个frame确保任何人打开的第一个代办表单都能准确定位
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="content-main"]/iframe[2]')
   click('xpath', '//*[@id="mainform"]/div[1]/div/a[2]/span')  # 点击办理按钮
   switch_to_default_content()  # 切换到其他frame前需要先切换回默认窗口
   switch_to_frame('xpath', '/html/body/div[15]/div[2]/iframe')  # 切换到选择步骤frame窗口
   click('xpath', '//*[@id="step_acecd3e1-0922-4c98-bf47-1361d65a8bd6"]')  # 选择具体处理步骤：请局领导阅示或处室阅处
   sleep(2)
   click('xpath', '//*[@id="mainform"]/table/tbody/tr[2]/td/input[3]')  # 点击选择步骤后出现的人员输入框后面的 选择 按钮
   sleep(2)
   switch_to_default_content()
   switch_to_frame('xpath', '/html/body/div[17]/div[2]/iframe')  # 切换到人员选择frame窗口
   click('xpath', '//*[@id="roadui_tree_3"]/ul/li[1]/span')  # 点击展开办公室人员树
   sleep(2)
   click('xpath', '//span[text()= "裘继海"]')  # 选择具体人员
   click('xpath', '//span[text()= "王维娜"]')
   click('xpath', '//*[@id="roadui_tree_18"]/ul/li[1]/span')  # 点击展开局领导人员树
   sleep(2)
   click('xpath', '//span[text()= "陈寿旦"]')  # 选择具体人员
   click('xpath', '//span[text()= "夏海明"]')
   click('xpath', '//*[@id="selectMemberTable"]/tbody/tr/td[2]/div[4]/button')  # 确认所选人员
   switch_to_default_content()
   switch_to_frame('xpath', '/html/body/div[15]/div[2]/iframe')  # 确认人员后 回到步骤窗口
   click('xpath', '//*[@id="mainform"]/div[4]/input[1]')  # 点击确定按钮
   switch_to_default_content()
   click('xpath', '//*[@id="layui-layer1"]/div[3]/a[1]')  # 在弹出对话框中，确认发送
   sleep(5)
   click('xpath', '/html/body/div[9]/div[2]/div[1]/div[2]/a[4]/img')  # 退出登录

########################################以下是陈寿旦办理选择分发给全体人员################################
   input_string('xpath', '//*[@id="txt_account"]', 'chensd')
   input_string('xpath', '//*[@id="txt_password"]', '123456')
   click('xpath', '//*[@id="login_button"]')
   # 等待个性化桌面frame
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="iframe578a7740-5681-4939-ba39-95bf1bce5d89"]')
   # 查找符合标题的代办任务
   text = '收文管理【' + title + '】'
   print(text)
   waitVisibilityOfElementLocated('xpath', "//a[text()='" + text + "']")
   click('xpath', "//a[text()='" + text + "']")
   sleep(2)
   switch_to_default_content()
   # 定位表单页面frame，使用上级元素往下查询第二个frame确保任何人打开的第一个代办表单都能准确定位
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="content-main"]/iframe[2]')
   click('xpath', '//*[@id="mainform"]/div[1]/div/a[2]/span')  # 点击办理按钮
   switch_to_default_content()  # 切换到其他frame前需要先切换回默认窗口
   switch_to_frame('xpath', '/html/body/div[15]/div[2]/iframe')  # 切换到选择步骤frame窗口
   click('xpath', '//*[@id="step_d199324e-55e2-402c-964f-e17feb3ef2f6"]')  # 选择具体处理步骤：分发
   sleep(2)
   click('xpath', '//*[@id="mainform"]/table/tbody/tr[14]/td/input[3]')  # 点击选择步骤后出现的人员输入框后面的 选择 按钮
   sleep(2)
   switch_to_default_content()
   switch_to_frame('xpath', '/html/body/div[17]/div[2]/iframe')  # 切换到人员选择frame窗口
   click('xpath', '//*[@id="selectMemberTable"]/tbody/tr/td[2]/div[2]/button')  # 点击全选按钮。
   waitVisibilityOfElementLocated('xpath', '//div[text()= "局领导"]')
   click('xpath', '//*[@id="selectMemberTable"]/tbody/tr/td[2]/div[4]/button')  # 确认所选人员
   switch_to_default_content()
   switch_to_frame('xpath', '/html/body/div[15]/div[2]/iframe')  # 确认人员后 回到步骤窗口
   click('xpath', '//*[@id="mainform"]/div[4]/input[1]')  # 点击确定按钮
   switch_to_default_content()
   click('xpath', '//*[@id="layui-layer1"]/div[3]/a[1]')  # 在弹出对话框中，确认发送
   sleep(5)
   click('xpath', '/html/body/div[9]/div[2]/div[1]/div[2]/a[4]/img')  # 退出登录

   ####################################以下是商海峰登录，查看附件在线预览。并点击已阅#######################################
   input_string('xpath','//*[@id="txt_account"]','xx005')
   input_string('xpath','//*[@id="txt_password"]','123456')
   click('xpath','//*[@id="login_button"]')
   #等待个性化桌面frame
   waitFrameToBeAvailableAndSwitchToIt('xpath','//*[@id="iframe578a7740-5681-4939-ba39-95bf1bce5d89"]')
   #查找符合标题的代办任务
   text = '收文管理【'+ title + '】'
   print(text)
   waitVisibilityOfElementLocated('xpath', "//a[text()='" + text + "']")
   click('xpath',"//a[text()='"+text+"']")
   sleep(2)
   switch_to_default_content()
   #定位表单页面frame，使用上级元素往下查询第二个frame确保任何人打开的第一个代办表单都能准确定位
   waitFrameToBeAvailableAndSwitchToIt('xpath','//*[@id="content-main"]/iframe[2]')
   click('xpath', '//a[text()= "建委投资管理系统任务书原型修改建议（韩松明）2019325.doc"]')
   switch_to_window(-1)
   sleep(2)
   switch_to_window(0)
   waitFrameToBeAvailableAndSwitchToIt('xpath', '//*[@id="content-main"]/iframe[2]')
   click('xpath','//*[@id="mainform"]/div[1]/div/a[5]/span')  #点击已阅按钮
'''




