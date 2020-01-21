from util.ObjectMap import now2
import time
from selenium import webdriver

driver= webdriver.Ie(executable_path= "c:\IEDriverServer")

driver.get('https://www.baidu.com/')
driver.maximize_window()
driver.find_element_by_xpath('//*[@id="kw"]') .send_keys('自动化测试')
driver.find_element_by_xpath('//*[@id="su"]').click()
driver.find_element_by_xpath('//*[@id="s_tab"]/div/a[4]').click()
