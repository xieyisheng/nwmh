import click as click
from selenium .webdriver.common.by import By
from selenium .webdriver.support.ui import WebDriverWait
from selenium .webdriver.support import expected_conditions as EC

import time

class WaitUtil(object):
    def __init__(self,driver):
        self.locationTypeDict = {
            "xpath":By.XPATH,
            "id":By.ID,
            "name":By.NAME,
            "css_selector":By.CSS_SELECTOR ,
            "class_name":By.CLASS_NAME ,
            "tag_name":By.TAG_NAME ,
            "link_text":By.LINK_TEXT ,
            "partial_link_text":By.PARTIAL_LINK_TEXT
        }
        self.driver = driver
        self.wait =WebDriverWait (self.driver,20)



    def presenceofElementlocated(self,locatorMethod,locatorExpression,*args ):
        try:
            if locatorMethod.lower() in self.locationTypeDict :
                element = self.wait.until(EC.presence_of_element_located ((self.locationTypeDict[locatorMethod .lower()],locatorExpression )))
                return element
            else:
                raise TypeError ("未找到定位方式，请确认定位方法是否正确")
        except Exception as e:
            raise e
    def frameToBeAvailableAndSwtitchToIt(self,locationType,locatorExpression,*args ):
        try:
            element = self.wait.until (EC.frame_to_be_available_and_switch_to_it ((self.locationTypeDict[locationType .lower()],locatorExpression )))
            return element
        except Exception  as e:
            raise e

    def visibilityOfElementLocated(self,locationType,locatorExpression,*args):
        try:
            element = self.wait.until(EC.visibility_of_element_located ((self.locationTypeDict[locationType .lower()],locatorExpression)))
            return element
        except Exception as e:
            raise e

    def elementToBeClickable(self,locationType,locatorExpression):
        try:
            element = self.wait.until(EC.element_to_be_clickable((self.locationTypeDict[locationType.lower()], locatorExpression)))
            return element
        except Exception as e:
            raise e


if __name__ =='__main__':
    from selenium import webdriver
    driver=webdriver.Chrome (executable_path= "c:\\chromedriver")
    driver.get("http://mail.126.com")
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="lbNormal"]').click()
    waituntil=WaitUtil(driver)
    waituntil.frameToBeAvailableAndSwtitchToIt("xpath",'// *[ @ id = "loginDiv"]/iframe[1]')
    waituntil.visibilityOfElementLocated("xpath","//input[@name='email']")
    driver.quit()


