from selenium.webdriver .support .ui import WebDriverWait
import datetime

now = datetime.datetime.now()
now2 = datetime.datetime.strftime(now,'%Y-%m-%d %H:%M:%S')
def getElement(driver,locationType,locatorExpression):
    try:
        element = WebDriverWait (driver ,30).until(lambda x: x.find_element(by= locationType ,value= locatorExpression ))
        return element
    except Exception as e:
        raise e

def getElements(driver,locationType,locatorExpression):
    try:
        elements = WebDriverWait(driver, 20).until(lambda x: x.find_elements(by=locationType, value=locatorExpression ))
        return elements
    except Exception as e:
        raise e

if __name__  == '__main__':
    from selenium import webdriver
    driver = webdriver .Chrome (executable_path= "c:\chromedriver")
    driver.get("http://www.baidu.com")
    seachbox=getElement(driver,"id","kw")
    print(seachbox.tag_name)
    alist=getElements(driver,"tag name","a")
    print(len(alist))
    driver.quit()

