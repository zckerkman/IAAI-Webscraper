# -*- coding: utf-8 -*-
"""
Created on Wed Jan 31 20:10:35 2018

@author: Zach
"""
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("load-extension=C:\\Users\\Zach\\Downloads\\cjpalhdlnbpafiamejdnhcphjbkeiagm\\1.14.18_0")
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.create_options()

def iterationNumber():
    iterations = input('How many cars are you saving? ')
    return iterations

def loadPage( str ):
    driver.set_page_load_timeout(10)
    try: 
        driver.get( str ) 
        return False
    except TimeoutException:
        driver.execute_script("window.stop();")
        return True

def loginIAAI():
    page = ('https://www.iaai.com/Login/LoginPage?ReturnUrl=/MyDashboard/Default')
    pageLoader = True
    while pageLoader & boolean:
        pageLoader = loadPage(str = page)
     
    WebDriverWait(driver, 100).until( lambda driver: driver.find_element_by_name("usrName"))
    
    usernameField = driver.find_element_by_name("usrName")
    usernameField.click()
    usernameField.send_keys("PUT YOUR USERNAME HERE")    
    
    pswdField = driver.find_element_by_name("pwd")
    pswdField.click()
    pswdField.send_keys("PUT YOUR PASSWORD HERE")
    pswdField.submit()
     
    dropDownMenu = driver.find_element_by_xpath("//nav[@id='mainNav']//a[contains(@href,'MyDashboard')]")
    
    WebDriverWait(driver, 100).until( lambda driver: dropDownMenu )
    
    driver.move_to_element(dropDownMenu)
    purchaseHistory = driver.find_element_by_xpath("//nav[@id='mainNav']//a[contains(@href,'PurchaseHistory')]")
    purchaseHistory.click()
    
def main():
    

if __name__ == "__main__":
    iterations = int(iterationNumber())
    loginIAAI()
    for x in range (0, iterations):
        main()
    
