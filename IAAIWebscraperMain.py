# -*- coding: utf-8 -*-
"""
Created on Wed Jan 31 20:10:35 2018

@author: Zach
"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options

from openpyxl import load_workbook

##Starts Selenium maximized and renames it to the variable 'driver'
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.create_options()

##Defines how many iterations do run main
def iterationNumber():
    iterations = input('How many cars are you saving? ')
    return int(iterations)

##Module to restart the page if it times out while loading
def loadPage( str ):
    pageLoader = False
    while not pageLoader:
        driver.set_page_load_timeout(10)
        try: 
            driver.get( str ) 
            pageLoader = True
        except TimeoutException:
            driver.execute_script("window.stop();")


##Calls the loadPage method on the IAAI login website and logs in, USER MUST INPUT CREDENTIALS IN THE CODE!
def loginIAAI():
    page = ('https://www.iaai.com/Login/LoginPage?ReturnUrl=/MyDashboard/Default')
    
    ##Runs the pageloader with the AIIA website 
    loadPage(str = page)
     
    ##Waits for the usrname field to become visible and then inputs usrname&pswrd and submits
    WebDriverWait(driver, 100).until( lambda driver: driver.find_element_by_name("usrName"))
    
    usernameField = driver.find_element_by_name("usrName")
    usernameField.click()
    usernameField.send_keys("PUT YOUR USERNAME HERE")    
    
    pswdField = driver.find_element_by_name("pwd")
    pswdField.click()
    pswdField.send_keys("PUT YOUR PASSWORD HERE")
    pswdField.submit()
    
    ##Waits for the required drop down menu to load, hovers over it, and clicks the 'purchase history' icon 
    dropDownMenu = driver.find_element_by_xpath("//nav[@id='mainNav']//a[contains(@href,'MyDashboard')]")
    WebDriverWait(driver, 100).until( lambda driver: dropDownMenu )
    
    driver.move_to_element(dropDownMenu)
    purchaseHistory = driver.find_element_by_xpath("//nav[@id='mainNav']//a[contains(@href,'PurchaseHistory')]")
    purchaseHistory.click()
    
def carInformation( int ):
    informationArray = []
    return informationArray
    

if __name__ == "__main__":
    iterations = iterationNumber()
    loginIAAI()
    carDictionaryKey = []
    carDictionary = {}
    for x in range (0, iterations):
        informationArray = carInformation( int = x )
        value = informationArray.pop[0]
        carDictionaryKey.append(value)
        carDictionary[value] = informationArray

    for keys in carDictionaryKey:
        wb = load_workbook(filename = 'Cost_of_Vehicle_Format_Master.xls')
        ws = wb.active
        valueArray = carDictionaryKey[keys]
        ws["A1"] = 