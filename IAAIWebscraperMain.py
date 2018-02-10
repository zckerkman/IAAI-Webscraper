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
    loginPage = ('https://www.iaai.com/Login/LoginPage?ReturnUrl=/MyDashboard/Default')
    
    ##Runs the pageloader with the AIIA website 
    loadPage(str = loginPage)
     
    ##Waits for the usrname field to become visible and then inputs usrname&pswrd and submits
    WebDriverWait(driver, 100).until( lambda driver: driver.find_element_by_name("usrName"))
    
    usernameField = driver.find_element_by_name("usrName")
    usernameField.click()
    usernameField.send_keys("PUT YOUR USERNAME HERE")    
    
    pswdField = driver.find_element_by_name("pwd")
    pswdField.click()
    pswdField.send_keys("PUT YOUR PASSWORD HERE")
    pswdField.submit()

def returnToPurchases ():
    ##defines the url for the purchase history page with access to cars recently purchase
    purchaseHistoryPage = ('https://www.iaai.com/PurchaseHistory/Default')
    loadPage(str = purchaseHistoryPage)
    
def carInformation( int ):
    ##clickes the car we want info on with respect to the iteration number
    driver.find_element_by_xpath('//a[@class id="stockitemID_%s"]/a[conatins(@href onclick,"return")]' % int).click
   
    ##creates an array to store the data in 
    informationArray = []
    
    ##locates and adds the vin number, car model, then scrolls to find the exterior color(seperated from the interior), and the mileage
    vinNumber = driver.find_element_by_xpath('//div[contains(@class,"vehicle-info-wrapper")]/span[contains(@id,"VIN_vehicleStats")]').text
    informationArray.append(vinNumber)
    
    carModel = driver.find_element_by_xpath('//div[@class="flex-item"]/h1[@class="pd-title-ymm"]').text
    informationArray.append(carModel)
    
    carColors = driver.find_element_by_xpath('//div[@class="tabs tab-vehicle waypoint-trigger"]/div[@class="row flex"][14]/div[contains(@class,"flex-self-end")]')
    carColors.location_once_scrolled_into_view
    ##splits the exterior and interior colors
    carColorsArray = carColors.text.split("/")
    informationArray.append(carColorsArray.pop[0])
    
    carMileage = driver.find_element_by_xpath('//div[@class="pd-condition-wrapper]/div[@class="row flex"][4]')
    ##removes the text after the mileage of the car
    carMileageArray = carMileage.text.split("mi")
    informationArray.append(carMileageArray[0])
    
    return informationArray
    
def addToExcelSpreadsheet ( str, list ):
    ##loads the template workbook
    templateExcelSpreadsheet = "Cost_of_Vehicle_Format_Master.xlsx"
    templateWorkbook = load_workbook(templateExcelSpreadsheet)
    
    ##adds the mileage, color, car model, and vin to their respective cells
    templateWorkbook[''] = list.pop
    templateWorkbook[''] = list.pop
    templateWorkbook['B1'] = list.pop
    templateWorkbook[''] = str
                    
    ##saves the excel file as the vinNumber
    fileName = str + ".xlsx"
    templateWorkbook.save(fileName, as_template=False)
    
def main():
    iterations = iterationNumber()
    loginIAAI()
    carDictionaryKey = []
    carDictionary = {}
    for x in range (0, iterations):
        returnToPurchases()
        informationArray = carInformation( int = x )
        vinNumber = informationArray.pop[0]
        carDictionaryKey.append(vinNumber)
        carDictionary[vinNumber] = informationArray

    for keys in carDictionaryKey:
        addToExcelSpreadsheet(str = keys, list = carDictionary[keys])
        
if __name__ == "__main__":
   main()
       

