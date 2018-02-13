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

import pyautogui

from openpyxl import load_workbook
from openpyxl.styles import Font

import os
import time 

##defines current folder, used to save car images before moving them into new folders
dirPath = os.path.dirname(os.path.realpath(__file__))

##Starts Selenium maximized and renames it to the variable 'driver'
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("download.default_directory=%s" % dirPath)
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
    
def createFolder(str):
    ##creates a new folder for the vehicle in the current script folder
    os.makedirs(os.path.join(dirPath,str))
    
def carInformation( int ):
    ##finds the price of the car
    carPrice = driver.find_element_by_xpath('//a[@class=])    
    
    ##clickes the car we want info on with respect to the iteration number
    driver.find_element_by_xpath('//a[@class id="stockitemID_%s"]/a[conatins(@href onclick,"return")]' % int).click
   

    ##locates and adds the vin number, car model, then scrolls to find the exterior color(seperated from the interior), and the mileage
    vinNumber = driver.find_element_by_xpath('//div[contains(@class,"vehicle-info-wrapper")]/span[contains(@id,"VIN_vehicleStats")]').text
    
    carModel = driver.find_element_by_xpath('//div[@class="flex-item"]/h1[@class="pd-title-ymm"]').text
    
    carColors = driver.find_element_by_xpath('//div[@class="tabs tab-vehicle waypoint-trigger"]/div[@class="row flex"][14]/div[contains(@class,"flex-self-end")]')
    carColors.location_once_scrolled_into_view
    ##splits the exterior and interior colors
    carColorsArray = carColors.text.split("/")

    carMileage = driver.find_element_by_xpath('//div[@class="pd-condition-wrapper]/div[@class="row flex"][4]')
    ##removes the text after the mileage of the car
    carMileageArray = carMileage.text.split("mi")
    
    ##concatenates all the parts of the typeOfVehicle
    typeOfVehicle = (carModel + " " + carMileageArray.pop[0] + " " + carColorsArray.pop[0] + " " + vinNumber)
    ##removes the spaces
    vehicleSaveName = typeOfVehicle.replace(" ","")
    ##creates a folder for the vehicle type
    createFolder(str=vehicleSaveName)
    
    driver.find_element_by_xpath('//a[@id="DownloadImages"]').click
    imageFolderName = vehicleSaveName + "Images.ZIP"
    pyautogui.typewrite(imageFolderName)
    pyautogui.press('enter')
    
    ##Defines the initial folder path and the final folder path for the images
    folderInitialPath = os.path.join(dirPath,imageFolderName)
    folderFinalPath = os.path.join(dirPath,vehicleSaveName,imageFolderName)
   
    ##checks to see if the image folder has been downloaded, if it has it adds it to the folder for the vehicle
    while not os.path.exists(folderInitialPath):
        time.sleep(1)
    os.rename(folderInitialPath,folderFinalPath)
    
    ##finds the damage on the vehicle
    typeOfDamage = driver.find_element_by_xpath('//div[@class="pd-condition-wrapper]/div[@flass="row flex"]/div[contains(@class,"flex-self-end")]').text                                  
                                               
    carInformationArray = [typeOfDamage,price,typeOfVehicle]
    
    return carInformationArray  
   

def addToExcelSpreadsheet ( str, list ):    
    
    ##loads the template workbook
    templateExcelSpreadsheet = "Cost_of_Vehicle_Format_Master.xlsx"
    templateWorkbook = load_workbook(templateExcelSpreadsheet)
    
    #sets the font to change the styles in A1 and A2 to bold and italics
    font = Font(name = 'Arial', size = 9, bold = True, italic = True)
    
    ##adds the price of the car
    templateWorkbook['B5'] = list.pop
    
    ##adds the carType and the text already in the template workbook
    initialA1Text = templateWorkbook['A1']
    templateWorkbook['A1'] = initialA1Text + str
      
    ##adds the type of damage
    initialA2Text = templateWorkbook['A2']
    templateWorkbook['A2'] = initialA2Text + list.pop
    
    ##bolds and italicizes the carType and damage
    templateWorkbook['A1'] = font
    templateWorkbook['A2'] = font
    
    ##saves the excel file as the Type of Vehicle in the folder with the images 
    vehicleSaveName = str.remove(" ","")
    fileName = vehicleSaveName + "Spreadsheet.xlsx"
    templateWorkbook.save(os.path.join(dirPath,vehicleSaveName,fileName), as_template=False))
    
def main():
    ##returns the interations function
    iterations = iterationNumber()
    ##logs in to the website
    loginIAAI()
    ##defines the array for typeOfVehicle keys
    vehicles = []
    ##defines the dictionary to hold the price and damage type
    carDictionary = {}
    for x in range (0, iterations):
        ##navigates to recent purchases
        returnToPurchases()
        carInformationArray = carInformation( int = x )
        ##adds the typeOfVehicle to vehciles and sets it as a the key with price and damage as valuse
        typeOfVehicle = carInformationArray.pop
        vehicles.append(typeOfVehicle)
        carDictionary[typeOfVehicle] = carInformationArray

    for cars in vehicles:
        ##adds each dictionary key and value to a spreadsheet
        addToExcelSpreadsheet(str = cars, list = carDictionary[cars])
        
if __name__ == "__main__":
   main()
       

