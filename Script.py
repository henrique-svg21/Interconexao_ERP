import os
import time
from datetime import datetime, timedelta
import pyautogui as py
from dotenv import load_dotenv
import sys
import openpyxl

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv() #load .env file, where email and password are stored
py.size() #define screen size
(1366,768)
py.FAILSAFE = True #just make sure it is enabled

chrome_options = Options()

# Open in Incognito & Dark Mode
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--force-dark-mode")
chrome_options.add_argument("--enable-features=WebUIDarkMode")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)

# Hide basic automation banners
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

servicce = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=servicce, options=chrome_options)

email = os.getenv("email")
password = os.getenv("password")

try:    
    # logging into google   
    browser.get("https://accounts.google.com/")

    email_input = WebDriverWait(browser, 30).until(
        EC.visibility_of_element_located((By.ID, "identifierId"))
    )
    email_input.send_keys(email + Keys.ENTER)

    password_input = WebDriverWait(browser, 30).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))
    )
    time.sleep(1)
    password_input.send_keys(password + Keys.ENTER)
    time.sleep(5)
    
    browser.get("https://forms.gle/SGy2pnLA9LdQfXeSA") #opens the form's link (test)

    # Usage
    def read_excel_data(file_path, sheet_name, cell_coordinate):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        cell_value = sheet[cell_coordinate].value
        return cell_value
    
    file_path = r"C:\Users\henrique_schorck\Documents\Base_dados.xlsx"

    # opening windows explorer and finding + opening the file
    py.hotkey('win', 'e')
    time.sleep(3)
    py.hotkey('ctrl', 'f')
    time.sleep(3)
    py.write(file_path)
    time.sleep(5)
    py.dragTo(381,156)
    time.sleep(3)
    py.leftClick()
    py.leftClick()
    time.sleep(5)

    #getting back to forms
    py.hotkey('alt','tab')
    time.sleep(3)
    py.hotkey('alt','f4')
    time.sleep(3)
    py.hotkey('alt','tab')
    time.sleep(3)

    #confirming email
    py.dragTo(390,640)
    time.sleep(0.5)
    py.leftClick()
    time.sleep(1)

    #DATA INSERTION IN THE FILE
    for i in range (1,10): 
        if i > 1:
            py.press('tab')
            time.sleep(1) 
        one_value = read_excel_data(file_path, 'Página1', 'B2')
        py.press('tab')
        time.sleep(1)
        py.press('enter')
        time.sleep(1)
        py.press('down', presses=i, interval=0.3)
        time.sleep(2)
        py.press('enter')
        time.sleep(1)
        py.press('tab')
        time.sleep(1)
        py.typewrite(str(one_value))    
        time.sleep(1)
        py.press('tab', presses=2, interval=0.5)
        time.sleep(1)

        if one_value >= 40: #OBJECTIVE REACHED
            py.dragTo(390,372)
            time.sleep(1)
            py.leftClick()
            time.sleep(1)
            py.press('tab', presses=2, interval=0.5)
            time.sleep(1)
            py.press('enter')
            time.sleep(1)

        else: #OBJECTIVE NOT REACHED
            py.dragTo(390, 587)
            time.sleep(1)
            py.leftClick()
            time.sleep(1)
            py.press('tab', presses=3, interval=0.5)
            time.sleep(1)
            py.press('enter')
            time.sleep(1)

except py.FailSafeException:
    print('Fail-safe activated. Exiting program.')
    sys.exit() 

