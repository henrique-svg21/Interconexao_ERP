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

# VARIÁVEL DE CONTROLE DE TEMPO (Altere aqui para mudar todos de uma vez)
DELAY = 2

def StringConcat():
    return "B" + str(i + 1)

load_dotenv() 
py.size() 
(1366,768)
py.FailSafeException = True 

chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--force-dark-mode")
chrome_options.add_argument("--enable-features=WebUIDarkMode")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

servicce = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=servicce, options=chrome_options)

email = os.getenv("email")
password = os.getenv("password")

try:    
    browser.get("https://accounts.google.com/")

    email_input = WebDriverWait(browser, 30).until(
        EC.visibility_of_element_located((By.ID, "identifierId"))
    )
    email_input.send_keys(email + Keys.ENTER)

    password_input = WebDriverWait(browser, 30).until(
        EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))
    )
    time.sleep(DELAY) 
    password_input.send_keys(password + Keys.ENTER)
    time.sleep(5) # Mantido fixo para carregamento da página
    
    browser.get("https://forms.gle/SGy2pnLA9LdQfXeSA") 

    def read_excel_data(file_path, sheet_name, cell_coordinate):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        cell_value = sheet[cell_coordinate].value
        return cell_value
    
    file_path = r"C:\Users\henrique_schorck\Documents\Base_dados.xlsx"

    py.hotkey('win', 'e')
    time.sleep(2)
    py.hotkey('ctrl', 'f')
    time.sleep(1.5)
    py.write(file_path)
    time.sleep(3)
    py.dragTo(381,156)
    time.sleep(2)
    py.leftClick()
    py.leftClick()
    time.sleep(3.5)

    py.hotkey('alt','tab')
    time.sleep(3)
    py.hotkey('alt','f4')
    time.sleep(3)
    py.hotkey('alt','tab')
    time.sleep(3)

    py.dragTo(390,640)
    time.sleep(0.5)
    py.leftClick()
    time.sleep(DELAY)

    # LAÇO FOR COM TEMPOS DINÂMICOS
    for i in range (1,10): 
        if i > 1:
            py.press('tab')
            time.sleep(DELAY)
            
        one_value = read_excel_data(file_path, 'Página1', StringConcat())
        py.press('tab')
        time.sleep(DELAY)
        
        py.press('enter')
        time.sleep(DELAY)
        
        py.press('down', presses=i, interval=0.3)
        time.sleep(2) 
        
        py.press('enter')
        time.sleep(DELAY)
        
        py.press('tab')
        time.sleep(DELAY)
        
        py.typewrite(str(one_value))    
        time.sleep(DELAY)
        
        py.press('tab', presses=2, interval=0.5)
        time.sleep(DELAY)

        if one_value >= 40: 
            if i > 1:
                py.press('tab')
            if i < 2:
                py.dragTo(390,372)
            else:
                py.dragTo(390, 516)
            time.sleep(DELAY)
            py.leftClick()
            time.sleep(DELAY)
            
            if i < 2:
                py.press('tab', presses=2, interval=0.5)
            else:
                py.press('tab', presses=3, interval=0.5)
            time.sleep(DELAY)
            py.press('enter')
            time.sleep(DELAY)

        else: 
            py.dragTo(390, 587)
            time.sleep(DELAY)
            py.leftClick()
            time.sleep(DELAY)
            py.press('tab', presses=2, interval=0.5)
            time.sleep(DELAY)
            py.press('enter')
            time.sleep(DELAY)

except py.FailSafeException:
    print('Fail-safe activated. Exiting program.')
    sys.exit()