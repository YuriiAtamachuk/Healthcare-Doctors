from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select
from time import sleep
import openpyxl


def wait_url(driver: webdriver.Chrome, url: str):
    while True:
        cur_url = driver.current_url
        if cur_url == url:
            break
        sleep(0.1)

def find_element(driver: webdriver.Chrome, whichBy, unique: str) -> WebElement:
    while True:
        try:
            element = driver.find_element(whichBy, unique)
            break
        except:
            pass
        sleep(1)
    return element

def find_elements(driver : webdriver.Chrome, whichBy, unique: str) -> list[WebElement]:
    while True:
        try:
            elements =driver.find_elements(whichBy, unique)
            break
        except:
            pass
        sleep(1)
    return elements


driver = webdriver.Chrome()
driver.maximize_window()

url = "https://www.ahpra.gov.au/"
driver.get(url)
sleep(3)

find_element(driver, By.CSS_SELECTOR, "#head > div.desktop-menu > div.desktop-menu__primary-mega-menu-container > div.desktop-menu__primary-nav-container > div > div > button").click()
print("Look UP!")
sleep(2)

find_element(driver, By.ID, "health-profession-dropdown").click()
print("Select Options")
sleep(1)
options = find_elements(driver, By.TAG_NAME, "li")
for option in options:
    if option.text == "Medical Practitioner":
        option.click()
        sleep(0.5)
        print("Option Clicked!")
        break
   
find_element(driver, By.ID, "predictiveSearchHomeBtn").click()
sleep(2)

people = find_elements(driver, By.CLASS_NAME, "search-results-table-row")   
for person in people:
    cols = person.find_elements(driver, By.CLASS_NAME, "col-span-row")
    for col in cols:
        speciality = col.find_element(driver, By.CLASS_NAME, "speciality")
    