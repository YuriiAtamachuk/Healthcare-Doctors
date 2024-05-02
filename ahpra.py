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


num = []
people = find_elements(driver, By.CLASS_NAME, "search-results-table-row")  
for person in people:
    others = person.find_elements(By.CLASS_NAME, "search-results-table-col")
    person_data = person.find_element(By.XPATH, ".//div/div[3]")
    specialities = person_data.find_elements(By.CLASS_NAME, "col-span-row")
    for speciality in specialities:
        speciality_txt = speciality.find_element(By.CLASS_NAME, "speciality").text
        words = ["Radiology", "Radiation"]
        if any(word in speciality_txt for word in words):
            print(others[0].text)
            num.append(people.index(person))
        else:
            pass
print(num)

for i in num:
    driver.get(url)
    sleep(2)

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
    select_person = people[i]
    topersondata = select_person.find_element(By.XPATH, ".//div/div[1]/div[2]/p/a")
    topersondata.click()
    print("Clicked!")
    sleep(0.5)
    dr_name = find_element(driver, By.CLASS_NAME, "practitioner-name").text
    print(dr_name)
    all_data = find_elements(driver, By.CLASS_NAME, "section-row")
    for data in all_data:
        data_txt = data.find_element(By.XPATH, ".//div[1]").text
        if data_txt == "Profession":
            profession = data.find_element(By.XPATH, ".//div[2]").text
            print(profession)
        elif data_txt == "Registration number":
            ahpra_reg_num = data.find_element(By.XPATH, ".//div[2]").text
            print(ahpra_reg_num)
        elif data_txt == "Registration status":
            reg_status = data.find_element(By.XPATH, ".//div[2]").text
            print(reg_status)
        elif data_txt == 'Conditions\nCondition':
            condition = data.find_element(By.XPATH, ".//div[2]").text
            print(condition)
            # condition = find_element(driver, By.CSS_SELECTOR, "#page-body > div.practitioner-detail-body > div:nth-child(1) > div:nth-child(6) > div.col-xs-12.col-sm-4.col-md-3 > div").text
            # print(condition)
          