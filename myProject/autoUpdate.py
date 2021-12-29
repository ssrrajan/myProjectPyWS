#Implicit wait  -
#Explicit Wait
import time

from selenium import webdriver
#pause the test for few seconds using Time class
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(executable_path="C:\\chromedriver.exe")
driver.maximize_window()
driver.implicitly_wait(5)

driver.get("http://rndprojects.precisionit.co.in/openproject/")

username = "PG0046"
pwd = "WEL(0^^E"

driver.find_element(By.XPATH, value="//div[@id='login-form']/form/div[1]/div/span/input").send_keys(username)
driver.find_element(By.XPATH, value="//div[@id='login-form']/form/div[2]/div/span/input").send_keys(pwd)
driver.find_element(By.XPATH, value="//div[@id='login-form']/form/input[4]").click()

driver.find_element(By.LINK_TEXT, value="Select a project").click()
driver.find_element(By.LINK_TEXT, value="» InnaITAadhaar Test Repositories").click()
driver.find_element(By.ID, value="main-menu-work-packages").click()
# driver.find_element(By.XPATH, value="//div[@class = 'toolbar']/ul/li[1]").click()
pglist = driver.find_elements(By.XPATH, value="//div[@class = 'pagination']/div/ul/li")

print(len(pglist))

for pg in pglist:
    print(pg.text)
    if pg.text == "300":
        pg.click()
        break

wplist = driver.find_elements(By.CLASS_NAME, value="results-tbody work-package--results-tbody")

for wp in wplist:
    print(wp)

# dropdown-menu

# print(list)
# for ls in list:
#     print(ls.text)
#     time.sleep(2)
#     if ls.text == "» InnaITAadhaar Modules":
#         break