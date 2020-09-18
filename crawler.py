import selenium as se
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

PATH = r'C:\Users\jameschang\Desktop\chromedriver.exe'
driver = webdriver.Chrome(PATH)

driver.get('http://next.epfrglobal.com')

try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "myDynamicElement"))
    )
finally:
    driver.quit()

