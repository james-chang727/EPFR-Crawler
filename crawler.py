from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import glob
import os
import pandas as pd

now = datetime.now()
date = now.strftime('%Y%m%d')

# Change chromedriver PATH and download default_directory to desired folder path if needed
PATH = r'C:\Users\jameschang\Desktop\chromedriver.exe'
PREFERENCES = {"download.default_directory": r"U:\EPFR Update\AutoUpdate Crawler\data"
              ,"safebrowsing.enabled": "False"}

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", PREFERENCES)
driver = webdriver.Chrome(executable_path=PATH, options=options)

driver.get('http://next.epfrglobal.com')
time.sleep(3)

search_email = driver.find_element_by_name("email")
search_pw = driver.find_element_by_name("password")
search_email.send_keys('mel@dotcomintl.com')
search_pw.send_keys('dotcom2117')
search_pw.send_keys(Keys.RETURN)

fund_rename_list = [f"{date}_EPFROutput_weekly_dataset fund flows.xlsx"
                   ,f"{date}_EPFROutput_weekly_dataset fund flows_asset universe all_investor type all_domicile all_level 3.xlsx"
                   ,f"{date}_EPFROutput_weekly_dataset fund flows_asset universe ETF_investor type all_domicile all_level 3.xlsx"]

country_rename_list = [f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type active_domicile all_level 3.xlsx"
                      ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type all_domicile all_level 3.xlsx"
                      ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type all_domicile all_level 3_institution.xlsx"
                      ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe ETF_investor type all_domicile all_level 3.xlsx"
                      ,f"{date}_EPFROutput_weekly_dataset country flows.xlsx"]

def file_rename(rename_list):
    PATH = PREFERENCES['download.default_directory']
    files_path = os.path.join(PATH, '*')
    files = sorted(glob.iglob(files_path), key=os.path.getctime, reverse=False)
    for i in range(len(rename_list)):
        dest = fr"{PATH}\{rename_list[i]}"
        os.rename(files[i], dest)
        print(dest)

try:
    # driver.find_element_by_xpath("// *[ @ id = 'saved-report-folder-list'] / ul[1] / li[8] / div[1] / div[1]").click()
    for i in range(8, 11):
        fund_flows = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='saved-report-folder-list']/div[1]/div[1]")))
        fund_flows.click()
        file = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"/ html / body / div[1] / div[2] / div[1] / div / div / div[2] / div / div / div[1] / div[2] / div[2] / ul[1] / li[{i}] / div[1] / div[1]")))
        file.click()
        excel = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='chart-table_wrapper']/div[1]/button[2]")))
        # text = driver.find_element_by_xpath('//*[@id="chart-table"]/tbody/tr[1]/td[1]').text
        excel.click()

        time.sleep(50)
        driver.refresh()

    for j in [3, 4, 6, 7]:
        country_flows = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='saved-report-folder-list']/div[2]/div[1]")))
        country_flows.click()
        file = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"/ html / body / div[1] / div[2] / div[1] / div / div / div[2] / div / div / div[1] / div[2] / div[2] / ul[2] / li[{j}] / div[1] / div[1]")))
        file.click()
        excel = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='chart-table_wrapper']/div[1]/button[2]")))
        excel.click()

        time.sleep(30)
        driver.refresh()

    country_flows = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='saved-report-folder-list']/div[2]/div[1]")))
    country_flows.click()
    file = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"/ html / body / div[1] / div[2] / div[1] / div / div / div[2] / div / div / div[1] / div[2] / div[2] / ul[2] / li[8] / div[1] / div[1]")))
    file.click()
    excel = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='chart-table_wrapper']/div[1]/button[2]")))
    excel.click()

    time.sleep(50)
    driver.refresh()

    file_rename(fund_rename_list + country_rename_list)

finally:
    driver.quit()