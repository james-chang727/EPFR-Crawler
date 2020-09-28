from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import datetime as dt
import calendar
import glob
import os
import pickle
import sys

# Would normally be executed on Friday, if not please change date to Friday in YYYYmmdd format
lastFri = dt.date.today()
delta = dt.timedelta(days=1)

#while lastFri.weekday() != calendar.FRIDAY:
#    lastFri -= delta

date =  lastFri.strftime('%Y%m%d')

# Change chromedriver PATH and download default_directory to desired folder path if needed
PATH = r'C:\Users\jameschang\Desktop\chromedriver.exe'  #the place your chromedriver are at #!!!!
root = r'C:\Users\jameschang\Desktop\auto-update-countryflow\auto-update-EPFR' #root #!!!!
datapath = fr"{root}\data" #the place data would be downloaded to
newpath = fr"{root}\{date}" #the place data finally would be moved to
if not os.path.exists(newpath):
    os.makedirs(newpath)
if not os.path.exists(datapath):
    os.makedirs(datapath)

#check if folder is empty or not, if not, end the process and send notification emails
record_lst = []    
if len(os.listdir(fr'{newpath}')) > 0 or len(os.listdir(fr'{datapath}')) > 0:
    record_lst.append("please keep the folder empty!")
    print("please keep the folder empty!")
    with open(fr'{root}\record_fail_list_'+date+'.pkl', 'wb') as f:
        pickle.dump(record_lst, f)
    sys.exit()

# start to crawl    
PREFERENCES = {"download.default_directory": fr'{datapath}',
              "safebrowsing.enabled": "False"}
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", PREFERENCES)
options.add_argument("--start-maximized")

driver = webdriver.Chrome(executable_path=PATH, options=options)
driver.get('http://next.epfrglobal.com')
driver.execute_script("document.body.style.transform='scale(0.6)'")
time.sleep(3)

#EPFR email and pw
search_email = driver.find_element_by_name("email")
search_pw = driver.find_element_by_name("password")
search_email.send_keys('mel@dotcomintl.com')
search_pw.send_keys('epfrdata')
search_pw.send_keys(Keys.RETURN)


name_list = [f"{date}_EPFROutput_weekly_dataset fund flows_asset universe all_investor type all_domicile all_level 3.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset fund flows_asset universe ETF_investor type all_domicile all_level 3.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset fund flows.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type active_domicile all_level 3.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type all_domicile all_level 3.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe all_investor type all_domicile all_level 3_institution.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset country flows_asset universe ETF_investor type all_domicile all_level 3.xlsx"
            ,f"{date}_EPFROutput_weekly_dataset country flows.xlsx"
            ]

def file_rename(rename_file,date,loc):
    PATH = PREFERENCES['download.default_directory']
    NEWPATH = newpath
    #print(PATH)
    #print(NEWPATH)
    files_path = os.path.join(PATH, '*')
    files = sorted(glob.iglob(files_path), key=os.path.getctime, reverse=False)
    #if len(rename_file) != 8:
    #    print('wrong file number')
    #else:
        #for j in range(len(rename_file)):
    dest = fr"{NEWPATH}\{rename_file}"
    print(dest)
    print(files[loc])
    os.rename(files[loc], dest)

def download_EPFR_excel(driver, ul_num, li_num, time1, time2, time3, name_list, date, loc, order):
    #driver.refresh()
    flows = WebDriverWait(driver, time1).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='saved-report-folder-list']/div[{ul_num}]/div[1]")))
    flows.click()
    file = WebDriverWait(driver, time2).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[1]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/ul[{ul_num}]/li[{li_num}]/div[1]/div[1]")))
    file.click()
    excel = WebDriverWait(driver, time3).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='chart-table_wrapper']/div[1]/button[2]")))
    # text = driver.find_element_by_xpath('//*[@id="chart-table"]/tbody/tr[1]/td[1]').text
    excel.click()

    time.sleep(time3)
    driver.refresh()
    
    file_rename(name_list[order], date, loc)

#html tag number
ul_lst = [1, 1, 1, 2, 2, 2, 2, 2]
li_lst = [9, 10, 8, 3, 4, 6, 7, 8]


try:
    j = 0
    fail = 0
    for ul, li in zip(ul_lst, li_lst):
        print(ul)
        print(li)
        
        try:
            download_EPFR_excel(driver, ul, li, 20, 20, 50, name_list, date, 0, j)
            j +=1
        except:
            try:
                download_EPFR_excel(driver, ul, li, 30, 30, 60, name_list, date, 0, j)
                j +=1
            except:
                try:
                    download_EPFR_excel(driver, ul, li, 40, 40, 70, name_list, date, 0, j)
                    j +=1
                except:
                    record_lst.append(name_list[j])
                    fail += 1
                    j += 1
                    print("fail!")
                    pass
        #print(j)
        
finally:
    #pass
    driver.quit()
    with open(fr'{root}\record_fail_list_'+date+'.pkl', 'wb') as f:
        pickle.dump(record_lst, f)
    #get_screenshot.screenshot()
    #sanity_check.send_mail_test()
    


