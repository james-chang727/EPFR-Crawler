import os
import glob
from datetime import datetime

now = datetime.now()
date = now.strftime('%Y%m%d')
# Change default directory to desired folder path if needed
PREFERENCES = {"download.default_directory": r"D:\Tarobo Training Materials\EPFR_crawler_project\EPFR-Crawler\data"
              ,"safebrowsing.enabled": "False"}


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

file_rename(fund_rename_list + country_rename_list)