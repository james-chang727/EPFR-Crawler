import os
import time
import re  
from PIL import ImageGrab
import win32gui
import pyautogui
import subprocess
import datetime as dt
import calendar
import pickle


lastFri = dt.date.today()
delta = dt.timedelta(days=1)
#while lastFri.weekday() != calendar.FRIDAY:
#    lastFri -= delta

date =  lastFri.strftime('%Y%m%d')
root = r'C:\Users\jameschang\Desktop\auto-update-countryflow\auto-update-EPFR' #!!!
screenshotpath = fr"{root}\screenshot"
newpath = fr"{root}\{date}" 
if not os.path.exists(newpath):
    os.makedirs(newpath)
if not os.path.exists(screenshotpath):
    os.makedirs(screenshotpath)

DIR_PATH = fr'{newpath}'
lst = [i for i in os.listdir(DIR_PATH)]
#string_test = fr'explorer /select,"C:\Users\miaolin\Desktop\auto-update-countryflow\auto-update-EPFR\{date}"'
try:
    popup = subprocess.Popen(fr'explorer /select, "{DIR_PATH}\{lst[0]}"')
    time.sleep(3)
    window = pyautogui.getWindowsWithTitle(date)
    hwnd = int(re.search(r'[0-9]{3,}', str(window)).group(0))
    # print(hwnd)
    
    win32gui.SetForegroundWindow(hwnd)
    bbox = win32gui.GetWindowRect(hwnd)
    img = ImageGrab.grab(bbox)
    # img.show()
    img.save(fr"{screenshotpath}\screenshot_{date}.jpg", 'JPEG')
except:
    error_msg = "no screenshot shown!"
    with open(fr'{root}\record_fail_list_{date}.pkl', 'wb') as f:
        pickle.dump(error_msg, f)




