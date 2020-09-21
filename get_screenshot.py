import os
import time
import re  
from PIL import ImageGrab
import win32gui
import pyautogui
import subprocess

DIR_PATH = r"U:\EPFR Update\AutoUpdate Crawler\data"
lst = [i for i in os.listdir(DIR_PATH)]

popup = subprocess.Popen(fr'explorer /select, {DIR_PATH}\{lst[0]}')
time.sleep(3)
window = pyautogui.getWindowsWithTitle('Data')
hwnd = int(re.search(r'[0-9]{3,}', str(window)).group(0))
print(hwnd)

win32gui.SetForegroundWindow(hwnd)
bbox = win32gui.GetWindowRect(hwnd)
img = ImageGrab.grab(bbox)
# img.show()
img.save(r"U:\EPFR Update\AutoUpdate Crawler\screenshot.jpg", 'JPEG')




