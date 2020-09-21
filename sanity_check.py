import os
import pandas as pd
import datetime 
from datetime import datetime 

# Check if most up to date data points on closest Wednesday, change when necessary (keep format)
DATE = '2020-09-16'
DIR_PATH = r"U:\EPFR Update\AutoUpdate Crawler\data"

# Sanity Check if on data dates
len = 0
for file in os.listdir(DIR_PATH):
    path = fr'{DIR_PATH}\{file}'
    df = pd.read_excel(path)
    df['Date'] = df['Date'].apply(lambda x: x.strftime('%Y-%m-%d'))

    if list(df['Date'].unique()) == [DATE]:
        print(f"{file}: Date Correct!")
        len += 1
    if len == 8:
        print("\nAll files correct!")