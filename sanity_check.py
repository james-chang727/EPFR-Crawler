import os
import pandas as pd
import datetime
import sys 
import datetime as dt
import calendar
import win32com.client  # for sending and getting e-mail

lastFri = dt.date.today()
delta = dt.timedelta(days=1)

#while lastFri.weekday() != calendar.FRIDAY:
 #   lastFri -= delta

dates =  lastFri.strftime('%Y%m%d')

# Check if most up to date data points on closest Wednesday, change when necessary (keep format)
lastWed = dt.date.today() #+ dt.timedelta(days=2)
if lastWed.weekday() != 2 and lastWed.weekday() != 3:
    while lastWed.weekday() != calendar.WEDNESDAY:
        lastWed -= delta
elif lastWed.weekday() == 2:
    lastWed -= dt.timedelta(days=7)
    print('!')
elif lastWed.weekday() == 3:
    lastWed -= dt.timedelta(days=8)
    print('!')

# path location

DATE = str(lastWed)
root = r'C:\Users\jameschang\Desktop\auto-update-countryflow\auto-update-EPFR' #!!!
DIR_PATH = fr"{root}\{dates}"
objects = pd.read_pickle(fr'{root}\record_fail_list_{dates}.pkl')

#send to who?
recipients = ['james.chang@dotcomintl.com', 'miao.lin@dotcomintl.com', 'alan.chen@dotcomintl.com', 'chang880727@gmail.com'] #!!!

"""Function to send email out with given email address, password and recipients included in list"""
def sendEmail(receiver_list,dates,DIR_PATH,root_path,contents):
    #print(receiver_list)
    ## Files to send and their paths
    filenames = os.listdir(DIR_PATH)
    print(filenames)
    for (i,rec) in enumerate(receiver_list):
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0) 
        mail.To = rec
        print(rec)
        mail.Subject = 'EPFR Autoupdate'
        
        content = contents
        mail.Body = content
        try:
            for filename in filenames:
                SourcePathName  = fr"{DIR_PATH}\{filename}"
                mail.Attachments.Add(SourcePathName)        
            attachment = mail.Attachments.Add(fr"{root_path}\screenshot\screenshot_{dates}.jpg")
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
            mail.HTMLBody = "<html><body>"+content+" <img src=""cid:MyId1""></body></html>"
        except:
            content = contents + "fail to attach files!"
            mail.HTMLBody = "<html><body>"+content+" <img src=""cid:MyId1""></body></html>"
        #mail.Attachments.Add('U:\EPFR auto update\AutoUpdate Crawler\screenshot.JPG')
        #mail.HTMLBody = content + '<br><img src="U:\EPFR auto update\AutoUpdate Crawler\screenshot.JPG">'
        mail.Send()

# Sanity Check if data dates match
lens = 0
mail_content = ''
for file in os.listdir(DIR_PATH):
    path = fr'{DIR_PATH}\{file}'
    # print(path)
    df = pd.read_excel(path)
    df['Date'] = df['Date'].apply(lambda x: x.strftime('%Y-%m-%d'))
    # print(df['Date'].unique())
    # print(DATE)
    if list(df['Date'].unique()) == [DATE]:
        print(f"{file}: Date Correct!")
        lens += 1
if lens == 8 and len(str(objects)) == 2:
    print("\nAll files correct!")
    mail_content = 'All files checked and correct!'
elif lens != 8:
    print("\nPlease rerun crawler_v2.py to get correct data!")
    mail_content = 'Something wrong! The number of files are not right \n 請手動下載沒有成功的檔案' + str(objects)
elif lens == 8 and objects == 'no screenshot shown!':
    print("\nPlease rerun get_screenshot.py to get correct data!")
    mail_content = 'Something wrong! ' + str(objects)
else:
    print("\nPlease rerun crawler_v2.py to get correct data!")
    mail_content = 'Something wrong!' + str(objects)



# Change to outbox email address and add recipients when needed
# print(object)
content = ''
try:
    sendEmail(recipients,dates,DIR_PATH,root,mail_content)
    print('\nSuccess: email sent\n')
except:
    print('\nFail: email not sent\n')
finally:
    sys.exit()



