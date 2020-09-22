import os
import pandas as pd
import datetime 
from datetime import datetime
import datetime as dt
import calendar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email import encoders

# Check if most up to date data points on closest Wednesday, change when necessary (keep format)
lastWed = dt.date.today() + dt.timedelta(days=2)
delta = dt.timedelta(days=1)

if lastWed.weekday() != 3:
    while lastWed.weekday() != calendar.WEDNESDAY:
        lastWed -= delta
else:
    lastWed -= dt.timedelta(days=8)

DATE = str(lastWed)
DIR_PATH = r"U:\EPFR Update\AutoUpdate Crawler\data"

# Sanity Check if data dates match
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
else:
    print("\nPlease rerun crawler.py to get correct data!")
    quit   

"""Function to send email out with given email address, password and recipients included in list"""
def sendEmail(EMAIL, PW, RECIPIENTS):
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = ', '.join(RECIPIENTS)
    msg['Subject'] = 'EPFR Autoupdate'
    body = 'All files checked and correct!'
    msg.attach(MIMEText(body, 'plain'))

    # Attach Screenshot,, still need tweaking 
    fp = open(r"screenshot.jpg", 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)

    ## Files to send and their paths
    filenames = os.listdir(DIR_PATH)

    for filename in filenames:
        SourcePathName  = fr"{DIR_PATH}\{filename}"

        ## Attachment of files to email
        attachment = open(SourcePathName, 'rb')
        part = MIMEBase('application', "octet-stream")
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)  ### put your relevant SMTP here
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(EMAIL, PW)  ### if applicable
    server.sendmail(msg['From'], RECIPIENTS, msg.as_string())
    server.quit()


# Change to outbox email address and type in your password in the terminal 
# Add recipients when needed
try:
    pw = input('\nPlease type in your password.')
    recipients = ['james.chang@dotcomintl.com', 'miao.lin@dotcomintl.com', 'alan.chen@dotcomintl.com', 'itdata@ugfunds.com']
    sendEmail('chang880727@gmail.com', pw, recipients) 
    print('\nSuccess: email sent\n')
except:
    print('\nFail: email not sent\n')




