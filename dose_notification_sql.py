import sqlite3
import pandas as pd
import openpyxl
import win32com.client as win32
import psutil
import py
import subprocess

# This is a test of the github repo on home computer

# Check if outlook is open.  If not, open it.
for item in psutil.pids():
    for item in psutil.pids():
        p = psutil.Process(item)
        flag = (p.name() == "OUTLOOK.EXE")
        if flag:
            break

    if flag:
        pass
    else:
        try:
            os.startfile("outlook")
            #subprocess.call(['C:\Program Files\Microsoft Office\Office16\Outlook.exe'])
            #os.system("C:\Program Files\Microsoft Office\Office16\Outlook.exe")
        except:
            print("Outlook didn't open successfully")


# path for local database
fileDb = py.path.local(r"C:\Users\clahn\AppData\Local\Continuum\anaconda3"
                       "\envs\env2.7\Lib\site-packages\openrem"
                       "\openremproject\openrem.db")

# Connect to the database. Need .strpath to work.
db = sqlite3.connect(fileDb.strpath)

# function that sends email


def send_notification():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # mail.To = getname
    mail.To = emailname
    mail.Subject = "Dose Notification Trigger"
    mail.body = ("Hello, \r\n \r\nThis is an automated message.  No reply is necessary."
                 "  \r\n \r\nAn exam was performed that exceeded our dose Notification limits.  \r\n \r\nExam: "
                 + protocol + "\r\n \r\nUID: " + uid + "\r\n \r\nCTDI: " + ctdi)
    mail.send



# selects data from database.  LIMIT will  limit results to specified number.
queries = """
SELECT acquisition_protocol as protocol, mean_ctdivol as ctdi, irradiation_event_uid as uid
FROM remapp_ctirradiationeventdata LIMIT 500
"""

# pandas dataframe
df = pd.read_sql_query(queries, db)
df['protocol'] = df['protocol'].astype(str)


# TODO: write a function that takes the uid and finds exam info: acc, location, etc.


# function creates a mask dataframe of single study type.
# looks for ctdi values above a set threshold.
# appends outlier data to a file and emails the physics email with study data.
def dose_limit(exam, limit):
    df2 = df[df['protocol'].str.contains(exam, case=False)]

    for idx, row in df2.iterrows():
        if row.at['ctdi'] > limit:
            # list for adding data to spreadsheet for tracking notifications.
            nt = []
            # global allowed for variables below to be called in outlook functions.
            # there is probably a better way to do this but this is all I know how to do right now.
            global emailname
            global protocol
            global uid
            global ctdi
            # TODO: change to physics@sanfordhealth.org
            emailname = "christopher.lahn@sanfordhealth.org"
            protocol = str(row.at["protocol"])
            nt.append(protocol)
            uid = str(row.at['uid'])
            nt.append(uid)
            ctdi = str(row.at['ctdi'])
            nt.append(ctdi)
            # write the notifications to a file.
            # TODO move file to a permanent place
            wb = openpyxl.load_workbook(r'W:\SHARE8 Physics\Software\python\scripts\clahn\sql dose limit notifications.xlsx')
            sheet = wb['Sheet1']
            # check if UID is already in file.  If so, pass.  If not, append and send notification.
            oldUid = []
            for col in sheet['B']:
                oldUid.append(col.value)
            if uid in oldUid:
                pass
            else:
                sheet.append(nt)
                wb.save(r'W:\SHARE8 Physics\Software\python\scripts\clahn\sql dose limit notifications.xlsx')
                wb.close()
                # calls the function that sends the email with these variables data.
                send_notification()
                wb.close()
                continue




dose_limit('cta', 30)
dose_limit('aaa', 30)
dose_limit('l-spine', 30)
db.close()
