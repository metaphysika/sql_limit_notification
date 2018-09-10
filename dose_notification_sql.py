import sqlite3
import pandas as pd
import openpyxl
import win32com.client as win32
import psutil
import os
import py
import subprocess
from emailsender import *

'''

TODO:  confine to specic dates so when it queries the database, it doesn't pull the full database
do this in the call to the database in building your dataframe.
'''

is_email = True

# path for local database
fileDb = py.path.local(r"C:\Users\clahn\Desktop\openrem.db")


# make a copy of databse file on my computer.
# This script will then perform operations on that file.
if fileDb.isfile():
    fileDb.remove()
py.path.local(r'W:\SHARE8 Physics\Software\python\data\openrem\openrem.db').copy(fileDb)

# Checks if outlook is open.  If not, opens it.
EmailSender().check_outlook()

# Connect to the database. Need .strpath to work.
db = sqlite3.connect(fileDb.strpath)

# selects data from database.  LIMIT will  limit results to specified number.
queries = """
SELECT acquisition_protocol as protocol, mean_ctdivol as ctdi, irradiation_event_uid as uid
FROM remapp_ctirradiationeventdata
"""

# pandas dataframe
df = pd.read_sql_query(queries, db)
df['protocol'] = df['protocol'].astype(str)


# function that takes the uid and finds exam accession number.
def get_accession(uid):
    uidrow = db.cursor().execute(f"SELECT ct_radiation_dose_id "
                                 f"FROM remapp_ctirradiationeventdata "
                                 f"WHERE irradiation_event_uid=?", (uid,)).fetchone()[0]
    ctdoseid = db.cursor().execute(f"SELECT general_study_module_attributes_id "
                                   f"FROM remapp_ctradiationdose "
                                   f"WHERE id=?", (uidrow,)).fetchone()[0]
    accnum = db.cursor().execute(f"SELECT accession_number "
                                 f"FROM remapp_generalstudymoduleattr "
                                 f"WHERE id=?", (ctdoseid,)).fetchone()[0]
    return accnum


def get_examdate(uid):
    uidrow = db.cursor().execute(f"SELECT ct_radiation_dose_id "
                                 f"FROM remapp_ctirradiationeventdata "
                                 f"WHERE irradiation_event_uid=?", (uid,)).fetchone()[0]
    raddate = db.cursor().execute(f"SELECT start_of_xray_irradiation "
                                  f"FROM remapp_ctradiationdose "
                                  f"WHERE id=?", (uidrow,)).fetchone()[0]
    return raddate

# function that takes the uid and finds site location.


def get_site(uid):
    uidrow = db.cursor().execute(f"SELECT ct_radiation_dose_id "
                                 f"FROM remapp_ctirradiationeventdata "
                                 f"WHERE irradiation_event_uid=?", (uid,)).fetchone()[0]
    ctdoseid = db.cursor().execute(f"SELECT general_study_module_attributes_id "
                                   f"FROM remapp_ctradiationdose "
                                   f"WHERE id=?", (uidrow,)).fetchone()[0]
    site = db.cursor().execute(f"SELECT institution_name "
                               f"FROM remapp_generalequipmentmoduleattr "
                               f"WHERE general_study_module_attributes_id=?", (ctdoseid,)).fetchone()[0]
    return site

# function that takes the uid and finds station name.


def get_station(uid):
    uidrow = db.cursor().execute(f"SELECT ct_radiation_dose_id "
                                 f"FROM remapp_ctirradiationeventdata "
                                 f"WHERE irradiation_event_uid=?", (uid,)).fetchone()[0]
    ctdoseid = db.cursor().execute(f"SELECT general_study_module_attributes_id "
                                   f"FROM remapp_ctradiationdose "
                                   f"WHERE id=?", (uidrow,)).fetchone()[0]
    station = db.cursor().execute(f"SELECT station_name "
                                  f"FROM remapp_generalequipmentmoduleattr "
                                  f"WHERE general_study_module_attributes_id=?", (ctdoseid,)).fetchone()[0]
    return station


def scanner_alert_limit(uid):
    uidrow = db.cursor().execute(f"SELECT id "
                                 f"FROM remapp_ctirradiationeventdata "
                                 f"WHERE irradiation_event_uid=?", (uid,)).fetchone()[0]
    scanalert = db.cursor().execute(f"SELECT ctdivol_notification_value "
                                    f"FROM remapp_ctdosecheckdetails "
                                    f"WHERE ct_irradiation_event_data_id=?", (uidrow,)).fetchone()[0]
    return scanalert

# function creates a mask dataframe of single study type.
# looks for ctdi values above a set threshold.
# appends outlier data to a file and emails the physics email with study data.


def dose_limit(exam, limit):
    df2 = df[df['protocol'].str.lower().str.contains(exam, case=False)]

    for idx, row in df2.iterrows():
        if row.at['ctdi'] > limit:
            # list for adding data to spreadsheet for tracking notifications.
            nt = []
            # global allowed for variables below to be called in outlook functions.
            # there is probably a better way to do this but this is all I know how to do right now.
            # global emailname
            # global protocol
            # global uid
            # global ctdi
            # global acc
            # global studydate
            # global siteadd
            # global stationname
            # global alert_limit
            # TODO: change to physics@sanfordhealth.org
            emailname = "christopher.lahn@sanfordhealth.org"
            protocol = str(row.at["protocol"])
            nt.append(protocol)
            uid = str(row.at['uid'])
            nt.append(uid)
            ctdi = str(row.at['ctdi'])
            nt.append(ctdi)
            alert_limit = str(limit)
            nt.append(alert_limit)
            scanalert = scanner_alert_limit(uid)
            nt.append(scanalert)
            # calls function that matches up uid with accession # in database.
            acc = get_accession(uid)
            nt.append(acc)
            # calls function that matches up uid with beginning of radiation event (study date) in database.
            studydate = get_examdate(uid)
            nt.append(studydate)
            # calls function that matches up uid with Site name in database.
            siteadd = get_site(uid)
            nt.append(siteadd)
            # calls function that matches up uid with station name in database.
            stationname = get_station(uid)
            nt.append(stationname)

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
                # calls the module that sends the email with these variables data.
                # if is_email is true, the email will get sent.  If false, it will not send email.
                if is_email:
                    EmailSender().send_email(emailname, "Dose Notification Trigger",
                                             "Hello, \r\n \r\nThis is an automated message.  No reply is necessary."
                                             "  \r\n \r\nAn exam was performed that exceeded our dose Notification limits.  \r\n \r\nExam: "
                                             + protocol + "\r\n \r\nAccession #: " + acc + "\r\n \r\nCTDI: " + ctdi +
                                             "\r\n \r\nAlert Limit: " + alert_limit + "\r\n \r\nStudy Date: " +
                                             studydate + "\r\n \r\nSite: " + siteadd + "\r\n \r\nStation name: " + stationname)
                else:
                    pass
                wb.close()
                continue


# set exams we are looking for and threshold value here.
dose_limit('cta', 150)
# dose_limit('aaa', 100)
# dose_limit('l-spine', 70)
# dose_limit('neck', 65)
# dose_limit('stone', 40)
db.close()
