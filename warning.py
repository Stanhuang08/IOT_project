# -*- coding: utf-8 -*-
"""
Created on Wed Oct  7 17:22:30 2020

@author: Stan
"""
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
import time
from pathlib import Path
import pyodbc




def check(df, col):
    warning = df.index[df[col] == '已達預警值'].tolist()
    alert = df.index[df[col] == '已達警戒值'].tolist()
    res = warning + alert
    result = df.loc[res, ['Channel ID','模組', col]]
    return result

def send(adress, result):
    if result.empty == False:    
       send_gmail(adress, result)
       

def send_gmail(email, result):
    gmail_user = 'ihmtmonitor@gmail.com'
    gmail_password = 'asml1qaz@WSX'
    mail_server = 'smtp.gmail.com'
    sender = gmail_user
    to_addr = email
    receiver = to_addr.split(',')
    subject = '阿里山五彎仔即時邊坡擋土監測平台預警通知'

    body = result.to_string()
    
    
    msg = MIMEText(body.encode('utf-8'), _charset='utf-8')  # Use utf-8 instead ascii.
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ",".join(receiver)

    s = smtplib.SMTP_SSL(mail_server,465)
    s.ehlo()
    s.login(gmail_user, gmail_password)
    s.sendmail(sender, receiver, msg.as_string())
    s.close()

       

def timer_reload(n):
    while True: 
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        df_STATUS_TABLE = pd.read_excel(STATUS_TABLE_Path)
        df_STATUS_TABLE = df_STATUS_TABLE.replace(to_replace='end', value='0', regex=True)
        df_STATUS_TABLE['field1'] = df_STATUS_TABLE.apply(lambda row: str(int(row['field1'])),axis=1)
        df_STATUS_TABLE['Channel ID'] = 'ID' + df_STATUS_TABLE['field1']
        ppp =  df_STATUS_TABLE[['localtime','Channel ID','模組','儀器狀態','水位管理值判定','傾角1管理值判定','傾角2管理值判定']]
        device_eror = ppp.index[ppp['儀器狀態'] == '異常'].tolist() 
        start = ppp.drop(device_eror)
        water = check(start, '水位管理值判定')
        tilt_1 = check(start, '傾角1管理值判定')
        tilt_2 = check(start, '傾角2管理值判定')
        send(mailgo, water)
        send(mailgo, tilt_1)
        send(mailgo, tilt_2)
        
        

        time.sleep(n)






SYS_path = os.path.dirname(os.path.abspath(__file__))
STATUS_TABLE_Path = SYS_path + '/STATUS_TABLE_OUTPUT/' + 'Monitoring_Status.xlsx'

USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
cnxn = pyodbc.connect(conn_str)
df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
df_NOTIFICATION = df[df['資料通知'] == '開啟']
toaddr = list(df_NOTIFICATION['帳號'])
mailgo = ",".join(toaddr)





timer_reload(259200)