import requests
import json
from pandas.io.json import json_normalize
import pandas as pd
import os
import pyodbc
from pathlib import Path
import smtplib
import time
import datetime  
import numpy as np
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.interval import IntervalTrigger
from dateutil import tz
from io import StringIO

def time_change(a):
    t2 = time.strptime(a, "%Y-%m-%d %H:%M:%S.%f") 
    #otherStyleTime = time.strftime("%Y/%m/%d %H:%M:%S", t2)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", t2)
    return otherStyleTime

def get_battery_level(voltage_now, max_capactiy, min_capacity):
    battery_level = (voltage_now-min_capacity)/(max_capactiy-min_capacity)*100    
    battery_level = int(battery_level)
    
    return battery_level

def get_true_angle(realtime_value ,zero_point, applicable):
    if applicable == 'Y':
        true_angle = np.round(abs(float(realtime_value)-float(zero_point)),3)
        return true_angle
    else:
        return '-'

def water_level(high,water8,water3,mode):
    if mode=='地下水位':
        level=np.round(high-water3,3)
    else:
        level=np.round(high-water8,3)	
    return level

def t_water_level(water8,water3,mode):
    if mode=='地下水位':
        level=np.round(water3,3)
    else:
        level=np.round(water8,3)	
    return level
    
def check_control_value(typename, realtime_value, prewarning_value, alert_value, applicable):
    if typename == '水位':
        if applicable == 'Y':
            warning_check_message = '正常'
            if float(realtime_value) <= float(prewarning_value):
                warning_check_message = '已達預警值'
            if float(realtime_value) <= float(alert_value):
                warning_check_message = '已達警戒值'
            return warning_check_message
        else:
            return '-'
    else:
        if applicable == 'Y':
            warning_check_message = '正常'
            if float(realtime_value) >= float(prewarning_value):
                warning_check_message = '已達預警值'
            if float(realtime_value) >= float(alert_value):
                warning_check_message = '已達警戒值'
            return warning_check_message
        else:
            return '-'    
            
def check_angle_value(realtime_value, prewarning_value, alert_value, applicable):
    if applicable == 'Y':
        warning_check_message = '正常'
        if float(realtime_value) >= float(prewarning_value):
            warning_check_message = '已達預警值'
        if float(realtime_value) >= float(alert_value): 
            warning_check_message = '已達警戒值'
        return warning_check_message
    else:
        return '-'

def typechange(x):
    ii = int(x)
    ss = str(ii)
    return ss

SYS_path = os.path.dirname(os.path.abspath(__file__))
CHANNEL_LIST = pd.read_excel(SYS_path + "/Channel_List.xlsx", engine = 'openpyxl')
CHANNEL_LIST = CHANNEL_LIST.loc[CHANNEL_LIST['Online'] == 'Y'] 
#print(CHANNEL_LIST)
CHANNEL_LIST['IOT_CHANNEL_ID'] = CHANNEL_LIST['IOT Channel ID'].apply(typechange)
CHANNEL_LIST['TS_ID_DB'] = CHANNEL_LIST['TS_ID(FOR DB)'].apply(typechange)
#print(CHANNEL_LIST['IOT_CHANNEL_ID'])

field = list(range(1,9))


def get_data(a):
    df = pd.DataFrame()
    for i in field:
        my_headers = {'CK': CHANNEL_LIST['Read API Keys'][a]}
        channelid = CHANNEL_LIST['IOT_CHANNEL_ID'][a]

        r = requests.get('https://iot.cht.com.tw/iot/v1/device/'+channelid+'/sensor/'+str(i)+'/rawdata?start=2021-03-16T16%3A00%3A00Z&end=2021-03-18T16%3A12%3A00Z&utcOffset=8', headers = my_headers)

        #print(r.status_code)

        #if r.status_code == requests.codes.ok:
            #print("OK")

        lines = r.text.split('{')
        newlines = []
        for line in lines:
            if len(line.split(',')) != 1:
                newlines.append(line)

        newlines.insert(0, '"id", "deviceID", "time", "lat", "lon", "value", "zzz"')
        data = pd.read_csv(StringIO("\n".join(newlines).replace('"', '')))
        data.columns = ["id", "deviceID", "time", "lat", "lon", "value", "zzz"]
        data = data.astype(str)
        data = data.apply(lambda s: s.str.replace('id:', ''))
        data = data.apply(lambda s: s.str.replace('deviceId:', ''))
        data = data.apply(lambda s: s.str.replace('time:', ''))
        data = data.apply(lambda s: s.str.replace('lat:', ''))
        data = data.apply(lambda s: s.str.replace('lon:', ''))
        data = data.apply(lambda s: s.str.replace('value:', ''))
        data = data.apply(lambda s: s.str.replace('[', ''))
        data = data.apply(lambda s: s.str.replace(']', ''))
        data = data.apply(lambda s: s.str.replace('}', ''))
        data = data.drop(columns = ['lat', 'lon', 'zzz'])
        #print(data)
        df = pd.concat([df, data], axis = 1)
    df.columns = ['d', 'deviceID', 'time', 'ID', 'd', 'd', 'd', 'voltage', 'd', 'd', 'd', 'tilt', 'd', 'd', 'd', 'w_25', 'd', 'd', 'd', 'w_60', 'd', 'd', 'd', 's_v', 'd', 'd', 'd', 'f7', 'd', 'd', 'd', 'f8']
    df = df.drop(columns = 'd')
    df['new_time'] = df['time'].apply(time_change)
    return df 
#df1.columns = ['d', 'deviceID', 'time', 'ID', 'd', 'd', 'd', 'voltage', 'd', 'd', 'd', 'tilt', 'd', 'd', 'd', 'w_25', 'd', 'd', 'd', 'w_60', 'd', 'd', 'd', 's_v', 'd', 'd', 'd', 'f7', 'd', 'd', 'd', 'f8']
#df1 = df1.drop(columns = 'd')
#df1['new_time'] = df1['time'].apply(time_change)
#print(df1)

#print(get_data(0, df1))

df0 = get_data(0)
df1 = get_data(1)
df2 = get_data(2)
df3 = get_data(3)
#print(df1)
#print(df2)
#print(df3)
dbname = Path(SYS_path + '/Database.mdb')

def insert_db(a, df):
    try:
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(dbname))
        cnxn = pyodbc.connect(conn_str)
        cur = cnxn.cursor()
        print('connect to db')

    except pyodbc.Error as e:
        print('Error', e)

    tablename = CHANNEL_LIST['TS_ID_DB'][a]
    #createtable = """CREATE TABLE %s (時間 Date PRIMARY KEY,UTC時間 Memo,entry_id Memo,field1 Memo,field2 Memo,field3 Memo,field4 Memo,field5 Memo,field6 Memo,field7 Memo,field8 Memo)"""%(tablename)
    #cur.execute(createtable)
    #cnxn.commit()

    for index, row in df.iterrows(): 
        with cnxn.cursor() as crsr:
            sql = """\
        INSERT INTO %s (時間,UTC時間,entry_id,field1,field2,field3,field4,field5,field6,field7,field8)
        SELECT ? as 時間, ? AS UTC時間, ? AS entry_id, ? AS field1, ? AS field2, ? AS field3, ? AS field4, ? AS field5, ? AS field6, ? AS field7, ? AS field8
        FROM (SELECT COUNT(*) AS n FROM %s) AS Dual
        WHERE NOT EXISTS (SELECT * FROM %s WHERE 時間 = ?)
        """%(tablename,tablename,tablename)
            params = (str(row['new_time']), str(row["time"]), str(row["ID"]), str(row["ID"]), str(row["voltage"]), str(row["tilt"]), str(row["w_25"]), str(row["w_60"]), str(row["s_v"]), str(row["f7"]), str(row["f8"]), str(row['new_time']))
            crsr.execute(sql, params)
            cnxn.commit()
    print('ID',a,'數據寫入完成')

def sort_out_new_data(a, b ,c ,d):
    latest_result0 = a.tail(1)      
    latest_result1 = b.tail(1)     
    latest_result2 = c.tail(1)      
    latest_result3 = d.tail(1)
    aa = latest_result0.append(latest_result1)
    aa = aa.append(latest_result2)
    aa = aa.append(latest_result3)
    aa = aa.reset_index(drop=True)
    return aa

SYS_var = pd.read_excel(SYS_path + '/Setting.xlsx', engine = 'openpyxl')
Sync_Freq = SYS_var['Sync_Freq(min)'][0]
Sync_Abnormal_Time = int(SYS_var['異常時間準則(小時)'][0])


df_STATUS_TABLE = pd.DataFrame()
df_STATUS_TABLE = sort_out_new_data(df0, df1, df2 ,df3)
df_STATUS_TABLE.columns = ['deviceID', 'created_at', 'field1', 'field2', 'field3', 'field4', 'field5', 'field6', 'field7', 'field8', 'localtime']

run_status = []
for i in range(0,4):    
    now_Sync = datetime.datetime.now()  
    data_lag = now_Sync - datetime.datetime.strptime(str(df_STATUS_TABLE['created_at'][i]), '%Y-%m-%d %H:%M:%S.%f')
    data_lag_day = data_lag.days*24
    data_lag_second = data_lag.seconds/3600        
    data_lag = data_lag_day + data_lag_second

    if data_lag < Sync_Abnormal_Time or now_Sync <=  datetime.datetime.strptime(str(df_STATUS_TABLE['created_at'][i]), '%Y-%m-%d %H:%M:%S.%f'):
        run_status.append('正常')
    else:
        run_status.append('異常')
df_STATUS_TABLE['儀器狀態'] = run_status

df_STATUS_TABLE = pd.concat([CHANNEL_LIST, df_STATUS_TABLE], axis = 1)
df_STATUS_TABLE = df_STATUS_TABLE.replace(to_replace='end', value='0', regex=True)



df_STATUS_TABLE["電池百分比"] = df_STATUS_TABLE.apply(lambda row: get_battery_level(float(row['field2']), float(row['電池電壓(最大)']), float(row['電池電壓(最小)'])),axis=1)
df_STATUS_TABLE["水位深度"] = df_STATUS_TABLE.apply(lambda row: water_level(float(row['高程']),float(row['field8']),float(row['field3']),str(row['模組'])),axis=1)
df_STATUS_TABLE["水位高程"] = df_STATUS_TABLE.apply(lambda row: t_water_level(float(row['field8']),float(row['field7']),str(row['模組'])),axis=1)
df_STATUS_TABLE["傾角1"] = df_STATUS_TABLE.apply(lambda row:get_true_angle(row['field3'], row['傾角1初始值'], str(row['傾角1判定'])) ,axis=1)
df_STATUS_TABLE["傾角2"] = df_STATUS_TABLE.apply(lambda row:get_true_angle(row['field4'], row['傾角2初始值'], str(row['傾角2判定'])) ,axis=1)

df_STATUS_TABLE["水位管理值判定"] = df_STATUS_TABLE.apply(lambda row: check_control_value('水位',float(row['水位深度']), float(row['預警值(水位)']), float(row['警戒值(水位)']), str(row['水位判定'])),axis=1)
df_STATUS_TABLE["傾角1管理值判定"] = df_STATUS_TABLE.apply(lambda row: check_angle_value(row['傾角1'], row['預警值(傾角)'], row['警戒值(傾角)'], str(row['傾角1判定'])),axis=1)
df_STATUS_TABLE["傾角2管理值判定"] = df_STATUS_TABLE.apply(lambda row: check_angle_value(row['傾角2'], row['預警值(傾角)'], row['警戒值(傾角)'],  str(row['傾角2判定'])),axis=1)
print(df_STATUS_TABLE)
STATUS_TABLE_Path = SYS_path + '/STATUS_TABLE_OUTPUT/' + 'Monitoring_Status.xlsx'   
writer = pd.ExcelWriter(STATUS_TABLE_Path)
df_STATUS_TABLE.to_excel(writer,'Sheet1')
writer.save()  








