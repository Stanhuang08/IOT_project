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
print('hi')

# 自訂表頭
my_headers = {'CK': 'DKRKZ44S2070YUTSKR'}
# 將自訂表頭加入 GET 請求中
r = requests.get('https://iot.cht.com.tw/iot/v1/device/14417566029/sensor/3/rawdata?start=2021-03-11T16%3A00%3A00Z&end=2021-03-13T16%3A12%3A00Z&utcOffset=8', headers = my_headers)
# 伺服器回應的狀態碼
print(r.status_code)
# 檢查狀態碼是否 OK
if r.status_code == requests.codes.ok:
  print("OK")
# 輸出網頁 HTML 原始碼
#print(r.text)
lines = r.text.split('{')

#for line in lines:
#  print(line)
#print(lines)

newlines = []
for line in lines:
  if len(line.split(',')) != 1:
    newlines.append(line)

newlines.insert(0, '"id", "deviceID", "time", "lat", "lon", "value", "zzz"')



df = pd.read_csv(StringIO("\n".join(newlines).replace('"', '')))
df.columns = ["id", "deviceID", "time", "lat", "lon", "value", "zzz"]
df = df.astype(str)
df = df.apply(lambda s: s.str.replace('id:', ''))
df = df.apply(lambda s: s.str.replace('deviceId:', ''))
df = df.apply(lambda s: s.str.replace('time:', ''))
df = df.apply(lambda s: s.str.replace('lat:', ''))
df = df.apply(lambda s: s.str.replace('lon:', ''))
df = df.apply(lambda s: s.str.replace('value:', ''))
df = df.apply(lambda s: s.str.replace('[', ''))
df = df.apply(lambda s: s.str.replace(']', ''))
df = df.apply(lambda s: s.str.replace('}', ''))
df = df.drop(columns = ['lat', 'lon', 'zzz'])
print(df)





