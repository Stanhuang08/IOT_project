# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 11:21:21 2019

@author: Makuro
"""

# -*- coding: utf-8 -*-

import dash
from dash.dependencies import Input, Output
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import pandas as pd
import flask
from flask import flash, render_template, request, session, send_from_directory
from plotly import graph_objs as go
import datetime
import pyodbc
from pathlib import Path
from datetime import date
import os
from io import BytesIO
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from werkzeug.security import generate_password_hash, check_password_hash



#################################################################################################################  
SYS_path = os.path.dirname(os.path.abspath(__file__))
SYS_var = pd.read_excel(Path(str(SYS_path)+'/Setting.xlsx'))

Sync_Freq = SYS_var['Sync_Freq(min)'][0]  

System_List = pd.read_excel(SYS_path + "/Channel_List.xlsx")

line_1_color = '#CCFFFF'
line_2_color = '#CCFF99'
aler_line_color = '#FFFF33'
act_line_color = '#FF9966'
allow_line_color = '#FF99CC'

System_List = System_List.loc[System_List['Online'] == 'Y']    

#Channel_LIST = System_List['Channel ID'].unique()
Channel_LIST = System_List['TS_ID(FOR DB)'].unique()
System_List['N座標'] = System_List.apply(lambda row: np.round(float(row['N座標']), 6),axis=1)
System_List['E座標'] = System_List.apply(lambda row: np.round(float(row['E座標']), 6),axis=1)

lats = list(System_List['N座標'])
lons = list(System_List['E座標'])
text = list('ID' + System_List['ID'].astype(str))

server_flask = flask.Flask(__name__)

#RECAPTCHA_ENABLED = True
#RECAPTCHA_SITE_KEY = '6LdxVaYUAAAAAPFo_LyWAHu-OEdXiLSLqsBdbgMf'
#RECAPTCHA_SECRET_KEY = '6LdxVaYUAAAAABUJGEJrqQwoxIqNG9oThAaYkGoz'
#RECAPTCHA_THEME = "dark"
#RECAPTCHA_TYPE = "image"
#RECAPTCHA_SIZE = "compact"
#RECAPTCHA_RTABINDEX = 10
#recaptcha = ReCaptcha(app=server_flask)

app = dash.Dash(__name__, server=server_flask, url_base_pathname='/AL/')
app.title='台20線50.7k處即時邊坡擋土監測平台'
app.config.suppress_callback_exceptions = True
#################################################################################################################    
                          
app.layout =  html.Div(
    [
         html.Div([
            # header
            html.Div([
    
                html.Span("台20線50.7k處即時邊坡擋土監測平台", className='app-title'),
                html.Div(
                    html.Img(src='https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Roundel_of_National_Cheng_Kung_University.svg/1024px-Roundel_of_National_Cheng_Kung_University.svg.png',height="95%")
                    ,style={"float":"right","height":"95%"}),                
                html.Div(
                    html.Img(src='https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Seal_of_Institute_of_Transportation_of_the_Republic_of_China.svg/1200px-Seal_of_Institute_of_Transportation_of_the_Republic_of_China.svg.png',height="95%")
                    ,style={"float":"right","height":"95%"}),

                ],
                className="row header"
                ),
    
            # tabs
            html.Div([
    
                dcc.Tabs(
                    id="tabs",
                    style={"height":"20","verticalAlign":"middle"},
                    value="sys_status_tab",
                )
    
                ],
                className="row tabs_div"
                ),
    
    
    
            # Tab content
            html.Div(id="tab_content", className="row", style={"margin": "2% 3%"}),
            dcc.Interval(
                    id='interval_component_tab',
                    interval=3600*1000, # in milliseconds
                    n_intervals=0
            ),        
    #        html.Link(href="https://use.fontawesome.com/releases/v5.2.0/css/all.css",rel="stylesheet"),
    #        html.Link(href="https://cdn.rawgit.com/plotly/dash-app-stylesheets/2d266c578d2a6e8850ebce48fdb52759b2aef506/stylesheet-oil-and-gas.css",rel="stylesheet"),
    #        html.Link(href="https://fonts.googleapis.com/css?family=Dosis", rel="stylesheet"),
    #        html.Link(href="https://fonts.googleapis.com/css?family=Open+Sans", rel="stylesheet"),
    #        html.Link(href="https://fonts.googleapis.com/css?family=Ubuntu", rel="stylesheet"),
    #        html.Link(href="https://cdn.rawgit.com/amadoukane96/8a8cfdac5d2cecad866952c52a70a50e/raw/cd5a9bf0b30856f4fc7e3812162c74bfc0ebe011/dash_crm.css", rel="stylesheet")
    
        ],
        className="row",
        style={"margin": "0%"},
    ),
            # header
        html.Div([


            html.Span("高等土壤力學實驗室 All rights reserved copyright © 2020 國立成功大學土木工程學系",
                  className="twelve columns indicator_text",
                  style={
                "color": "black",
                'text-align': 'center'
            },),
  
            
            ],
            className="navbar"
            ),
    ])

@app.callback(Output("tabs", "children"), [Input("interval_component_tab", "n_intervals")])
def RENDER_TAB_LIST(n_intervals):
    if session.get('role') == 'super_admin':
        children=[
                    dcc.Tab(label="即時資訊", value="sys_status_tab"),
                    dcc.Tab(label="資料下載", value="history_search_tab"),
                    dcc.Tab(label="使用者管理", value="user_management_tab"),
                    
                ]

    elif session.get('role') == 'admin':
        children=[
                    dcc.Tab(label="即時資訊", value="sys_status_tab"),
                    dcc.Tab(label="資料下載", value="history_search_tab"),
                    
                ]

    else:
        children=[
                    dcc.Tab(label="即時資訊", value="sys_status_tab"),                  
                ]
    return children        
        
@app.callback(Output("tab_content", "children"), [Input("tabs", "value")])
def RENDER_CONTENT(tab):
    if not session.get('logged_in'):
         return html.Div([   html.Br(),
                        html.H1('請先登入後，始可使用本系統功能', style={'textAlign': 'center', 'fontSize': 20}),
                        html.Br(),
                        html.A([html.Button('登入')],
                        href='/')
                    ], style={'textAlign': 'center'}) 
    else:
        if session.get('role')=='super_admin':
            if tab == "sys_status_tab":
                return sys_status_layout
            elif tab == "history_search_tab":
                return history_search_layout
            elif tab == "user_management_tab":
                return user_management_layout
            else:
                return None
        elif session.get('role')=='admin':
            if tab == "sys_status_tab":
                return sys_status_layout
            elif tab == "history_search_tab":
                return history_search_layout
            elif tab == "user_management_tab":
                return html.Div([   html.Br(),
                        html.H1('本功能僅開放給管理者', style={'textAlign': 'center', 'fontSize': 20}),
                        html.Br(),
                        html.A([html.Button('返回即時資訊')],
                        href='/')
                    ], style={'textAlign': 'center'}) 
            else:
                return None
        elif session.get('role')=='user':
            if tab == "sys_status_tab":
                return sys_status_layout
            elif tab == "history_search_tab":
                return html.Div([   html.Br(),
                        html.H1('本功能僅開放給管理者', style={'textAlign': 'center', 'fontSize': 20}),
                        html.Br(),
                        html.A([html.Button('返回即時資訊')],
                        href='/')
                    ], style={'textAlign': 'center'}) 
            elif tab == "user_management_tab":
                return html.Div([   html.Br(),
                        html.H1('本功能僅開放給管理者', style={'textAlign': 'center', 'fontSize': 20}),
                        html.Br(),
                        html.A([html.Button('返回即時資訊')],
                        href='/')
                    ], style={'textAlign': 'center'}) 
            else:
                return None            
#################################################################################################################
STATUS_TABLE_Path = SYS_path + '/STATUS_TABLE_OUTPUT/' + 'Monitoring_Status.xlsx'
df_STATUS_TABLE = pd.read_excel(STATUS_TABLE_Path)
mapbox_access_token = 'pk.eyJ1Ijoid2lsc29uY2hhbzI5MjkiLCJhIjoiY2prd29kbWFqMGIzYjN2bW05emFxYTdhaCJ9.yeiZxq4JlOzCcBSHqca4BQ'

sys_status_layout = html.Div([
        html.Br(),                
        html.A(
            "登出",
            href = 'logout',
            className="button button--primary",
            style={
                "height": "24",
                "background": "##87CEFA",
                "border": "1px solid #FFF",
                "color": "white",
                'float': 'right',
            },
        ),
        html.Br(), 
        html.Div(
                    id="time_value",
                    className="twelve columns indicator_text",
            style={
                "color": "black",
                'float': 'left',
            },
                ),

        html.Br(),
        html.Br(),
###############################################################################
    html.Div([
        html.Div([
		    html.Img(src='https://i.imgur.com/RfxOZrY.png',
		        style={
		        'height':'436px',
			   },
			   className="nine columns chart_div",
				),
			]),
			
		html.Div([
                    html.P('監測儀器分布'),
                    dcc.Graph(
                        id = "mapbox",
                        style={"height": "360px"},
                    ),           
                    dcc.Dropdown(
                    id='BaseMap_type',
                    options=[
                        {'label': '衛星影像圖', 'value': 'satellite-streets'},                        {'label': '基本地圖(深色)', 'value': 'mapbox://styles/wilsonchao2929/cjkwoeat03dxs2rmrkbycrvt9'},
                        {'label': '基本地圖(淺色)', 'value': 'light'},                        {'label': '基本地圖', 'value': 'basic'},
                        {'label': '地形圖', 'value': 'outdoors'},
                   ],
                    value='satellite-streets',
                    placeholder="切換地圖型式",
                    ), 
                            ],
                           style={"height": "400px"},
                            className="three columns chart_div",
                            ),
        ],
		className="row",
		style={"marginTop": "10px","marginBottom": "20px"},
    ),
	
	html.Br(),
	html.Br(),


###############################################################################
    html.Div(
        [
         
                    html.Div([
                        html.P('監測儀器狀態'),
                    dash_table.DataTable(
                        id = 'table_all_list',
            
                        columns=[
                            {"name": '模組編號', "id": 'Channel ID'},
                            {"name": '時間', "id": 'localtime'},
                            {"name": '運轉狀態', "id": '儀器狀態'},
                            {"name": '註解', "id": '模組'},
                            {"name": '水位管理值判定', "id": '水位管理值判定'},
                            {"name": '傾角1管理值判定', "id": '傾角1管理值判定'},
                            {"name": '傾角2管理值判定', "id": '傾角2管理值判定'},
                            ],                    
                        style_as_list_view=True,
                        style_header={
                                        'backgroundColor': 'white',
                                        'fontWeight': 'bold',
                                        'fontSize': '13px',
                                        'color': 'black',
                                        'font-family': '"Microsoft JhengHei',
                                        'overflow': 'hidden',
                                        'textOverflow': 'ellipsis',
                                    },
                        style_cell={'textAlign': 'center',
                                    'fontSize': '13px',
                                    'font-family': '"Microsoft JhengHei'},    
                                    
                        style_data_conditional=[
                            {
                                'if': {
                                    'column_id': '水位管理值判定',
                                    'filter_query': '{水位管理值判定} eq "已達預警值"'
                                },
                                'backgroundColor': '#FFDD55',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '水位管理值判定',
                                    'filter_query': '{水位管理值判定} eq "已達警戒值"'
                                },
                                'backgroundColor': '#FF8888',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '傾角1管理值判定',
                                    'filter_query': '{傾角1管理值判定} eq "已達預警值"'
                                },
                                'backgroundColor': '#FFDD55',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '傾角1管理值判定',
                                    'filter_query': '{傾角1管理值判定} eq "已達警戒值"'
                                },
                                'backgroundColor': '#FF8888',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '傾角2管理值判定',
                                    'filter_query': '{傾角2管理值判定} eq "已達預警值"'
                                },
                                'backgroundColor': '#FFDD55',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '傾角2管理值判定',
                                    'filter_query': '{傾角2管理值判定} eq "已達警戒值"'
                                },
                                'backgroundColor': '#FF8888',
                                'color': 'black',
                            },
                            {
                                'if': {
                                    'column_id': '儀器狀態',
                                    'filter_query': '{儀器狀態} eq "異常"'
                                },
                                'backgroundColor': '#FF8888',
                                'color': 'black',
                            },
                                     
                        ],                     
                            
                        style_table={"height": "400px",'overflowX': 'scroll','overflowY': 'scroll'},
                        
                    ),
                             
                                ],
                                style={"height": "400px"},
                                className="twelve columns chart_div",
                            ),
        ],
        className="row",
        style={"marginTop": "10px","marginBottom": "50px"},
    ),
    
   
###############################################################################
                    
    html.Br(),
    html.Br(),
    
###############################################################################     
    html.Div(
        [
                    html.Div([
                        html.P('最新體積含水量資訊'),
                    dash_table.DataTable(
                        id = 'table_water_content',
            
                        columns=[
                            {"name": '模組編號', "id": 'Channel ID'},
                            {"name": '時間', "id": 'localtime'},                 
                            {"name": '深度25cm處(%)', "id": 'field4'},
                            {"name": '深度60cm處(%)', "id": 'field5'},                            
                            ],                    
                        style_as_list_view=True,
                        style_header={
                                        'backgroundColor': 'white',
                                        'fontWeight': 'bold',
                                        'fontSize': '13px',
                                        'color': 'black',
                                        'font-family': '"Microsoft JhengHei',
                                        'overflow': 'hidden',
                                        'textOverflow': 'ellipsis',
                                    },
                        style_cell={'textAlign': 'center',
                                    'fontSize': '13px',
                                    'font-family': '"Microsoft JhengHei'},    
                            
                        style_table={"height": "200px",'overflowX': 'scroll','overflowY': 'scroll'}
                        
                    ),
                            
                                ],
                                style={"height": "200px"},
                                className="three columns chart_div",
                            ),
                    html.Div([
                        html.P('最新傾角資訊'),
                    dash_table.DataTable(
                        id = 'table_wall_rotation',
            
                        columns=[
                            {"name": '模組編號', "id": 'Channel ID'},
                            {"name": '時間', "id": 'localtime'},                  
                            {"name": '方向1(坡向)(度)', "id": 'field3'},
                            {"name": '方向2(正交於坡向)(度)', "id": 'field4'},
                            {"name": '累積傾角變量(方向1)(度) ', "id": '傾角1'},
                            {"name": '累積傾角變量(方向2) (度)', "id": '傾角2'},
                            {"name": '預警值(度)', "id": '預警值(傾角)'},
                            {"name": '警戒值(度)', "id": '警戒值(傾角)'},
                            ],                    
                        style_as_list_view=True,
                        style_header={
                                        'backgroundColor': 'white',
                                        'fontWeight': 'bold',
                                        'fontSize': '13px',
                                        'color': 'black',
                                        'font-family': '"Microsoft JhengHei',
                                        'overflow': 'hidden',
                                        'textOverflow': 'ellipsis',
                                    },
                        style_cell={'textAlign': 'center',
                                    'fontSize': '13px',
                                    'font-family': '"Microsoft JhengHei'},    
                            
                        style_table={"height": "200px",'overflowX': 'scroll','overflowY': 'scroll'},
                        
                    ),
                             
                                ],
                                style={"height": "200px"},
                                className="nine columns chart_div",
                            ),
        ],
        className="row",
        style={"marginTop": "10px","marginBottom": "50px"},
    ),
###############################################################################
                    
    html.Br(),
    html.Br(),
    
###############################################################################    
html.Div(
        [
                    html.Div([
                       
                        html.Div(id = "tilt_title"),
                        
                        dcc.Graph(
                            id = "tilt_graph",
                            style={"height": "360px"},
                        ),           
                        dcc.Dropdown(
                        id='tilt_id',
                        options=[
                            {'label': 'ID1', 'value': 'ID1'},                         
                            {'label': 'ID2', 'value': 'ID2'},
                            {'label': 'ID3', 'value': 'ID3'},
                        ],
                        value = '-',
                        placeholder="請選擇儀器ID",
                        ), 
                        dcc.Dropdown(
                        id = 'data_num',
                        options=[
                            
                            {'label' : '最近50筆資料', 'value': '最近50筆資料'},
                            {'label' : '最近100筆資料', 'value': '最近100筆資料'},
                            {'label' : '最近200筆資料', 'value': '最近200筆資料'},
                            {'label' : '最近1000筆資料', 'value': '最近1000筆資料'},
                        ],
                        value = '最近100筆資料',
                        placeholder="選擇資料",
                        ),
                             
                                ],
                                style={"height": "400px"},
                                className="six columns chart_div",
                            ),


                    html.Div([
                       
                        html.Div(id = "water_content_title"),
                        
                        dcc.Graph(
                            id = "water_content_graph",
                            style={"height": "360px"},
                        ),           
                        dcc.Dropdown(
                        id='water_content_id',
                        options=[
                            {'label': 'ID1', 'value': 'ID1'},                             
                            {'label': 'ID2', 'value': 'ID2'},
                            {'label': 'ID3', 'value': 'ID3'},
                        ],
                        value = '-',
                        placeholder="請選擇儀器ID",
                        ), 
                        dcc.Dropdown(
                        id = 'waterdata_num',
                        options=[
                            
                            {'label' : '最近50筆資料', 'value': '最近50筆資料'},
                            {'label' : '最近100筆資料', 'value': '最近100筆資料'},
                            {'label' : '最近200筆資料', 'value': '最近200筆資料'},
                            {'label' : '最近1000筆資料', 'value': '最近1000筆資料'},
                        ],
                        value = '最近100筆資料',
                        placeholder="選擇資料",
                        ),
                             
                                ],
                                style={"height": "400px"},
                                className="six columns chart_div",
                            ),


        ],
        className="row",
        style={"marginTop": "10px","marginBottom": "100px"},
    ),



###############################################################################
                    
    html.Br(),
    html.Br(),
    
############################################################################### 
        dcc.Interval(
                id='interval_component_sys_status',
                interval=600*1000, # in milliseconds
                n_intervals=0
        ),
    ],
    className='ten columns offset-by-one'
)
                    
                    
###############################################################################
###############################################################################
@app.callback(
    dash.dependencies.Output('mapbox', 'figure'),
    [dash.dependencies.Input('BaseMap_type', 'value')])
def callback_maptype(maptype):
    figure={
                "data": [
                    dict(
                        type = "scattermapbox",
                        lat = lats,
                        lon = lons,
                        mode = "markers+text",
                        marker=dict(
                        size=17,
                        color='rgb(255, 0, 0)',
                        opacity=0.7),
                        text = text
                    ),
                    dict(
                        type = "scattermapbox",
                        lat = lats,
                        lon = lons,
                        mode = "markers+text",
                        marker=dict(
                        size=8,
                        color='rgb(242, 177, 172)',
                        opacity=0.7),
                        text = text,
                        textfont=dict(
                        family='sans serif',
                        size=18,
                        color='black'
                        )

                    ), 
               ],
                "layout": dict(
                    autosize = True,
                    hovermode = None,
                    showlegend=False,
                    margin = dict(l = 0, r = 0, t = 0, b = 0),
                    mapbox = dict(
                        accesstoken = mapbox_access_token,
                        bearing = 0,	
                        center = dict(lat = 23.072972, lon = 120.544528),
                        style = maptype,
                        pitch = 0,
                        zoom = 16,
                        layers = []
                    )   
                )
            }
    return figure
 
@app.callback(
    Output("table_all_list", "data"),
    [Input('interval_component_sys_status', 'n_intervals')]
)
def table_all_list_df(n):  
    df_STATUS_TABLE = pd.read_excel(STATUS_TABLE_Path)
    df_STATUS_TABLE = df_STATUS_TABLE.replace(to_replace='end', value='0', regex=True)
    df_STATUS_TABLE['field1'] = df_STATUS_TABLE.apply(lambda row: str(int(row['field1'])),axis=1)
    df_STATUS_TABLE['Channel ID'] = 'ID' + df_STATUS_TABLE['field1']
    
    df_STATUS_TABLE = df_STATUS_TABLE[['localtime','Channel ID', '電池電壓(最小)', '電池電壓(最大)', 'field2', 'field3', 'field4','field5','field6','field7','field8','模組','電池百分比','儀器狀態','水位管理值判定','傾角1管理值判定','傾角2管理值判定']]   
    df_STATUS_TABLE['field2'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field2']), 3),axis=1)
    df_STATUS_TABLE['field3'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field3']), 3),axis=1)
    df_STATUS_TABLE['field4'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field4']), 3),axis=1)
    df_STATUS_TABLE['field5'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field5']), 3),axis=1)
    df_STATUS_TABLE['field6'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field6']), 3),axis=1)
    df_STATUS_TABLE['field7'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field7']), 3),axis=1)
    df_STATUS_TABLE['field8'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field8']), 3),axis=1)

    #df_STATUS_TABLE1 = df_STATUS_TABLE[df_STATUS_TABLE['Channel ID'] == 'ID1']
    #df_STATUS_TABLE1 = df_STATUS_TABLE1.reset_index(drop=True)
    
    #df_STATUS_TABLE1['Channel ID'] = 'Gateway'
    #df_STATUS_TABLE1['模組'] = '資料蒐集器'
    #df_STATUS_TABLE1['電池百分比'] = np.round((float(df_STATUS_TABLE1['field2'][0]) - df_STATUS_TABLE1['電池電壓(最小)'][0])/(df_STATUS_TABLE1['電池電壓(最大)'][0] - df_STATUS_TABLE1['電池電壓(最小)'][0])*100, 1)
    #df_STATUS_TABLE = df_STATUS_TABLE.append(df_STATUS_TABLE1)
    
    return (df_STATUS_TABLE.to_dict("rows"))


# updates time
@app.callback(
    Output("time_value", "children"),[Input('interval_component_sys_status', 'n_intervals')]
)
def time_value_callback(n):
   
    return html.P(
        '現在時間：'+datetime.datetime.now().strftime('%Y-%m-%d') + ' ' + datetime.datetime.now().strftime('%H:%M:%S'),
        className="twelve columns indicator_text"
    )
    




@app.callback(
    Output("table_water_content", "data"),
    [Input('interval_component_sys_status', 'n_intervals')]
)
def table_water_content_df(n):  
    df_STATUS_TABLE = pd.read_excel(STATUS_TABLE_Path)
    df_STATUS_TABLE = df_STATUS_TABLE.replace(to_replace='end', value='0', regex=True)
    df_STATUS_TABLE = df_STATUS_TABLE[df_STATUS_TABLE['模組'] == '分層含水量&傾斜儀']
    df_STATUS_TABLE['field1'] = df_STATUS_TABLE.apply(lambda row: str(int(row['field1'])),axis=1)
    df_STATUS_TABLE['Channel ID'] = 'ID' + df_STATUS_TABLE['field1']
    df_STATUS_TABLE = df_STATUS_TABLE[['localtime','Channel ID', 'field2', 'field3', 'field4','field5','field6','field7','field8','模組','電池百分比','儀器狀態']]
    df_STATUS_TABLE['field2'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field2']), 3),axis=1)
    df_STATUS_TABLE['field3'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field3']), 3),axis=1)
    df_STATUS_TABLE['field4'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field4']), 3),axis=1)
    df_STATUS_TABLE['field5'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field5']), 3),axis=1)
    df_STATUS_TABLE['field6'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field6']), 3),axis=1)
    df_STATUS_TABLE['field7'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field7']), 3),axis=1)
    df_STATUS_TABLE['field8'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field8']), 3),axis=1)
    return (df_STATUS_TABLE.to_dict("rows"))    

@app.callback(
    Output("table_wall_rotation", "data"),
    [Input('interval_component_sys_status', 'n_intervals')]
)
def table_wall_rotation_df(n):  
    df_STATUS_TABLE = pd.read_excel(STATUS_TABLE_Path)
    df_STATUS_TABLE = df_STATUS_TABLE.replace(to_replace='end', value='0', regex=True)
    df_STATUS_TABLE = df_STATUS_TABLE[df_STATUS_TABLE['模組'] != '雨量計']
    df_STATUS_TABLE = df_STATUS_TABLE[df_STATUS_TABLE['模組'] != '地下水位']
    df_STATUS_TABLE1 = df_STATUS_TABLE[df_STATUS_TABLE['模組'] != '裂縫計']
    df_STATUS_TABLE1['field4'] = '-' 
#    df_STATUS_TABLE2 = df_STATUS_TABLE[df_STATUS_TABLE['模組'] == '裂縫計']
#    df_STATUS_TABLE2['field4'] = df_STATUS_TABLE2.apply(lambda row: np.round(float(row['field4']), 3),axis=1)
#    df_STATUS_TABLE = df_STATUS_TABLE1.append(df_STATUS_TABLE2)
    df_STATUS_TABLE = df_STATUS_TABLE1
    df_STATUS_TABLE['field1'] = df_STATUS_TABLE.apply(lambda row: str(int(row['field1'])),axis=1)
    df_STATUS_TABLE['Channel ID'] = 'ID' + df_STATUS_TABLE['field1']
    df_STATUS_TABLE = df_STATUS_TABLE[['localtime','Channel ID', 'field2', 'field3', 'field4','field5','field6','field7','field8','模組','電池百分比','儀器狀態','傾角1','傾角2','預警值(傾角)','警戒值(傾角)']]
    df_STATUS_TABLE['field2'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field2']), 3),axis=1)
    df_STATUS_TABLE['field3'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field3']), 3),axis=1)
    df_STATUS_TABLE['field5'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field5']), 3),axis=1)
    df_STATUS_TABLE['field6'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field6']), 3),axis=1)
    df_STATUS_TABLE['field7'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field7']), 3),axis=1)
    df_STATUS_TABLE['field8'] = df_STATUS_TABLE.apply(lambda row: np.round(float(row['field8']), 3),axis=1)
    return (df_STATUS_TABLE.to_dict("rows"))   



######TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
@app.callback(
    dash.dependencies.Output('tilt_graph', 'figure'),
    [dash.dependencies.Input('tilt_id', 'value'),dash.dependencies.Input('data_num', 'value')])
def tilt_graph(tilt_id_input,data_num_input):
    if tilt_id_input == 'ID1':
        id_channel = '26254079528'
        ini_value = System_List.loc[0,"傾角1初始值"]
    elif tilt_id_input == 'ID2':
        id_channel = '26254196776'
        ini_value = System_List.loc[1,"傾角1初始值"]
    elif tilt_id_input == 'ID3':
        id_channel = '26257495976'
        ini_value = System_List.loc[2,"傾角1初始值"]

    DB_PATH = SYS_path+"/DATA/"      
    filename1 = Path(DB_PATH + "Database.mdb")
    
    try:
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(filename1))
        cnxn = pyodbc.connect(conn_str)
    except Exception as e:
        print(e)     

    try:
        tablename= str(id_channel)
        df_input = pd.read_sql("SELECT * FROM %s" %(tablename), cnxn)

    except (BaseException): 
        None    

    
    if data_num_input == '最近1000筆資料':
        x=1000
    elif data_num_input == '最近50筆資料':
        x=50
    elif data_num_input == '最近100筆資料':
        x=100
    elif data_num_input == '最近200筆資料':
        x=200
    df_input = df_input.sort_values(by=['時間'])
    df_input = df_input.tail(x)

    df_input['field3'] = df_input.apply(lambda row: np.round(float(row['field3']), 3),axis=1)

    figure={
                        'data': [go.Scatter(
                        x= df_input['時間'] ,
                        y= df_input['field3']-ini_value,
                        text= df_input['field3']-ini_value,
                        name = tilt_id_input,
                        mode = 'lines',
                        line = dict(color = line_1_color),
                        marker={
                            'size': 10,
                            'opacity': 0.5,
                            'line': {'width': 0.5, 'color': '#45B39D'}
                        }),
                        
                     ],            
                    'layout': go.Layout(  
                        xaxis=dict(
                                        type='date',
                                        title='時間',
                                    ),
                        yaxis={
                            'title': '傾角(Degree)',
                            'type': 'linear'
                        },
                    autosize=True,
                    font=dict(color='black'),
                    titlefont=dict(color='black', size=14),
                    margin=dict(
                        l=75,
                        r=35,
                        b=55,
                        t=45
                    ),
                    hovermode="closest",
                    plot_bgcolor="#666666",
                    paper_bgcolor="#DDDDDD",
                    legend=dict(font=dict(size=14)),
                    )
            }
    return figure

@app.callback(
    dash.dependencies.Output('water_content_graph', 'figure'),
    [dash.dependencies.Input('water_content_id', 'value'),dash.dependencies.Input('waterdata_num', 'value')])
def water_content_graph(water_content_id_input,waterdata_num_input):
    if water_content_id_input == 'ID1':
        id_channel = '26254079528'
    elif water_content_id_input == 'ID2':
        id_channel = '26254196776'
    elif water_content_id_input == 'ID3':
        id_channel = '26257495976'

    DB_PATH = SYS_path+"/DATA/"      
    filename1 = Path(DB_PATH + "Database.mdb")
    
    try:
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(filename1))
        cnxn = pyodbc.connect(conn_str)
    except Exception as e:
        print(e)     

    try:
        tablename= str(id_channel)
        df_input = pd.read_sql("SELECT * FROM %s" %(tablename), cnxn)

    except (BaseException): 
        None    

    
    if waterdata_num_input == '最近1000筆資料':
        x=1000
    elif waterdata_num_input == '最近50筆資料':
        x=50
    elif waterdata_num_input == '最近100筆資料':
        x=100
    elif waterdata_num_input == '最近200筆資料':
        x=200
    df_input = df_input.sort_values(by=['時間'])
    df_input = df_input.tail(x)

    df_input['field4'] = df_input.apply(lambda row: np.round(float(row['field4']), 3),axis=1)
    df_input['field5'] = df_input.apply(lambda row: np.round(float(row['field5']), 3),axis=1)
    figure={
                        'data': [go.Scatter(
                        x= df_input['時間'] ,
                        y= df_input['field4'],
                        text= df_input['field4'],
                        name = water_content_id_input+'(25cm)',
                        mode = 'lines',
                        line = dict(color = line_1_color),
                        marker={
                            'size': 10,
                            'opacity': 0.5,
                            'line': {'width': 0.5, 'color': '#45B39D'}
                        }),
                        
                        go.Scatter(
                        x= df_input['時間'] ,
                        y= df_input['field5'],
                        text= df_input['field5'],
                        name = water_content_id_input+'(60cm)',
                        mode = 'lines',
                        line = dict(color = aler_line_color),
                        marker={
                            'size': 10,
                            'opacity': 0.5,
                            'line': {'width': 0.5, 'color': '#45B39D'}
                        }),
                     ],            
                    'layout': go.Layout(  
                        xaxis=dict(
                                        type='date',
                                        title='時間',
                                    ),
                        yaxis={
                            'title': 'water content',
                            'type': 'linear'
                        },
                    autosize=True,
                    font=dict(color='black'),
                    titlefont=dict(color='black', size=14),
                    margin=dict(
                        l=75,
                        r=35,
                        b=55,
                        t=45
                    ),
                    hovermode="closest",
                    plot_bgcolor="#666666",
                    paper_bgcolor="#DDDDDD",
                    legend=dict(font=dict(size=14)),
                    )
            }
    return figure
# updates time
@app.callback(
    Output("rain_graph_title", "children"),[Input('rain_graph_period', 'value')]
)
def rain_graph_title_callback(rain_graph_period_input):
   
    return html.P(rain_graph_period_input + '時雨量圖')    

# updates time
@app.callback(
    Output("water_level_title", "children"),[Input('water_level_period', 'value')]
)
def water_level_title_callback(water_level_period_input):
   
    return html.P(water_level_period_input + '水位高程變化圖')    
 #update time
@app.callback(
    Output("tilt_title", "children"),[Input('tilt_id', 'value')]
)
def tilt_title_callback(tilt_id_input):
   
    return html.P('邊坡監測-'+tilt_id_input + '傾角變化圖') 

 #update time
@app.callback(
    Output("water_content_title", "children"),[Input('water_content_id', 'value')]
)
def water_content_title_callback(water_content_id_input):
   
    return html.P('邊坡監測-'+water_content_id_input + '體積含水量變化圖') 
       
#################################################################################################################
history_search_layout =html.Div([
        html.Br(),
        html.A(
            "登出",
            href = 'logout',
            className="button button--primary",
            style={
                "height": "24",
                "background": "##87CEFA",
                "border": "1px solid #FFF",
                "color": "white",
                'float': 'right',
            },
        ),
        html.Br(),
        html.Div(
                    id="time_value_history_search",
                    className="twelve columns indicator_text",
            style={
                "color": "black",
                'float': 'left',
            },
                ),

        html.Br(),
        html.Div(
            [
                html.P('欲下載之日期區間:',style={'margin-top': '20'}),  # noqa: E501
                dcc.DatePickerRange(
                    id='date_picker_history_search',
                    min_date_allowed=datetime.datetime(2018, 1, 1),
                    max_date_allowed=datetime.datetime(2049, 12, 31),
                    initial_visible_month=datetime.datetime(2018, 1, 1),
                    end_date=datetime.datetime(2018, 1, 1)
                ),
                   html.Button(
                    id='button_history_search',
                    children='下載',
                    n_clicks=0,
                    style={'height': '48px','text-align': 'center','float': 'right','font-size': '1.8rem', 'font-weight':'bold',
                        'position': 'relative'}
                ),
            ],
            className='row'
        ),

        html.Div(
            [
                html.P('Channel ID:',style={'margin-top': '20'}),  # noqa: E501
                dcc.Dropdown(
                    id='Channel_ID_history_search',
                    options=[{'label': 'ID' + str(i), 'value': i} for i in System_List['ID'].unique()],
                    multi=True,
                    placeholder="請選擇儀器ID",
                ),                
            ],
            style={'margin-top': '20'}
        ),

        html.Br(),
        html.Div(id='sensor_graphics_history_search'),
        dcc.Interval(
                id='interval_component_history_search',
                interval=86400*1000, # in milliseconds
                n_intervals=0
        ),
    ],
    className='ten columns offset-by-one'
)

def Channel_ID_Data_Download(n_clicks,Channel_ID_List,start_date,end_date):
    try:
        for i_Channel_ID in Channel_ID_List:
            print('查詢之Channel_ID：'+str(i_Channel_ID))
            
            Channel_ID_Download_List = System_List[System_List['ID'] == i_Channel_ID ]
            
            Channel_ID_Download_List = Channel_ID_Download_List.reset_index(drop=True)
            
            i_Channel_ID = Channel_ID_Download_List['TS_ID(FOR DB)'][0]

            return html.Div([
                html.A('Channel_ID：'+str(i_Channel_ID)+' 原始資料下載', id=str(i_Channel_ID)+'downloadbtn', 
                       href='/dash/urlToDownload?value={}'.format(str(i_Channel_ID)+'__'+start_date+'__'+end_date),
                       className="button button--primary",
                       style={
                            "height": "34",
                            "background": "##87CEFA",
                            "border": "1px solid #FFF",
                            "color": "white",
                            'float': 'right',}),       
                html.Br(),
                html.Br()])
    except Exception as e:
        print(e)
        return html.Div([
            '查詢失敗，可能無資料或系統錯誤。'
        ])


# updates time
@app.callback(
    Output("time_value_history_search", "children"),[Input('interval_component_history_search', 'n_intervals')]
)
def time_value_callback_history_search(n):
   
    return html.P(
        '現在時間：'+datetime.datetime.now().strftime('%Y-%m-%d') + ' ' + datetime.datetime.now().strftime('%H:%M:%S'),
        className="twelve columns indicator_text"
    )

@app.callback(Output('date_picker_history_search', 'initial_visible_month'),
              [Input('interval_component_history_search', 'n_intervals')])
def STARTDATE_RENEW_HISTORY_SEARCH(n):
      return date.today()
@app.callback(Output('date_picker_history_search', 'end_date'),
              [Input('interval_component_history_search', 'n_intervals')])
def ENDDATE_RENEW_HISTORY_SEARCH(n):
      return date.today()

@app.callback(
    dash.dependencies.Output('sensor_graphics_history_search', 'children'),
    [dash.dependencies.Input('button_history_search', 'n_clicks')],
    [dash.dependencies.State('Channel_ID_history_search', 'value'),
     dash.dependencies.State('date_picker_history_search', 'start_date'),
     dash.dependencies.State('date_picker_history_search', 'end_date'),

     ])
def Channel_ID_Data_Download_Callback(n_clicks,Channel_ID_input,start_date,end_date):
    if n_clicks>0:
        if Channel_ID_input == None:
                return html.Div([
                    '請選擇Channel_ID。'
                ])
        elif start_date == None:
                return html.Div([
                    '請選擇起始日期。'
                ])
        elif end_date == None:
                return html.Div([
                    '請選擇結束日期。'
                ])                            
        else:
            try:
                children = [Channel_ID_Data_Download(n_clicks,i,start_date,end_date) for i in zip(Channel_ID_input)]
                return children
            except Exception as e:
                print(e)
                return html.Div([
                    '下載失敗，可能無資料或系統錯誤。'
                    ])
@app.server.route('/dash/urlToDownload')
def DOWNLOAD_RAWDATA():
    if session.get('logged_in') and session.get('role')!='user': 
        value = flask.request.args.get('value')
        value = value.split('__')
        
        id_channel = str(value[0])

        DB_PATH = SYS_path+"/DATA/"      
        filename1 = Path(DB_PATH + "Database.mdb")
        
        try:
            conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(filename1))
            cnxn = pyodbc.connect(conn_str)
        except Exception as e:
            print(e)     
    
        try:
            tablename= str(id_channel)
            df_input = pd.read_sql("SELECT * FROM %s" %(tablename), cnxn)
        except (BaseException): 
            None    
                           
        df_input = df_input.set_index('時間') # 将date设置为index        
        df_input = df_input.sort_index(ascending=[True])
        
        value[2] = datetime.datetime.strptime(value[2], '%Y-%m-%d')
        value[2] = value[2] + datetime.timedelta(days = 1)
        value[2] = datetime.datetime.strftime(value[2], '%Y-%m-%d')
        
        df_input = df_input.truncate(before = value[1])
        df_input = df_input.truncate(after = value[2])      
        
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
        #taken from the original question
        df_input.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = value[0])
    
    
        #the writer has done its job
        writer.close()
    
        #go back to the beginning of the stream
        output.seek(0)
    
        #finally return the file
        return flask.send_file(output, attachment_filename=str(value[0])+'.xlsx', as_attachment=True)
    else:
        return render_template('index.html')    
  

#################################################################################################################
def modal_add_user_management():
    USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
    conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
    cnxn = pyodbc.connect(conn_str)
    df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
    return html.Div(
        html.Div(
            [
                html.Div(
                    [   

                        # modal header
                        html.Div(
                            [
                                html.Span(
                                    "設定使用者權限",
                                    style={
                                        "color": "#000",
                                        "fontWeight": "bold",
                                        "fontSize": "20",
                                    },
                                ),
                                html.Span(
                                    "×",
                                    id="add_close_user_management",
                                    n_clicks=0,
                                    style={
                                        "float": "right",
                                        "cursor": "pointer",
                                        "marginTop": "0",
                                        "marginBottom": "17",
                                    },
                                ),
                            ],
                            className="row",
                            style={"borderBottom": "1px solid #C8D4E3"},
                        ),

                        # modal form
                        html.Div(
                            [

                                html.P(
                                    "使用者帳號",
                                    style={
                                        "textAlign": "left",
                                        "marginBottom": "2",
                                        "marginTop": "4",
                                    },
                                ),
                                dcc.Dropdown(
                                    id="new_no_user_management_list",
                                    options=[{'label': i, 'value': i} for i in df['帳號']],
                                    value="",
                                ),

                                html.P(
                                    "使用者權限",
                                    style={
                                        "textAlign": "left",
                                        "marginBottom": "2",
                                        "marginTop": "4",
                                    },
                                ),
                                dcc.Dropdown(
                                    id="new_role_user_management",
                                    options=[
                                        {
                                            "label": "訪客",
                                            "value": "user",
                                        },
                                        {
                                            "label": "計畫成員",
                                            "value": "admin",
                                        },
                                        {
                                            "label": "系統管理者",
                                            "value": "super_admin",
                                        },
                                    ],

                                ),
                                html.P(
                                    "開通狀態",
                                    style={
                                        "textAlign": "left",
                                        "marginBottom": "2",
                                        "marginTop": "4",
                                    },
                                ),
                                dcc.Dropdown(
                                    id="new_status_user_management",
                                    options=[
                                        {
                                            "label": "未開通",
                                            "value": "未開通",
                                        },
                                        {
                                            "label": "開通",
                                            "value": "開通",
                                        },

                                    ],
 
                                ),
                                html.P(
                                    "資料通知",
                                    style={
                                        "textAlign": "left",
                                        "marginBottom": "2",
                                        "marginTop": "4",
                                    },
                                ),
                                dcc.Dropdown(
                                    id="new_notification_user_management",
                                    options=[
                                        {
                                            "label": "開啟",
                                            "value": "開啟",
                                        },
                                        {
                                            "label": "關閉",
                                            "value": "關閉",
                                        },

                                    ],
 
                                ),
                            ],
                            className="row",
                            style={"padding": "2% 8%"},
                        ),

                        # submit button
                        html.Span(
                            "確認",
                            id="submit_new_user_management",
                            n_clicks=0,
                            className="button button--primary add"
                        ),
                    ],
                    className="modal-content",
                    style={"textAlign": "center"},
                )
            ],
            className="modal",
        ),
        id="modal_add_user_management",
        style={"display": "none"},
    )

def modal_dele_user_management():
    USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
    conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
    cnxn = pyodbc.connect(conn_str)
    df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
    return html.Div(
        html.Div(
            [
                html.Div(
                    [   

                        # modal header
                        html.Div(
                            [
                                html.Span(
                                    "刪除使用者",
                                    style={
                                        "color": "#000",
                                        "fontWeight": "bold",
                                        "fontSize": "20",
                                    },
                                            
                                ),
                                html.Span(
                                    "×",
                                    id="dele_close_user_management",
                                    n_clicks=0,
                                    style={
                                        "float": "right",
                                        "cursor": "pointer",
                                        "marginTop": "0",
                                        "marginBottom": "17",
                                    },
                                ),
                            ],
                            className="row",
                            style={"borderBottom": "1px solid #C8D4E3"},
                        ),

                        # modal form
                        html.Div(
                            [
                                html.P(
                                    [
                                        "請選擇欲刪除之使用者",
                                        
                                    ],
                                    style={
                                        "float": "left",
                                        "marginTop": "4",
                                        "marginBottom": "2",
                                    },
                                    className="row",
                                ),
                                dcc.Dropdown(
                                    id="dele_no_user_management",
                                    options=[{'label': i, 'value': i} for i in df['帳號']],
                                    value="",
                                ),
                            ],
                            className="row",
                            style={"padding": "2% 8%"},
                        ),

                        # submit button
                        html.Span(
                            "確認",
                            id="submit_dele_user_management",
                            n_clicks=0,
                            className="button button--primary add"
                        ),
                    ],
                    className="modal-content",
                    style={"textAlign": "center"},
                )
            ],
            className="modal",
        ),
        id="modal_dele_user_management",
        style={"display": "none"},
    )

user_management_layout = [
    html.Br(),
        html.Div(
        [
        html.A(
            "登出",
            href = 'logout',
            className="button button--primary",
            style={
                "height": "24",
                "background": "##87CEFA",
                "border": "1px solid #FFF",
                "color": "white",
                'float': 'right',
            },
        ),
        html.Br(),
        html.Div(
                    id="time_value_user_management",
                    className="twelve columns indicator_text",
            style={
                "color": "black",
                'float': 'left',
            },
                ),
        html.Br(),   
    # top controls
    html.Br(),
    html.Br(),
    html.Br(),                

    modal_add_user_management(),
    modal_dele_user_management(),
html.Div([

        html.Div(
            [
                # add button
                html.Div(
                    html.Span(
                        "設定使用者權限",
                        id="add_new_button_user_management",
                        n_clicks=0,
                        className="button button--primary add",
    #                    style={
    #                        "height": "34",
    #                        "background": "#119DFF",
    #                        "border": "1px solid #119DFF",
    #                        "color": "white",
    #                    },
                    ),
                    style={"float": "right"},
                ),
            ],
            className="row",
            style={"marginBottom": "10"},
        ),
                html.Div(
                    id="table_user_management",
                    style={
                        "maxHeight": "350px",
                        "overflowY": "scroll",
                        "padding": "8",
                        "marginTop": "5",
                        "marginBottom": "10",
                        "backgroundColor":"white",
                        "border": "1px solid #C8D4E3",
                        "borderRadius": "3px"
                    },
                ),
                html.Div(
                    [
                        # add button
                        html.Div(
                            html.Span(
                                "刪除使用者帳號",
                                id="dele_button_user_management",
                                n_clicks=0,
                                className="button button--primary add",
            #                    style={
            #                        "height": "34",
            #                        "background": "#119DFF",
            #                        "border": "1px solid #119DFF",
            #                        "color": "white",
            #                    },
                            ),
                            style={"float": "right"},
                        ),
                    ],
                    className="row",
                    style={"marginBottom": "10"},
                ),

            ],
            className='row'
        ),

    html.Br(),
    dcc.Interval(
            id='interval_component_user_management',
            interval=86400*1000, # in milliseconds
            n_intervals=0
    ),
    ],
    className='ten columns offset-by-one'
)            
]

# return html Table with dataframe values  
def df_to_table(df):
    return html.Table(
        # Header
        [html.Tr([html.Th(col,style={'textAlign': 'center'}) for col in df.columns])] +
        
        # Body
        [
            html.Tr(
                [
                    html.Td(df.iloc[i][col],style={'textAlign': 'center'})
                    for col in df.columns
                ]
            )
            for i in range(len(df))
        ]

    )

# updates time
@app.callback(
    Output("time_value_user_management", "children"),[Input('interval_component_user_management', 'n_intervals')]
)
def time_value_callback_user_management(n):
   
    return html.P(
        '現在時間：'+datetime.datetime.now().strftime('%Y-%m-%d') + ' ' + datetime.datetime.now().strftime('%H:%M:%S'),
        className="twelve columns indicator_text"
    )


@app.callback(
    dash.dependencies.Output('table_user_management', 'children'),
    [dash.dependencies.Input('submit_new_user_management', 'n_clicks'),
     dash.dependencies.Input('submit_dele_user_management', 'n_clicks'),
     dash.dependencies.Input('add_new_button_user_management', 'n_clicks'),
     dash.dependencies.Input('dele_button_user_management', 'n_clicks'),
     dash.dependencies.Input("add_close_user_management", "n_clicks"),
     dash.dependencies.Input("dele_close_user_management", "n_clicks")
     ],
    [dash.dependencies.State('new_no_user_management_list', 'value'),
     dash.dependencies.State('dele_no_user_management', 'value'),
     dash.dependencies.State('new_role_user_management', 'value'),
     dash.dependencies.State('new_status_user_management', 'value'),     
     dash.dependencies.State('new_notification_user_management', 'value'),  
  
     ])
def USER_MANAGEMENT_EDIT(n_clicks1,n_clicks2,n_clicks3,n_clicks4,n_clicks5,n_clicks6,new_no,dele_no,new_role,new_status,new_notification):
  
    if n_clicks1 > 0:
        if n_clicks3 == 0 and n_clicks5 == 0 and new_no != None and new_status != None and new_role != None and new_notification != None:
            USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
            conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
            cnxn = pyodbc.connect(conn_str)
            tablename = 'USERLIST'
            cur = cnxn.cursor()
            try:
                createtable = """CREATE TABLE %s (帳號 Memo PRIMARY KEY,密碼 Memo,權限 Memo,開通狀態 Memo,資料通知 Memo )"""%(tablename)
                cur.execute(createtable)
                cnxn.commit()
            except Exception as e:
                print(e)
                USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
                conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
                cnxn = pyodbc.connect(conn_str)
                df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
                df = df.drop(columns=['密碼'])
            update_count1 = 0
            update_count2 = 0
            update_count3 = 0
            try:    
                with cnxn.cursor() as crsr:
                    sql = """\
                    UPDATE %s 
                    SET 開通狀態 = ?
                    WHERE 帳號 = ?
                    """%(tablename)
                    params = (new_status,new_no)
                    crsr.execute(sql, params)
                    crsr.commit()
                    update_count1 = update_count1 + crsr.rowcount 
                with cnxn.cursor() as crsr:
                    sql = """\
                    UPDATE %s 
                    SET 權限 = ?
                    WHERE 帳號 = ?
                    """%(tablename)
                    params = (new_role,new_no)
                    crsr.execute(sql, params)
                    crsr.commit()
                    update_count2 = update_count2 + crsr.rowcount 

                with cnxn.cursor() as crsr:
                    sql = """\
                    UPDATE %s 
                    SET 資料通知 = ?
                    WHERE 帳號 = ?
                    """%(tablename)
                    params = (new_notification,new_no)
                    crsr.execute(sql, params)
                    crsr.commit()
                    update_count3 = update_count3 + crsr.rowcount 
            
                fromaddr = 'ihmtmonitor@gmail.com'
                toaddr = new_no
                    
                if update_count1 != 0 and new_status == '開通':
                    msg = MIMEMultipart()
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "【帳號開通通知】「台20線50.7k處即時邊坡擋土監測平台」"
                    body3 = str(toaddr)+" 您好，\n\n"+"管理員已成功開通您的帳號"
                    msg.attach(MIMEText(body3, 'plain'))
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.ehlo()
                    server.starttls()
                    server.login('ihmtmonitor@gmail.com', 'asml1qaz@WSX')
                    text = msg.as_string()
                    server.sendmail(fromaddr, toaddr, text)
                    server.quit()
                if update_count2 != 0:
                    msg = MIMEMultipart()
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "【權限設定通知】「台20線50.7k處即時邊坡擋土監測平台」"
                    body3 = str(toaddr)+" 您好，\n\n"+"管理員已將您的帳號權限設定為"+str(new_role)+ "\n\n"
                    msg.attach(MIMEText(body3, 'plain'))
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.ehlo()
                    server.starttls()
                    server.login('ihmtmonitor@gmail.com', 'asml1qaz@WSX')
                    text = msg.as_string()
                    server.sendmail(fromaddr, toaddr, text)
                    server.quit()
                if update_count3 != 0 and new_notification == '開啟':
                    msg = MIMEMultipart()
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "【資料通知】「台20線50.7k處即時邊坡擋土監測平台」"
                    body3 = str(toaddr)+" 您好，\n\n"+"管理員已將您的帳號設定為資料通知對象，您後續可以收到資料同步狀態相關訊息" + "\n\n"
                    msg.attach(MIMEText(body3, 'plain'))
                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.ehlo()
                    server.starttls()
                    server.login('ihmtmonitor@gmail.com', 'asml1qaz@WSX')
                    text = msg.as_string()
                    server.sendmail(fromaddr, toaddr, text)
                    server.quit()
                
            except Exception as e:
                print(e)
                USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
                conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
                cnxn = pyodbc.connect(conn_str)
                df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
                df = df.drop(columns=['密碼'])
                return df_to_table(df) 
            df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
            df = df.drop(columns=['密碼'])
            return df_to_table(df)

      
        else:
            USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
            conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
            cnxn = pyodbc.connect(conn_str)
            df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
            df = df.drop(columns=['密碼'])
            return df_to_table(df)           
    elif n_clicks2 > 0:
        if n_clicks4 == 0 and n_clicks6 == 0 and dele_no != None:
            USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
            try:        
                conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
                cnxn = pyodbc.connect(conn_str)
                tablename = 'USERLIST'
                try:
                    with cnxn.cursor() as crsr:
                        sql = """\
                    DELETE FROM %s WHERE 帳號 = ?
                    """%(tablename)
                        params = dele_no
                        crsr.execute(sql, params)
                        cnxn.commit()
                    df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
                    df = df.drop(columns=['密碼'])
                    return df_to_table(df)
                except Exception as e:
                    print(e)
                    return df_to_table(df) 
            except Exception as e:
                print(e)
                USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
                conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
                cnxn = pyodbc.connect(conn_str)
                df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
                df = df.drop(columns=['密碼'])
                return df_to_table(df)       
        else:
            USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
            conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
            cnxn = pyodbc.connect(conn_str)
            df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
            df = df.drop(columns=['密碼'])
            return df_to_table(df)     
    else:                   
        USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
        cnxn = pyodbc.connect(conn_str)
        df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
        df = df.drop(columns=['密碼'])
        return df_to_table(df)     


# hide/show modal_add
@app.callback(Output("modal_add_user_management", "style"), [Input("add_new_button_user_management", "n_clicks")])
def display_modal_add_user_management_callback(n_add):
    if n_add > 0:
        return {"display": "block"}
    return {"display": "none"}

# reset to 0 add button n_clicks property 
@app.callback(
    Output("add_new_button_user_management", "n_clicks"),
    [Input("add_close_user_management", "n_clicks"), Input("submit_new_user_management", "n_clicks")],
)
def close_modal_add_callback_user_management(n1_add, n2_add):
    return 0

# reset to 0 add close_button n_clicks property 
@app.callback(
    Output("add_close_user_management", "n_clicks"),
    [Input("submit_new_user_management", "n_clicks")]
)
def close_button_add_callback_user_management(n_add_close):
    return 0

# hide/show modal_dele
@app.callback(Output("modal_dele_user_management", "style"), [Input("dele_button_user_management", "n_clicks")])
def display_modal_dele_user_management_callback(n_dele):
    if n_dele > 0:
        return {"display": "block"}
    return {"display": "none"}

# GET 警報單號清單 modal_dele
@app.callback(Output("dele_no_user_management", "options"), 
              [Input("add_new_button_user_management", "n_clicks"),
               Input("dele_button_user_management", "n_clicks")])
def updatelsit_modal_dele_user_management_callback(n_add,n_dele):
    if n_dele > 0:
        WARNLSIT_DB = Path(SYS_path + "/USERLIST.mdb")       
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(WARNLSIT_DB))
        cnxn = pyodbc.connect(conn_str)
        df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
        return [{'label': i, 'value': i} for i in df['帳號']]
    else:
        return {'label': '-', 'value': '-'}
    
# GET 警報單號清單 modal_dele
@app.callback(Output("new_no_user_management_list", "options"), 
              [Input("add_new_button_user_management", "n_clicks"),
               Input("dele_button_user_management", "n_clicks")])
def updatelsit_modal_new_user_management_callback(n_add,n_dele):
    if n_add > 0:
        WARNLSIT_DB = Path(SYS_path + "/USERLIST.mdb")       
        conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(WARNLSIT_DB))
        cnxn = pyodbc.connect(conn_str)
        df = pd.read_sql("SELECT * FROM USERLIST", cnxn)
        return [{'label': i, 'value': i} for i in df['帳號']]
    else:
        return {'label': '-', 'value': '-'}

# reset to 0 dele button n_clicks property 
@app.callback(
    Output("dele_button_user_management", "n_clicks"),
    [Input("dele_close_user_management", "n_clicks"), Input("submit_dele_user_management", "n_clicks")],
)
def close_modal_dele_callback_user_management(n1_dele, n2_dele):
    return 0

# reset to 0 dele close_button n_clicks property 
@app.callback(
    Output("dele_close_user_management", "n_clicks"),
    [Input("submit_dele_user_management", "n_clicks")]
)
def close_button_dele_callback_user_management(n_dele_close):
    return 0

@app.callback(Output('date_picker_user_management', 'initial_visible_month'),
              [Input('interval_component_user_management', 'n_intervals')])
def STARTDATE_RENEW_USER_MANAGEMENT(n):
      return date.today()
  
@app.callback(Output('date_picker_user_management', 'end_date'),
              [Input('interval_component_user_management', 'n_intervals')])
def ENDDATE_RENEW_USER_MANAGEMENT(n):
      return date.today()

#################################################################################################################

@server_flask.route('/')
def home():

    if not session.get('logged_in'):
        return render_template('index.html')
    else:
        return flask.redirect('/AL') 

@server_flask.route('/register')
def register():

    if not session.get('logged_in'):
        return render_template('register.html')
    else:
        return flask.redirect('/AL') 
    
import json
import requests

def is_human(captcha_response):
    """ Validating recaptcha response from google server
        Returns True captcha test passed for submitted form else returns False.
    """
    secret = "6LdxVaYUAAAAABUJGEJrqQwoxIqNG9oThAaYkGoz"
    payload = {'response':captcha_response, 'secret':secret}
    response = requests.post("https://www.google.com/recaptcha/api/siteverify", payload)
    response_text = json.loads(response.text)
    return response_text['success']

@server_flask.route('/login', methods=['POST'])
def do_admin_login():
#    captcha_response = request.form['g-recaptcha-response']
#    
#
#        
#    if is_human(captcha_response):
    # Process request here

    USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
    conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
    cnxn = pyodbc.connect(conn_str)
    
    df_USERLIST = pd.read_sql("SELECT * FROM USERLIST", cnxn)
    df_USERLIST = df_USERLIST.loc[df_USERLIST['帳號'] == request.form['username']]
    df_USERLIST = df_USERLIST.reset_index()
    if len(df_USERLIST['帳號']) == 0 :
        flash('您輸入之EMAIL帳號尚未註冊!')
        return home()
    else:
        LOGIN_PASSWORD = df_USERLIST['密碼'][0]
        LOGIN_ROLE = df_USERLIST['權限'][0]
        LOGIN_ACTIVE_STATUS = df_USERLIST['開通狀態'][0]
#        print(df_USERLIST['密碼'][0])
#        print(df_USERLIST['權限'][0])
#        print(df_USERLIST['開通狀態'][0])
        
        if check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'super_admin' and LOGIN_ACTIVE_STATUS == '開通':
            session['logged_in'] = True      
            session['role'] = 'super_admin'
        elif check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'admin' and LOGIN_ACTIVE_STATUS == '開通':
            session['logged_in'] = True      
            session['role'] = 'admin'
        elif check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'user' and LOGIN_ACTIVE_STATUS == '開通':
            session['logged_in'] = True      
            session['role'] = 'user'
        elif check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'super_admin' and LOGIN_ACTIVE_STATUS == '未開通':
            flash('您輸入之EMAIL帳號尚未開通，請聯絡管理員開通!')
        elif check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'admin' and LOGIN_ACTIVE_STATUS == '未開通':
            flash('您輸入之EMAIL帳號尚未開通，請聯絡管理員開通!') 
        elif check_password_hash(LOGIN_PASSWORD,request.form['password']) and LOGIN_ROLE == 'user' and LOGIN_ACTIVE_STATUS == '未開通':
            flash('您輸入之EMAIL帳號尚未開通，請聯絡管理員開通!')
        else:
            flash('您輸入之EMAIL帳號或密碼錯誤!')
    return home()

#    else:
#         # Log invalid attempts
#        status = "圖片驗證碼失敗!"
#
#        flash(status)
#        return home()

@server_flask.route('/register', methods=['POST'])
def do_admin_register():
#    captcha_response = request.form['g-recaptcha-response']
        
#    if is_human(captcha_response):
    # Process request here
    
    USERLIST_DB = Path(SYS_path + "/USERLIST.mdb")
    conn_str = (r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};''DBQ=%s;' %(USERLIST_DB))
    cnxn = pyodbc.connect(conn_str)
    tablename = 'USERLIST'
    cur = cnxn.cursor()
    try:
        createtable = """CREATE TABLE %s (帳號 Memo PRIMARY KEY,密碼 Memo,權限 Memo,開通狀態 Memo,資料通知 Memo )"""%(tablename)
        cur.execute(createtable)
        cnxn.commit()
    except Exception as e:
        print(e)
        
    with cnxn.cursor() as crsr:
        sql = """\
    INSERT INTO %s (帳號,密碼,權限,開通狀態,資料通知)
    SELECT ? as 帳號, ? AS 密碼, ? AS 權限, ? AS 開通狀態, ? AS 資料通知
    FROM (SELECT COUNT(*) AS n FROM %s) AS Dual
    WHERE NOT EXISTS (SELECT * FROM %s WHERE 帳號 = ?)
    """%(tablename,tablename,tablename)
        params = (request.form['email'],generate_password_hash(request.form['psw']),'user','未開通','關閉',request.form['email'])
        crsr.execute(sql, params)
        cnxn.commit()
        update_count = crsr.rowcount

    fromaddr = "ihmtmonitor@gmail.com"
    toaddr = request.form['email']
        
    if update_count != 0:
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = "【註冊成功通知】「 台20線50.7k處即時邊坡擋土監測平台」"
        body3 = str(toaddr)+" 您好，\n\n"+"您已成功註冊「台20線50.7k處即時邊坡擋土監測平台」"+ "\n\n" + "目前帳號為尚未開通狀態，請聯絡管理員開通及設定權限"+ "\n\n"
        msg.attach(MIMEText(body3, 'plain'))
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login('ihmtmonitor@gmail.com', 'asml1qaz@WSX')        
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()

        df = pd.read_sql("SELECT * FROM USERLIST", cnxn)        
        df_NOTIFICATION = df[df['權限'] == 'super_admin']        
        toaddr_admin = list(df_NOTIFICATION['帳號'])
        
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = ", ".join(toaddr_admin)
        msg['Subject'] = "【帳號註冊通知】「台20線50.7k處即時邊坡擋土監測平台」"
        body3 = " 您好，\n\n"+"帳號 "+str(toaddr)+" 已註冊「台20線50.7k處即時邊坡擋土監測平台」"+ "\n\n" + "目前帳號為尚未開通狀態，請協助開通及設定權限"+ "\n\n"
        msg.attach(MIMEText(body3, 'plain'))
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login('ihmtmonitor@gmail.com', 'asml1qaz@WSX')        
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr_admin , text)
        server.quit()

        flash('您輸入之EMAIL帳號已註冊成功!')
        return flask.redirect('/register')
    elif update_count == 0:
        flash('您輸入之EMAIL帳號已註冊，請直接登入!')
        return flask.redirect('/register')
    else:
        return flask.redirect('/register')

#    else:
#         # Log invalid attempts
#        status = "圖片驗證碼失敗!"
#
#        flash(status)
#        return flask.redirect('/register')

@server_flask.route('/css/style.css')
def css():
  return render_template('style.css') # the file variable is created
                               # when a the <file> is something.

@server_flask.route("/<path:path>")
def template_mount(path):
    TEMPLATE_DIRECTORY = Path(SYS_path + "/templates/")
    return send_from_directory(TEMPLATE_DIRECTORY, path, as_attachment=False)

@server_flask.route("/AL/logout")
def logout():
    session['logged_in'] = False
    return flask.redirect('/')
  
if __name__ == "__main__":   
        
    server_flask.secret_key = os.urandom(12)
    server_flask.run(
                #debug=False, host="127.0.0.1", port=80, threaded=True)
                debug=False, host="210.61.148.55", port=5005, threaded=True)


