# -*- coding: utf-8 -*-
"""
Created on Sun Feb 11 22:56:17 2024

@author: antoi
"""

"""

This project aims to replicate a weekly economic calendar by scraping data from 
internet, and manipulates excel sheet directly through python to get an usable
calendar. 
Note : "Détails" column in the excel output is kept empty to let the user
add commentary regarding the data.

"""

import pandas as pd
import requests
import xlsxwriter
import os
from datetime import datetime, timedelta
from datetime import datetime

date_pm = input("Enter a date (YYYY-MM-DD) :")

url = 'https://economic-calendar.tradingview.com/events'

# Date format management : 
today = pd.Timestamp(f'{date_pm} 00:00:00').normalize()
start = (today + pd.offsets.Hour(23)).isoformat() + '.000Z'
end = (today + pd.offsets.Day(6) + pd.offsets.Hour(22)).isoformat() + '.000Z'

payload = {
    'from': start,
    'to': end,
    'countries': ','.join(['US','FR',"DE",'EU','AU','CA','UK','JP'])
}
data = requests.get(url, params=payload).json()
df = pd.DataFrame(data['result'])

# We set our degree of news importance to always keep a minimum of news.
if len(df[df["importance"]>=1])<13 : 
    importance=0
    importance_text = "Importance : two and three stars"
else : 
    importance=1
    importance_text = "Importance : only three stars" 

calendar = df[df["importance"]>=importance]


# Cleaning data : 
## Deleting useless tickers 
list_ticker_useless = ['USGSCH','USCSC','USCOSC','CAPTE','CAFTE','USAHEYY',
                       'USAHE','USLFPR','AUSPMI','USEC','CALFPR','DEUC','DEUP',
                       'USFO','DEUC','DEUR', 'DEUP ', 'USMEMP', 'USCSC', 
                       'USEC', 'USCOSC', 'USGSCH', 'USAHEYY', 'USAHE', 
                       'USLFPR', 'USFO', 'DEEXP', 'DEFO', 'FRBOT', 'USEXP', 
                       'USIMP', 'USBOT', 'USEOI', 'USCPI', 'USGBV', 'DEWPIMM', 
                       'DEWPIYY', 'USBI', 'USHMI', 'USEMCI', 'USEMCIW', 
                       'USEMCIB','USCSHPIMM', 'USCSHPIYY', 'USPHSIMM', 
                       'USPHSIYY', 'USTICNLF','USEHS', 'USEHSMM']
 
for ticker in list_ticker_useless : 
    calendar = calendar.drop(calendar[calendar["ticker"]==ticker].index)

## Deleting two stars news from "secondary" country (the ones we are less
## interested in) : 
secondary_country = ['AU','CA','UK','JP']
for i in secondary_country  :
    print(calendar[(calendar["country"]==i) & \
    (calendar["importance"]==0)].index)
    calendar = calendar.drop(calendar[(calendar["country"]==i) & \
    (calendar["importance"]==0)].index)


calendar = calendar[["title","country","period","actual","previous","forecast",
                    "date","unit","scale"]]

## Managing date and hour to get usable format
def get_new_date(row) : 
    return str(row[8:10]+"/"+row[5:7]+"/" +row[0:4])
calendar["newdate"] = calendar["date"].apply(get_new_date)

def get_new_hour(row) : 
    return str(row[11:16])
calendar["newhour"] = calendar["date"].apply(get_new_hour)


def transfo_str(row):
    return str(row)
calendar["newdate"] = calendar["newdate"].apply(transfo_str)
calendar["newhour"] = calendar["newhour"].apply(transfo_str)

## Applying function and deleting NaN (from speech for example)
final_cal = calendar
final_cal['scale'] = final_cal['scale'].fillna("")
final_cal['unit'] = final_cal['unit'].fillna("")
final_cal["forecast"] = final_cal["forecast"].fillna("")

final_cal["actual"] = final_cal["actual"].apply(transfo_str)
final_cal["previous"] = final_cal["previous"].apply(transfo_str)
final_cal["forecast"] = final_cal["forecast"].apply(transfo_str)

final_cal["title"] = calendar["title"] + " ("+calendar["period"]+")"
final_cal["actual"] = final_cal["actual"]+final_cal["scale"]+final_cal["unit"]
final_cal["previous"] = final_cal["previous"]+final_cal["scale"]+final_cal["unit"]
final_cal["forecast"] = final_cal["forecast"]+final_cal["scale"]+final_cal["unit"]

final_cal = final_cal[["newdate","newhour","country","title","actual","previous",
                       "forecast"]]

date_list = final_cal["newdate"].unique()

day_position = [1,]

# Creating list of weekly day for more readable output : 
jours_semaine = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 
'Dimanche']
mois = ['', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 
'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']

# Starting mapping the dataframe to get the expected result : 
mapping = pd.DataFrame(columns =["newdate","newhour","country","title","actual",
                                 "previous","forecast","details"] )
for ddate in date_list :
    yyyy = int(ddate[6:11])
    mm = int(ddate[3:5])
    dd = int(ddate[0:2])
    ddate_dt = datetime(yyyy, mm, dd)
    jour_i = str(jours_semaine[ddate_dt.weekday()])
    date_jour_i = str(dd)
    mois_jour_i = str(mois[ddate_dt.month])
    date_long = jour_i +" "+ date_jour_i +" "+ mois_jour_i
    d = {'newdate': date_long, 'newhour':"",'country':"", "title":"",
         "actual":"Actuel","previous":"Précédent","forecast":"Consensus",
         "details":"Détails" }
    temporaire = pd.DataFrame(data=d,index=[0])
    temporaire_2 = final_cal[final_cal["newdate"]==ddate]
    temporaire_2["newdate"]=""
    temporaire_2 = pd.concat([temporaire,temporaire_2],axis=0)
    mapping = pd.concat([mapping,temporaire_2],axis=0)
    long = len(mapping)
    day_position.append(long+1)
day_position.remove(day_position[-1])  
long_table = len(mapping)
mapping = mapping.rename(columns={"newdate":"Date","newhour" : ""
                                  , "country":" ","title": ""
                                  ,"details":"Détails"})

# Starting manipulate the excel file : 
output_path = "cal_eco.xlsx"

# Visual modification : 
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    mapping.to_excel(writer, index=False, sheet_name='report')
    worksheet = writer.sheets['report']
    workbook = writer.book
    
    # Defining different cell format
    cell_format = workbook.add_format({
        'bold': False,
        'font_color': 'black',
        'bg_color': '#FFFFFF',
        'border': 0,
        'align': 'vcenter'

    })
    cell_format_2 = workbook.add_format({
        'bold': False,
        'font_color': 'black',
        'bg_color': '#FFFFFF',
        'border': 0,
        'align': 'vcenter',
        'border': 1,
        'bottom': 1,
        'top' : 0,
        'right' : 0,
        'left' : 0,
        'border_color' : "E7E7E7"

    })
    detail_format = workbook.add_format({
        'bold': False,
        'font_color': 'black',
        'bg_color': '#FFFFFF',
        'border': 0,
        'align': 'vcenter'
    })
    
    day_row_format = workbook.add_format({
        'bold': False,
        'font_color': 'black',
        'bg_color': '#EFEFEF',
        'border': 1,
        'bottom': 1,
        'top' : 0,
        'right' : 0,
        'left' : 0,
        'align': 'vcenter'
    })
    
    # Replacing ISO code of the news by the flag concerned
    for row in range(0,long_table) : 
        country_id=mapping.iloc[row,2]
        try : 
            row_excel = row+2
            worksheet.insert_image(f"C{row_excel}", f"{country_id}_FLAG.png",
                                   {"x_scale": 0.18, "y_scale": 0.18,
                                     "x_offset":1, "y_offset": 3.5})
        except :    
            pass
    
    # Applying cell format : 
    worksheet.set_column('A:H', None, cell_format)
    worksheet.set_column('F:G', 14,cell_format)
    worksheet.set_column('D:D', 40, cell_format)
    worksheet.set_column('H:H', 60, cell_format)

    for j in range(1,long_table+1): 
        worksheet.set_row(j,21,cell_format_2)
    
    for i in day_position :
        worksheet.set_row(i,24,day_row_format)
workbook.close()

