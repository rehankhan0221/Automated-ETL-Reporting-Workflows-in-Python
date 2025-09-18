# -*- coding: utf-8 -*-
"""
Created on Thu Mar 28 10:22:59 2019

@author: B15400
"""

import datetime
import os
import pandas as pd
import calendar
import math
import numpy as np


os.chdir(r'C:\Python')

dff=pd.read_pickle(r'C:\Python\pk\sale_jan_all')
dfk=pd.read_pickle(r'C:\Python\pk\sale_jan')
tr=pd.read_csv(r'C:\Python\Prod\pr_br.csv')
df=pd.read_excel(r'C:\Python\Prod\daily_sales_exv.xlsx',skiprows=1)
df=df.drop(['COMPANY'],axis=1)
df.columns=['COUNTRY', 'REGION', 'AR', 'BRANCH', 'YR', 'MN', 'WK', 'DATE',
       'SALECY_AED', 'BUDCY_AED', 'GP_AED', 'QTYCY', 'CUSTCY', 'FFCY', 'GPLY',
       'CUSTLY', 'FFLY', 'SALELY_AED', 'BUDLY_AED', 'SALECY', 'BUDCY', 'GPV',
       'GP', 'SALELY', 'BUDLY','QTY_LY','BUDGP','SALEXCY','BUDXCY_AED','SALEXLY_AED','BUDXCY','SALEXLY']
df['DATE']=pd.to_datetime(df['DATE'])

td=pd.to_datetime('today')-pd.to_timedelta(1, unit='d')
df.drop(df[df['DATE']>td].index, inplace=True)


df['YR']=df['YR'].astype(int)
df['SALECY']=df['SALECY'].astype(float)
df['SALECY_AED']=df['SALECY_AED'].astype(float)
df['BUDCY_AED']=df['BUDCY_AED'].astype(float)
df['BUDCY']=df['BUDCY'].astype(float)
df['GP']=df['GP'].astype(float)
df['GPV']=df['GPV'].astype(float)
df['GP_AED']=df['GP_AED'].astype(float)
df['QTYCY']=df['QTYCY'].astype(int)
df['CUSTCY']=df['CUSTCY'].astype(int)
df['GPLY']=df['GPLY'].astype(float)
df['FFCY']=df['FFCY'].astype(int)
df['CUSTLY']=df['CUSTLY'].astype(int)
df['FFLY']=df['FFLY'].astype(int)
df['SALELY']=df['SALELY'].astype(float)
df['SALELY_AED']=df['SALELY_AED'].astype(float)
df['BUDLY']=df['BUDLY'].astype(float)
df['SALEXCY']=df['SALEXCY'].astype(float)
df['BUDXCY_AED']=df['BUDXCY_AED'].astype(float)
df['SALEXYLY_AED']=df['SALEXLY_AED'].astype(float)
df['BUDXCY']=df['BUDXCY'].astype(float)
df['SALEXLY']=df['SALEXLY'].astype(float)

dfa=dff.append(df)
dfa['AR']=np.where(dfa['BRANCH']=='13124 - RT BANDER MALL','SR3',dfa['AR'])
dfa.to_pickle(r'C:\Python\pk\sale_all')

df=df[(df['REGION']=='ERO')|(df['REGION']=='CRO')|(df['REGION']=='CRN')|(df['REGION']=='SRO')|(df['REGION']=='WRO')|(df['REGION']=='ESA')]

#df.drop(df[df['DATE']<'2019-9-29'].index, inplace=True)
df=dfk.append(df)
df['AR']=np.where(df['BRANCH']=='13124 - RT BANDER MALL','SR3',df['AR'])
df['AR']=np.where(df['REGION']=='ESA','ES1',df['AR'])
df['BRANCH']=np.where(df['REGION']=='ESA','13800 - RT KSA - Ecom',df['BRANCH'])
pd.options.display.float_format = '{:.2f}'.format


yd=td-pd.to_timedelta(1, unit='d')
yd
yd=td.normalize()
print(df[(df['DATE']==yd)&(df['SALECY']==0)&(df['BRANCH'].isin(tr['BRANCH']))]['BRANCH'])
dfill=(df[(df['DATE']==yd)&(df['SALECY']==0)&(df['BRANCH'].isin(tr['BRANCH']))]['BRANCH']).reset_index()
dfill
for b in dfill['BRANCH']:
    print(b)
    tx="Sale  "+b
    mis= int(input("Sale"))
    mis=mis*1.00
    df.loc[(df['BRANCH'] == b)&(df['DATE']==yd), 'SALECY']=mis

df.to_pickle(r'C:\Python\pk\sale')    
sl=df[df['DATE']==yd]['SALECY'].sum()
sl_aed=(sl*.975/1000000).round(2)
t=pd.to_timedelta(364,unit='d')
ly=((((df[df['DATE']==yd]['SALECY'].sum())/df[df['DATE']==(yd-t)]['SALECY'].sum())-1)*100).round(2)
lw=((((df[df['DATE']==yd]['SALECY'].sum())/df[df['DATE']==(yd-pd.to_timedelta(7,unit='d')).normalize()]['SALECY'].sum())-1)*100).round(2)
bud=((((df[df['DATE']==yd]['SALECY'].sum())/df[df['DATE']==yd]['BUDCY'].sum())-1)*100).round(2)
day=calendar.day_name[yd.weekday()]
wk=math.ceil((td-datetime.datetime(2022,7,2)).days/7)
print("WK ",wk,day,td.strftime("%B %d, %Y"),"Sales",sl_aed,"m LY",ly,"% LW",lw,"% Bud",bud,"%")

slw_aed=(df[(df['DATE']<=yd)&(df['DATE']>(yd-pd.to_timedelta(7,unit='d')))]['SALECY'].sum()*.975/1000000).round(2)
ly_w=(((df[(df['DATE']<=yd)&(df['DATE']>(yd-pd.to_timedelta(7,unit='d')))]['SALECY'].sum()/df[(df['DATE']<=(yd-t))&(df['DATE']>(yd-pd.to_timedelta(371,unit='d')))]['SALECY'].sum())-1)*100).round(2)
lw_w=(((df[(df['DATE']<=yd)&(df['DATE']>(yd-pd.to_timedelta(7,unit='d')))]['SALECY'].sum()/df[(df['DATE']<=(yd-pd.to_timedelta(7,unit='d')))&(df['DATE']>(yd-pd.to_timedelta(14,unit='d')))]['SALECY'].sum())-1)*100).round(2)
bud_w=(((df[(df['DATE']<=yd)&(df['DATE']>(yd-pd.to_timedelta(7,unit='d')))]['SALECY'].sum()/df[(df['DATE']<=yd)&(df['DATE']>(yd-pd.to_timedelta(7,unit='d')))]['BUDCY'].sum())-1)*100).round(2)
print("WK",wk,"Sales",slw_aed,"M LY",ly_w,"% LW",lw_w,"% Bud",bud_w,"%")
