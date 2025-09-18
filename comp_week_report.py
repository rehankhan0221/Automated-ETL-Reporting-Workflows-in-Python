# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 08:46:38 2020

@author: B15400
"""

import pandas as pd
import numpy as np
import xlwings
import datetime

df=pd.read_pickle(r'C:\Python\pk\sale')
dc=pd.read_csv(r'C:\Python\Prod\ksa_comp.csv')
dw=pd.read_csv(r'C:\Python\Prod\weeks adj.csv',parse_dates=['DATE'],dayfirst=True)
dm=pd.read_csv(r'C:\Python\Prod\type.csv')

df['BRANCH']=df['BRANCH'].replace('13124 - RT BANDER MALL','13185 - RT BANDER MALL - NEW')
df['BRANCH']=df['BRANCH'].replace('13067 - RT NAKHEEL PLAZA','13186 - RT NAKHEELPLAZA - NEW')

dbr=df[df['BRANCH']=='13185 - RT BANDER MALL - NEW'].groupby(['COUNTRY', 'REGION', 'AR', 'BRANCH', 'YR', 'MN', 'WK', 'DATE']).agg({'SALECY_AED':'sum', 'BUDCY_AED':'sum', 'GP_AED':'sum', 'QTYCY':'sum', 'CUSTCY':'sum', 'FFCY':'sum', 'GPLY':'sum','CUSTLY':'sum', 'FFLY':'sum', 'SALELY_AED':'sum', 'BUDLY_AED':'sum', 'SALECY':'sum', 'BUDCY':'sum', 'GPV':'sum','GP':'sum', 'SALELY':'sum', 'BUDLY':'sum', 'QTY_LY':'sum', 'BUDGP':'sum', 'SALEXCY':'sum', 'BUDXCY_AED':'sum','SALEXLY_AED':'sum', 'BUDXCY':'sum', 'SALEXLY':'sum','SALEXYLY_AED':'sum'}).reset_index()
df=df[df['BRANCH']!='13185 - RT BANDER MALL - NEW']
dnp=df[df['BRANCH']=='13186 - RT NAKHEELPLAZA - NEW'].groupby(['COUNTRY', 'REGION', 'AR', 'BRANCH', 'YR', 'MN', 'WK', 'DATE']).agg({'SALECY_AED':'sum', 'BUDCY_AED':'sum', 'GP_AED':'sum', 'QTYCY':'sum', 'CUSTCY':'sum', 'FFCY':'sum', 'GPLY':'sum','CUSTLY':'sum', 'FFLY':'sum', 'SALELY_AED':'sum', 'BUDLY_AED':'sum', 'SALECY':'sum', 'BUDCY':'sum', 'GPV':'sum','GP':'sum', 'SALELY':'sum', 'BUDLY':'sum', 'QTY_LY':'sum', 'BUDGP':'sum', 'SALEXCY':'sum', 'BUDXCY_AED':'sum','SALEXLY_AED':'sum', 'BUDXCY':'sum', 'SALEXLY':'sum','SALEXYLY_AED':'sum'}).reset_index()
df=df[df['BRANCH']!='13186 - RT NAKHEELPLAZA - NEW']
df=df.append(dbr)
df=df.append(dnp)
del dbr
del dnp

st_ty=datetime.datetime(2022,7,2)#check for dates
st_ly=datetime.datetime(2021,7,3)
td=pd.to_datetime('today')
td=td.normalize()

ed_ty=td
ed_ly=td-pd.to_timedelta(364, unit='d')

dc.columns=['BRANCH','COMP']
dc['COMP'].loc[dc['COMP']=='COMPARABLE']='C'
dc['COMP'].loc[dc['COMP']=='NON COMPARABLE']='NC'
df=pd.merge(df,dc,on='BRANCH',how='left')
df=pd.merge(df,dw,on='DATE',how='left')


df=df[df['BRANCH']!='13175 - RT-KSA Online']
df=df[df['BRANCH']!='13192 - RT-KSA ONLINE']

dm=pd.read_csv(r'C:\Python\Prod\type.csv')
df=pd.merge(df,dm,on='BRANCH',how='left')
df['BUDGPCY']=df['BUDXCY']*df['BUDGP']/100

dft=df[(df['DATE']>st_ty)&(df['DATE']<ed_ty)]
dfl=df[(df['DATE']>st_ly)&(df['DATE']<ed_ly)]


td=pd.to_datetime('today')
ly=dfl.groupby(['BRANCH','AR','TYPE','COMP','WEEK']).agg({'SALECY':'sum','SALEXCY':'sum',
        'BUDCY':'sum','BUDXCY':'sum',
        'QTYCY':'sum','CUSTCY':'sum',
        'GPV':'sum',
        'FFCY':'sum'}).reset_index()
ly.columns=['BRANCH', 'AR','TYPE','COMP','WEEK','SALELY','SALEXLY', 'BUDLY','BUDXLY', 'QTYLY', 'CUSTLY', 'GPVLY',
       'FFLY']

ty=dft.groupby(['BRANCH','AR','TYPE','COMP','WEEK']).agg({'SALECY':'sum','SALEXCY':'sum',
        'BUDCY':'sum','BUDXCY':'sum',
        'QTYCY':'sum','CUSTCY':'sum',
        'GPV':'sum',
        'BUDGPCY':'sum','FFCY':'sum'}).reset_index()

fin=pd.merge(ty,ly,on=['AR','BRANCH','COMP','TYPE','WEEK'],how='left')
fin=fin.set_index('BRANCH')
#fin['TRADE']=np.where((fin['SALECY']*fin['SALELY']>0.0),"T","NT")


wb = xlwings.Book("C:\Python\\Prod\\Templates\\WK_KPI_store_tpl.xlsx")
Sheet1 = wb.sheets[2]
Sheet1.range(1,1).options(index=True).value = fin
wb.save("C:\Python\KPIs\\\WK_KPI_store.xlsx")





