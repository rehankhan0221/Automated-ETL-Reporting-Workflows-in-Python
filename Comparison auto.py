# -*- coding: utf-8 -*-
"""
Created on Tue Aug 30 12:54:31 2022

@author: rehan
"""
    
import pandas as pd
import numpy as np
import xlwings as xw


df=pd.read_pickle(r'C:\Python\pk\sale')
t4=pd.read_pickle(r'C:\Python\pk\t4')
dty=df[(df['DATE']>='2022-11-23')&(df['DATE']<='2022-12-03')]#rt ty dates
dly=df[(df['DATE']>='2021-11-14')&(df['DATE']<='2021-11-27')]#rt ly dates
dt4y=df[(df['DATE']>='2022-11-20')&(df['DATE']<='2022-12-10')]#t4 ty dates
dl4y=df[(df['DATE']>='2021-10-17')&(df['DATE']<='2021-10-30')]#t4 ly dates

dty=dty.groupby(['COUNTRY','DATE']).agg({'SALECY':'sum'})
dty['days']=range(1,1+len(dty))
dty=pd.merge(dty, df, on='DATE')
dty['BUDGP%']=dty['BUDGP']/100
dty['BUDGPV']=dty['BUDGP%']*dty['BUDXCY']
dty=dty[['BRANCH','AR','DATE','SALECY_y','SALEXCY','BUDCY','QTYCY','CUSTCY','GPV','FFCY','BUDXCY','BUDGP','BUDGPV','days']]
dty['BRANCHDTTY']=dty['BRANCH'] + dty['days'].astype('str')

dly=dly.groupby(['COUNTRY','DATE']).agg({'SALECY':'sum'})
dly['days']=range(1,1+len(dly))
dly=pd.merge(dly, df, on='DATE')
dly['BUDGP%']=dly['BUDGP']/100
dly['BUDGPV']=dly['BUDGP%']*dly['BUDXCY']
dly=dly[['BRANCH','DATE','SALECY_y','SALEXCY','BUDCY','QTYCY','CUSTCY','GPV','FFCY','BUDXCY','BUDGP','BUDGPV','days']]
dly['BRANCHDTTY']=dly['BRANCH'] + dly['days'].astype('str')

dff=dty.merge(dly,on=['BRANCHDTTY'],how= 'left')
dc=pd.read_excel(r'C:\Python\Prod\Comptype2022.xlsx')
dc.drop('SUB REGION',axis=1,inplace=True)
dc.rename(columns = {'BRANCH NAME':'BRANCH_x'}, inplace = True)
fdf=pd.merge(dff,dc,on=['BRANCH_x'],how='left')
fdf=fdf[['BRANCH_x','AR','DATE_x','SALECY_y_x', 'SALEXCY_x', 'BUDCY_x',
       'QTYCY_x', 'CUSTCY_x', 'GPV_x', 'FFCY_x', 'BUDXCY_x', 'BUDGP_x',
       'BUDGPV_x', 'days_x','DATE_y','SALECY_y_y',
       'SALEXCY_y', 'BUDCY_y', 'QTYCY_y', 'CUSTCY_y', 'GPV_y', 'FFCY_y',
       'BUDXCY_y', 'BUDGP_y', 'BUDGPV_y','Trade','TYPE','C/NC','Crazy']]
fdf.columns=['BRANCH NAME','SUB REGION','DT','TY LCL NET SALES','TY EX VAT LCL NET SALES','TY LCL BUD SALES','TY NET QTY','TY CUSTOMERS','TY LCL GP','TY FOOTFALL','TY LCL Ex BUD SALES','38 TY BUD GP%','TY BUD GP V','Day','DT LY','LY LCL NET SALES','LY EX VAT LCL NET SALES','LY LCL BUD SALES','LY NET QTY','LY CUSTOMERS','LY LCL GP','LY FOOTFALL','LY LCL Ex BUD SALES','38 LY BUD GP%','LY BUD GP V','Trade','TYPE','C/NC','Crazy Deal']
fdf.sort_values(['BRANCH NAME','DT'],inplace=True)
#t4
dt4y=dt4y.groupby(['COUNTRY','DATE']).agg({'SALECY':'sum'})
dt4y['days']=range(1,1+len(dt4y))
dt4y=pd.merge(dt4y, t4, on='DATE')
dt4y['BUDGP%']=dt4y['BUDGP']/100
dt4y['BUDGPV']=dt4y['BUDGP%']*dt4y['BUDXCY']
dt4y=dt4y[['BRANCH','AR','DATE','SALECY_y','SALEXCY','BUDCY','QTYCY','CUSTCY','GPV','FFCY','BUDXCY','BUDGP','BUDGPV','days']]
dt4y['BRANCHDTTY']=dt4y['BRANCH'] + dt4y['days'].astype('str')

dl4y=dl4y.groupby(['COUNTRY','DATE']).agg({'SALECY':'sum'})
dl4y['days']=range(1,1+len(dl4y))
dl4y=pd.merge(dl4y, t4, on='DATE')
dl4y['BUDGP%']=dl4y['BUDGP']/100
dl4y['BUDGPV']=dl4y['BUDGP%']*dl4y['BUDXCY']
dl4y=dl4y[['BRANCH','DATE','SALECY_y','SALEXCY','BUDCY','QTYCY','CUSTCY','GPV','FFCY','BUDXCY','BUDGP','BUDGPV','days']]
dl4y['BRANCHDTTY']=dl4y['BRANCH'] + dl4y['days'].astype('str')

df4=dt4y.merge(dl4y,on=['BRANCHDTTY'],how= 'left')
d4c=pd.read_excel(r'C:\Python\Prod\Comptype2022T4.xlsx')
d4c.rename(columns = {'BRANCH NAME':'BRANCH_x'}, inplace = True)
dft4=pd.merge(df4,d4c,on=['BRANCH_x'],how='left')
dft4=dft4[['BRANCH_x','AR','DATE_x','SALECY_y_x', 'SALEXCY_x', 'BUDCY_x',
       'QTYCY_x', 'CUSTCY_x', 'GPV_x', 'FFCY_x', 'BUDXCY_x', 'BUDGP_x',
       'BUDGPV_x', 'days_x','DATE_y','SALECY_y_y',
       'SALEXCY_y', 'BUDCY_y', 'QTYCY_y', 'CUSTCY_y', 'GPV_y', 'FFCY_y',
       'BUDXCY_y', 'BUDGP_y', 'BUDGPV_y','Trade','C/NC']]
dft4.columns=['BRANCH NAME','SUB REGION','DT','TY LCL NET SALES','TY EX VAT LCL NET SALES','TY LCL BUD SALES','TY NET QTY','TY CUSTOMERS','TY LCL GP','TY FOOTFALL','TY LCL Ex BUD SALES','38 TY BUD GP%','TY BUD GP V','Day','DT LY','LY LCL NET SALES','LY EX VAT LCL NET SALES','LY LCL BUD SALES','LY NET QTY','LY CUSTOMERS','LY LCL GP','LY FOOTFALL','LY LCL Ex BUD SALES','38 LY BUD GP%','LY BUD GP V','Trade','C/NC']
dft4.sort_values(['BRANCH NAME','DT'],inplace=True)
datesrt=fdf.groupby(['DT','DT LY']).agg({'TY LCL NET SALES':'sum'})
datest4=dft4.groupby(['DT','DT LY']).agg({'TY LCL NET SALES':'sum'})

wb = xw.Book(r"C:\Python\Prod\Comparison Template.xlsx")
sheetCompt4 = wb.sheets['Data T4']
sheetCompt4.range(5,29).value= dft4.set_index('BRANCH NAME')

sheetComp = wb.sheets['Data']
sheetComp.range(5,29).value= fdf.set_index('BRANCH NAME')

dtrt=wb.sheets['print']
dtrt.range(2,34).value=datesrt

dtrt=wb.sheets['print']
dtrt.range(66,34).value=datest4


#dff=pd.merge(dty,dly,  how='cross', left_on=['BRANCH','DATE'], right_on = ['BRANCH','DATE'],right_index=True,left_index=True)
#dff= dty.merge(dly, on=['BRANCH','DATE'])
