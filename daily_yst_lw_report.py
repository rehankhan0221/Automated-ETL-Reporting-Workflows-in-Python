# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 23:13:14 2020

@author: B15400
"""

import pandas as pd
import numpy as np
import datetime
import xlwings
import wtd_mtd_new_asp_gp as wd


df=pd.read_pickle(r'C:\Python\pk\sale')
t4=pd.read_pickle(r'C:\Python\pk\t4')
t4['SEL']=np.where(((t4['BRANCH']=='23072 - T4 SHIFA')&(t4['DATE']>'2021-03-06')&(t4['YR']==2020)),1,0)
t4=t4[t4['SEL']==0]
t4=t4.drop('SEL',axis=1)

dm=pd.read_csv(r'C:\Python\Prod\t60_ksa.csv')
dm.columns=['BRANCH','TYPE']
dmt=pd.read_csv(r'C:\Python\Prod\type_t4.csv')

dc=pd.read_csv(r'C:\Python\Prod\ksa_comp.csv')
dct=pd.read_excel(r'C:\Python\Prod\syscomp_t4.xlsx',skiprows=1)
dsm=pd.read_csv(r'C:\Python\Prod\sm_contact.csv')

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

dc.columns=['BRANCH','COMP']
dc['COMP'].loc[dc['COMP']=='COMPARABLE']='C'
dc['COMP'].loc[dc['COMP']=='NON COMPARABLE']='NC'

dct.columns=['BRANCH','COMP']
dct['COMP'].loc[dct['COMP']=='COMPARABLE']='C'
dct['COMP'].loc[dct['COMP']=='NON COMPARABLE']='NC'

df=pd.merge(df,dc,on='BRANCH',how='left')
t4=pd.merge(t4,dct,on='BRANCH',how='left')

df=pd.merge(df,dm,on='BRANCH',how='left')
t4=pd.merge(t4,dmt,on='BRANCH',how='left')

tr=pd.read_csv(r'C:\Python\Prod\pr_br.csv')
df=df[df['BRANCH'].isin(tr['BRANCH'])]

td=pd.to_datetime('today')-pd.to_timedelta(1, unit='d')
td=td.normalize()

yd=td-pd.to_timedelta(1, unit='d')
yd=yd.normalize()

lw=td-pd.to_timedelta(7, unit='d')
lw=lw.normalize()

ly=td-pd.to_timedelta(364, unit='d')
ly=ly.normalize()

lm=td-pd.to_timedelta(28, unit='d')
lm=lm.normalize()

stw=df[df['WK']==df[(df['DATE']==td)]['WK'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
stm=df[df['MN']==df[(df['DATE']==td)]['MN'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
sty=df[df['YR']==df[(df['DATE']==td)]['YR'].unique()[0]]['DATE'].nsmallest(1).iloc[0]


def daysale(br):
    if br=='T4':
        dtemp=t4
    else:
        dtemp=df
    df1=dtemp[dtemp['DATE']==td].groupby(['REGION','BRANCH','TYPE','COMP']).agg({'SALECY':'sum','BUDCY':'sum'}).reset_index()
    df1.columns=['REGION','BRANCH','TYPE','LFL','SALE','BUD']
    df2=dtemp[dtemp['DATE']==yd].groupby(['BRANCH'])['SALECY'].sum().reset_index()
    df2.columns=['BRANCH','YST']
    df3=dtemp[dtemp['DATE']==lw].groupby(['BRANCH'])['SALECY'].sum().reset_index()
    df3.columns=['BRANCH','LW']
    df4=dtemp[dtemp['DATE']==ly].groupby(['BRANCH'])['SALECY'].sum().reset_index()
    df4.columns=['BRANCH','LY']
    df1=pd.merge(df1,df2,on='BRANCH',how='left')
    df1=pd.merge(df1,df3,on='BRANCH',how='left')
    df1=pd.merge(df1,df4,on='BRANCH',how='left')
    dt=[td,yd,lw,ly]
    for d in dt:
        dft=dtemp[dtemp['DATE']==d].groupby(['BRANCH']).agg({'QTYCY':'sum','FFCY':'sum','CUSTCY':'sum'}).reset_index()
        df1=pd.merge(df1,dft,on='BRANCH',how='left')
    df1.columns=['REGION', 'BRANCH', 'TYPE','LFL', 'SALE', 'BUD', 'YST', 'LW','LY' ,'QTYCY_td',
       'FFCY', 'CUSTCY', 'QTY_yst', 'FF_yst', 'CUST_yst', 'QTY_lw', 'FF_lw',
       'CUST_lw','QTY_ly','FF_ly','CUST_ly']
    dfw=dtemp[(dtemp['DATE']>=stw)&(dtemp['DATE']<=td)].groupby('BRANCH').agg({'SALECY':'sum','SALELY':'sum','QTYCY':'sum','QTY_LY':'sum'})
    dfm=dtemp[(dtemp['DATE']>=stm)&(dtemp['DATE']<=td)].groupby('BRANCH').agg({'SALECY':'sum','SALELY':'sum','QTYCY':'sum','QTY_LY':'sum'})
    dfy=dtemp[(dtemp['DATE']>=sty)&(dtemp['DATE']<=td)].groupby('BRANCH').agg({'SALECY':'sum','SALELY':'sum','QTYCY':'sum','QTY_LY':'sum'})
    df1=pd.merge(df1,dfw,on='BRANCH',how='left')
    df1=pd.merge(df1,dfm,on='BRANCH',how='left')
    df1=pd.merge(df1,dfy,on='BRANCH',how='left')
    return(df1)

df1=daysale('RT')
dft4=daysale('T4')


# For region wise bottom 5 and KPI

bt=(df[(df['DATE']==td)&(df['SALECY']>0)].groupby(['REGION','BRANCH'])['SALECY'].sum()/df[df['DATE']==lw].groupby(['REGION','BRANCH'])['SALECY'].sum())-1
d5=bt.groupby(level=0,group_keys=False).apply(lambda x: x.nsmallest()).reset_index()
d5.columns=['REGION','BRANCH','Vs LW']
bty=(df[(df['DATE']==yd)&(df['SALECY']>0)].groupby(['REGION','BRANCH'])['SALECY'].sum()/df[df['DATE']==(lw-pd.to_timedelta(1, unit='d'))].groupby(['REGION','BRANCH'])['SALECY'].sum())-1
d5y=bty.groupby(level=0,group_keys=False).apply(lambda x: x.nsmallest()).reset_index()

def kpi(reg):
    if reg=='BRANCH':
        dly=((df[df['DATE']==td].groupby([reg])['SALECY'].sum()/df[df['DATE']==td].groupby([reg])['SALELY'].sum())-1).reset_index()
        dly.columns=[reg,'Vs LY']
    else:
        dly=((df[(df['DATE']==td)&(df['COMP']=='C')].groupby([reg])['SALECY'].sum()/df[(df['DATE']==td)&(df['COMP']=='C')].groupby([reg])['SALELY'].sum())-1).reset_index()
        dly.columns=[reg,'Vs LY']
    dy=((df[df['DATE']==td].groupby([reg])['SALECY'].sum()/df[df['DATE']==(td-pd.to_timedelta(1, unit='d'))].groupby([reg])['SALECY'].sum())-1).reset_index()
    dy.columns=[reg,'Vs Yst']
    dw=((df[df['DATE']==td].groupby([reg])['SALECY'].sum()/df[df['DATE']==lw].groupby([reg])['SALECY'].sum())-1).reset_index()
    dw.columns=[reg,'Vs LW']
    db=((df[df['DATE']==td].groupby(reg)['SALECY'].sum()/df[df['DATE']==td].groupby([reg])['BUDCY'].sum())-1).reset_index()
    db.columns=[reg,'Vs Bud']
    dy_1=((df[df['DATE']==(td-pd.to_timedelta(1, unit='d'))].groupby([reg])['SALECY'].sum()/df[df['DATE']==(td-pd.to_timedelta(2, unit='d'))].groupby([reg])['SALECY'].sum())-1).reset_index()
    dy_1.columns=[reg,'Vs Yst -1']
    dw_1=((df[(df['DATE']==(td-pd.to_timedelta(1, unit='d')))&(df['SALECY']>0)].groupby([reg])['SALECY'].sum()/df[df['DATE']==(lw-pd.to_timedelta(1, unit='d'))].groupby([reg])['SALECY'].sum())-1).reset_index()
    dw_1.columns=[reg,'Vs LW -1']
    da=(((df[(df['DATE']==td)].groupby([reg])['QTYCY'].sum()/df[df['DATE']==td].groupby([reg])['CUSTCY'].sum())/(df[(df['DATE']==lw)].groupby([reg])['QTYCY'].sum()/df[df['DATE']==lw].groupby([reg])['CUSTCY'].sum()))-1).reset_index()
    da.columns=[reg,'ABS Var']
    dc=(((df[(df['DATE']==td)&(df['CUSTCY']>0)].groupby([reg])['CUSTCY'].sum()/df[(df['DATE']==td)&(df['CUSTCY']>0)].groupby([reg])['FFCY'].sum())/(df[(df['DATE']==lw)].groupby([reg])['CUSTCY'].sum()/df[df['DATE']==lw].groupby([reg])['FFCY'].sum()))-1).reset_index()
    dc.columns=[reg,'Conv Var']
    dff=((df[df['DATE']==td].groupby([reg])['FFCY'].sum()/df[(df['DATE']==lw)].groupby([reg])['FFCY'].sum())-1).reset_index()
    dff.columns=[reg,'FF Var']
    return dy,dw,dly,db,dy_1,dw_1,da,dc,dff

dy,dw,dly,db,dy_1,dw_1,da,dc,dff=kpi('BRANCH')

d5=pd.merge(d5,dy,on='BRANCH',how='left')
d5=pd.merge(d5,dly,on='BRANCH',how='left')
d5=pd.merge(d5,db,on='BRANCH',how='left')
d5=pd.merge(d5,dw_1,on='BRANCH',how='left')
d5=pd.merge(d5,dy_1,on='BRANCH',how='left')
d5=pd.merge(d5,da,on='BRANCH',how='left')
d5=pd.merge(d5,dc,on='BRANCH',how='left')
d5=pd.merge(d5,dff,on='BRANCH',how='left')


d5r=pd.DataFrame()

dyr,dwr,dly_r,db_r,dyrl1,dwr_1,dar,dcr,dffr=kpi('REGION')
d5r['REGION']=['ERO','CRO','CRN','WRO','SRO']
d5r['BRANCH']=d5r['REGION']
d5r['OCC']=["","","","",""]
d5r=pd.merge(d5r,dyr,on='REGION',how='left')
d5r=pd.merge(d5r,dly_r,on='REGION',how='left')
d5r=pd.merge(d5r,db_r,on='REGION',how='left')
d5r=pd.merge(d5r,dwr,on='REGION',how='left')
d5r=pd.merge(d5r,dwr_1,on='REGION',how='left')
d5r=pd.merge(d5r,dyrl1,on='REGION',how='left')
d5r=pd.merge(d5r,dar,on='REGION',how='left')
d5r=pd.merge(d5r,dcr,on='REGION',how='left')
d5r=pd.merge(d5r,dffr,on='REGION',how='left')

d5k=pd.DataFrame()

dyk,dwk,dly_k,db_k,dykl1,dwk_1,dak,dck,dffk=kpi('COUNTRY')
d5k['COUNTRY']=['KSA']
d5k['BRANCH']=d5k['COUNTRY']
d5k['OCC']=[""]
d5k=pd.merge(d5k,dyk,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dly_k,on='COUNTRY',how='left')
d5k=pd.merge(d5k,db_k,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dwk,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dwk_1,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dykl1,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dak,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dck,on='COUNTRY',how='left')
d5k=pd.merge(d5k,dffk,on='COUNTRY',how='left')
d5k=d5k.rename(columns={"COUNTRY":"REGION"})

# for count of days in month

df['LW']=df['DATE']-pd.to_timedelta(7,unit='d')
dfl=df[df['DATE']>'2020-6-1']
dfl=dfl[['BRANCH','DATE','SALECY']]
dfl.columns=['BRANCH','LW','SALELW']
df=pd.merge(df,dfl,on=['BRANCH','LW'],how='left')

lwv=((df[(df['DATE']>lm)&(df['SALECY']>0)&(df['DATE']<=td)&(df['BRANCH']!='13192 - RT-KSA ONLINE')].groupby(['REGION','DATE','BRANCH'])['SALECY'].sum()/df[(df['DATE']>lm)&(df['SALECY']>0)&(df['DATE']<=td)&(df['BRANCH']!='13192 - RT-KSA ONLINE')].groupby(['REGION','DATE','BRANCH'])['SALELW'].sum())-1).reset_index()
lwv.columns=['REGION', 'DATE', 'BRANCH', 'VAR']
tc=lwv.sort_values(['REGION','VAR']).groupby(['REGION','DATE']).head(5)
occ=tc.groupby(['BRANCH'])['VAR'].count().reset_index()
occ.columns=['BRANCH','OCC']

d5=pd.merge(d5,occ,on='BRANCH',how='left')
d5=d5[['REGION', 'BRANCH', 'OCC', 'Vs LW', 'Vs Yst', 'Vs LY','Vs Bud', 'Vs LW -1', 'Vs Yst -1',
       'ABS Var', 'Conv Var', 'FF Var']]

d5r=d5r[['REGION', 'BRANCH', 'OCC', 'Vs LW', 'Vs Yst', 'Vs LY','Vs Bud','Vs LW -1', 'Vs Yst -1',
       'ABS Var', 'Conv Var', 'FF Var']]
d5k=d5k[['REGION', 'BRANCH', 'OCC', 'Vs LW', 'Vs Yst', 'Vs LY','Vs Bud','Vs LW -1', 'Vs Yst -1',
       'ABS Var', 'Conv Var', 'FF Var']]

d5r=d5.append(d5r)
d5r=d5r.append(d5k)
d5r=d5r.sort_values(by=['Vs LW','REGION'])

dst=pd.DataFrame()
dst['REGION']=['ERO','CRO','CRN','WRO','SRO','KSA']
dst['OD']=[1,2,3,4,5,6]

d5r=pd.merge(d5r,dst,on='REGION',how='left')
#d5r=d5r.sort_values(by=['REGION'])
d5r=d5r.sort_values(by=['OD','REGION','Vs LW'],ascending=['TRUE','TRUE','TRUE'])
d5r=pd.merge(d5r,dsm,on=['BRANCH'],how='left')
d5r = d5r.fillna('')
d5r=d5r[['REGION', 'BRANCH', 'SM','OCC', 'Vs LW', 'Vs Yst', 'Vs LY', 'Vs Bud',
       'Vs LW -1', 'Vs Yst -1', 'ABS Var', 'Conv Var', 'FF Var']]

# kpis vs LY, from wtd file

stys=td-pd.to_timedelta(1,unit='d')
std=td
stw=df[df['WK']==df[(df['DATE']==td)]['WK'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
stm=df[df['MN']==df[(df['DATE']==td)]['MN'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
#stl=df[df['MN']==df[(df['DATE']==(stm-pd.to_timedelta(1,unit='d')))]['MN'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
sty=df[df['YR']==df[(df['DATE']==td)]['YR'].unique()[0]]['DATE'].nsmallest(1).iloc[0]

edys=stys
edd=td
edw=td
edm=td
edl=df[df['MN']==df[(df['DATE']==(stm-pd.to_timedelta(1,unit='d')))]['MN'].unique()[0]]['DATE'].nlargest(1).iloc[0]
edy=td
en=['REGION','COUNTRY']
sh=0
ddf=pd.DataFrame()
brn=['RT','T4','RT_T4map']
for br in brn:    
    for e in en:
        dd=wd.kpis(std,edd,e,br)
        dw=wd.kpis(stw,edw,e,br)
        dm=wd.kpis(stm,edm,e,br)
        dy=wd.kpis(sty,edy,e,br)
        fin=pd.merge(dd,dw,on=[e],how='left')
        fin=pd.merge(fin,dm,on=[e],how='left')
        fin=pd.merge(fin,dy,on=[e],how='left')
        fin.columns=['REGION', 'SALE_D', 'ABS_D', 'QTY_D', 'FF_D', 'CONV_D','ASP_D','GP_D','GPLY_D','SALE_W',
           'ABS_W', 'QTY_W', 'FF_W', 'CONV_W','ASP_W','GP_W','GPLY_W','SALE_M','ABS_M', 'QTY_M', 'FF_M',
           'CONV_M','ASP_M','GP_M','GPLY_M','SALE_Y','ABS_Y', 'QTY_Y', 'FF_Y', 'CONV_Y','ASP_Y','GP_Y','GPLY_Y']
        fin=fin[['REGION', 'SALE_D', 'SALE_W','SALE_M','SALE_Y','QTY_D','QTY_W','QTY_M','QTY_Y', 
                 'ABS_D','ABS_W','ABS_M','ABS_Y','FF_D','FF_W','FF_M','FF_Y',
                 'CONV_D','CONV_W','CONV_M','CONV_Y','ASP_D','ASP_W','ASP_M','ASP_Y','GP_D','GP_W','GP_M','GP_Y','GPLY_D','GPLY_W','GPLY_M','GPLY_Y']]
        ddf=ddf.append(fin)

#

wb = xlwings.Book("C:\Python\\Prod\\Templates\\daily_y_lw.xlsx")
Sheet1 = wb.sheets[0]
Sheet1.range(3,1).value = df1.set_index('REGION')
Sheet2 = wb.sheets[1]
Sheet2.range(3,1).value = dft4.set_index('REGION')
Sheet3 = wb.sheets[2]
Sheet3.range(3,1).value = d5r.set_index('REGION')
Sheet4 = wb.sheets[3]
Sheet4.range(3,1).value = ddf.set_index('REGION')
wb.save("C:\Python\Daily\\trend.xlsx")

