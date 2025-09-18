# -*- coding: utf-8 -*-
"""
Created on Mon Jun 22 23:13:14 2020

@author: rehan
"""

import pandas as pd
import numpy as np
#import datetime
import xlwings


df=pd.read_pickle(r'C:\Python\pk\sale')
dt4=pd.read_pickle(r'C:\Python\pk\t4')
dt4['SEL']=np.where(((dt4['BRANCH']=='23072 - T4 SHIFA')&(dt4['DATE']>'2021-3-6')&(dt4['YR']==2020)),1,0)
dt4=dt4[dt4['SEL']==0]
dt4=dt4.drop('SEL',axis=1)
dm=pd.read_csv(r'C:\Python\Prod\type.csv')
dc=pd.read_csv(r'C:\Python\Prod\ksa_comp.csv')
dct=pd.read_excel(r'C:\Python\Prod\syscomp_t4.xlsx',skiprows=1)
dw=pd.read_csv(r'C:\Python\Prod\weeks.csv',parse_dates=['DATE'],dayfirst=True)
arb=pd.read_csv(r'C:\Python\Prod\arbr_all.csv')
di=pd.read_csv(r'C:\Python\Prod\rtt4.csv')

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
df=pd.merge(df,dw,on='DATE',how='left')
dt4=pd.merge(dt4,dw,on='DATE',how='left')


df=pd.merge(df,dc,on='BRANCH',how='left')
df=pd.merge(df,dm,on='BRANCH',how='left')
dt4=pd.merge(dt4,dct,on='BRANCH',how='left')


tr=pd.read_csv(r'C:\Python\Prod\pr_br.csv')

df=df[df['BRANCH'].isin(tr['BRANCH'])]

td=pd.to_datetime('today')-pd.to_timedelta(1, unit='d')
td=td.normalize()
'''
dcv=(df[(df['YR']==2019)&(df['DATE']<'2020-2-23')].groupby('BRANCH')['SALECY'].sum()/df[(df['YR']==2019)&(df['DATE']<'2020-2-23')].groupby('BRANCH')['SALELY'].sum()).reset_index()
dcv.columns=['BRANCH','CVC']
dnc=df[(df['DATE']>'2020-3-14')&(df['DATE']<'2020-6-28')]
dnc=pd.merge(dnc,dcv,on='BRANCH',how='left')
dnc['CV']=dnc['SALELY']*dnc['CVC']
dnc=dnc[['BRANCH','DATE','CV']]
dnc['DATE']=dnc['DATE']+pd.to_timedelta(364,unit='d')

df=pd.merge(df,dnc,on=['BRANCH','DATE'],how='left')
'''
def kpis(st,ed,reg,br):
    if br=='RT':
        dfa=df[df['COMP']=='C']
        dfl=df[(df['COMP']=='C')]
        dfl['DATE']=dfl['DATE']+pd.to_timedelta(364,unit='d')
    
    elif br=='RT_T4map':
        dfa=df[(df['COMP']=='C')&(df['BRANCH'].isin(di['RT BRANCH']))]
        dfl=df[(df['COMP']=='C')&(df['BRANCH'].isin(di['RT BRANCH']))]
        dfl['DATE']=dfl['DATE']+pd.to_timedelta(364,unit='d')
        
    else:
        dfa=dt4[dt4['COMP']=='C']
        dfl=dt4[(dt4['COMP']=='C')]
        dfl['DATE']=dfl['DATE']+pd.to_timedelta(364,unit='d')
        
    sl=((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['SALECY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['SALECY'].sum())-1).reset_index()
    #bud=((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['SALECY'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['BUDCY'].sum())-1).reset_index()
    #nc=((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['SALECY'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['CV'].sum())-1).reset_index()
    qt=((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['QTYCY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['QTYCY'].sum())-1).reset_index()
    ab=(((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['QTYCY'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['CUSTCY'].sum())/(dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['QTYCY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['CUSTCY'].sum()))-1).reset_index()
    asp=(((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['SALECY'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['QTYCY'].sum())/(dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['SALECY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['QTYCY'].sum()))-1).reset_index()
    gp=(dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['GPV'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['SALEXCY'].sum()).reset_index()
    gply=(dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['GPV'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['SALEXCY'].sum()).reset_index()
    co=(((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['CUSTCY'].sum()/dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['FFCY'].sum())/(dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['CUSTCY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['FFCY'].sum()))-1).reset_index()
    ff=((dfa[(dfa['DATE']>=st)&(dfa['DATE']<=ed)].groupby([reg])['FFCY'].sum()/dfl[(dfl['DATE']>=st)&(dfl['DATE']<=ed)].groupby([reg])['FFCY'].sum())-1).reset_index()
    #sl=pd.merge(sl,bud,on=[reg],how='left')
    #sl=pd.merge(sl,nc,on=[reg],how='left')
    sl=pd.merge(sl,ab,on=[reg],how='left')
    sl=pd.merge(sl,qt,on=[reg],how='left')
    sl=pd.merge(sl,ff,on=[reg],how='left')
    sl=pd.merge(sl,co,on=[reg],how='left')
    sl=pd.merge(sl,asp,on=[reg],how='left')
    sl=pd.merge(sl,gp,on=[reg],how='left')
    sl=pd.merge(sl,gply,on=[reg],how='left')
    sl.columns=[reg,'SALE','ABS','QTY','FF','CONV','ASP','GP','GPLY']
    return sl

stys=td-pd.to_timedelta(1,unit='d')
std=td
stw=df[df['WK']==df[(df['DATE']==td)]['WK'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
stm=df[df['MN']==df[(df['DATE']==td)]['MN'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
stl=df[df['MN']==df[(df['DATE']==(stm-pd.to_timedelta(1,unit='d')))]['MN'].unique()[0]]['DATE'].nsmallest(1).iloc[0]
sty=df[df['YR']==df[(df['DATE']==td)]['YR'].unique()[0]]['DATE'].nsmallest(1).iloc[0]

edys=stys
edd=td
edw=td
edm=td
edl=df[df['MN']==df[(df['DATE']==(stm-pd.to_timedelta(1,unit='d')))]['MN'].unique()[0]]['DATE'].nlargest(1).iloc[0]
edy=td

en=['REGION','AR','BRANCH']
sh=0
for e in en:
    dd=kpis(std,edd,e,'RT')
    dw=kpis(stw,edw,e,'RT')
    dm=kpis(stm,edm,e,'RT')
    dl=kpis(stl,edl,e,'RT')
    dy=kpis(sty,edy,e,'RT')
    dys=kpis(stys,edys,e,'RT')
    fin=pd.merge(dys,dd,on=[e],how='left')
    fin=pd.merge(fin,dw,on=[e],how='left')
    fin=pd.merge(fin,dm,on=[e],how='left')
    fin=pd.merge(fin,dl,on=[e],how='left')
    fin=pd.merge(fin,dy,on=[e],how='left')
    if e=='BRANCH':
        fin=pd.merge(fin,arb,on=['BRANCH'],how='left')
        fin.insert(1,'AREA',fin['AR'])
        fin.insert(0,'REG',fin['REGION'])        
        fin=fin.drop(['AR','REGION'],axis=1)        
    elif e=='AR':
        fin=pd.merge(fin,arb,on=['AR'],how='left')
        fin=fin.drop_duplicates(subset='AR')
        fin.insert(1,'REG',fin['REGION'])
        fin=fin.drop(['BRANCH','REGION'],axis=1)
    '''
    wb = xlwings.Book("E:\PythonSaves\\Prod\\Templates\\wtd_mtd.xlsx")
    Sheet1 = wb.sheets[sh]
    Sheet1.range(2,1).value = fin.set_index(e)
    sh=sh+1
    '''
        

    

