    # -*- coding: utf-8 -*-
"""
Created on Fri Jan  6 23:45:11 2023

@author: 807181
"""

import pandas as pd
import numpy as np
import xlwings as xw

disp = pd.read_excel(r'C:\WMS Py\Prod\02 Dispatch Detail Report - History.xlsx')
h5 = pd.read_excel(r"C:\WMS Py\Prod\H5 Details.xlsx",sheet_name="Sheet1")
disp['H1'] = disp['H1'].astype(str).str.replace('\s+', '').replace('nan', np.nan)
disp['H2'] = disp['H2'].astype(str).str.replace('\s+', '').replace('nan', np.nan)
td=pd.to_datetime('today').normalize()
yst=td-pd.to_timedelta(1,unit='d')
disp['MERGE']=disp['H1']+disp['H2']
disp=pd.merge(disp,h5,on='MERGE',how='left')
disp['REMARKS']=disp['INVOICENO'].str[:3]
disp['DISTRI']=disp['REGION']+disp['H1_x']+disp['H5']
disp=disp[['BRAND', 'COMPANY', 'VIR WH', 'INVOICENO', 'SHIPDATE', 'SHOPCODE',
       'SHOPNAME', 'REGION', 'CARRIERID', 'TRAILERID', 'PALLETID', 'CARTONID',
       'SKU', 'SKUDesc', 'H1_x', 'H2_x', 'H3', 'H4', 'PROCGROUP',
       'SHIPSCHEDULE', 'COMPONENT QTY', 'PACK QTY', 'RATIO', 'ITEM FLAG',
       'MERGE', 'H5','REMARKS','DISTRI']]

disp.columns=['BRAND', 'COMPANY', 'VIR WH', 'INVOICENO', 'SHIPDATE', 'SHOPCODE',
       'SHOPNAME', 'REGION', 'CARRIERID', 'TRAILERID', 'PALLETID', 'CARTONID',
       'SKU', 'SKUDesc', 'H1', 'H2', 'H3', 'H4', 'PROCGROUP',
       'SHIPSCHEDULE', 'COMPONENT QTY', 'PACK QTY', 'RATIO', 'ITEM FLAG',
       'MERGE', 'H5','REMARKS','DISTRI']

disp['BRAND']=np.where((disp['SHOPCODE']==130211),'eCom',disp['BRAND'])
disp['SHIPDATE']=disp['SHIPDATE'].dt.date
disppiv=disp[(disp['BRAND']=='REDTAG')&(disp['H1']!='SFA')&(disp['H1']!='UNI')&(disp['H1']!='DIS')].groupby(['DISTRI','SHIPDATE']).agg({'COMPONENT QTY':'sum'}).unstack()
dispbea=disp[(disp['BRAND']=='REDTAG')&(disp['H1']=='BEA')].groupby(['H5','SHIPDATE']).agg({'COMPONENT QTY':'sum'}).unstack()
dispshop=disp[(disp['BRAND']=='REDTAG')].groupby(['SHOPCODE','SHIPDATE']).agg({'COMPONENT QTY':'sum'}).unstack()
alloc=pd.read_excel(r'C:\WMS Py\Prod\42.xlsx',skiprows=3)
regions=pd.read_excel(r'C:\WMS Py\Prod\Shop_Code_Region_2022 - Copy.xlsb')
alloc=pd.merge(alloc,regions,on='Toloc',how='left')
alloc['Allocation DATE']=alloc['Allocation DATE'].dt.strftime("%Y-%m-%d")
alloc['Allocation DATE']=pd.to_datetime(alloc['Allocation DATE'])
allocregion=alloc[(alloc['storerkey']=='REDTAG')&(alloc['H1']!='SFA')&(alloc['H1']!='UNI')&(alloc['H1']!='DIS')&(alloc['H1']!='BEA')&(alloc['Allocation DATE']<=yst)].groupby(['REGION_CODE']).agg({'COMP QTY':'sum'})
allocitem=alloc[(alloc['storerkey']=='REDTAG')&(alloc['H1']!='SFA')&(alloc['H1']!='UNI')&(alloc['H1']!='DIS')].groupby(['ITEM FLAG','Allocation DATE']).agg({'COMP QTY':'sum'}).unstack()
alloch1=alloc[(alloc['storerkey']=='REDTAG')&(alloc['H1']!='SFA')&(alloc['H1']!='UNI')&(alloc['H1']!='DIS')].groupby(['H1','Allocation DATE']).agg({'COMP QTY':'sum'}).unstack()

asn=pd.read_excel(r'C:\WMS Py\Prod\ASN vs ORDER KSA RDC Report.xlsx',sheet_name='RT ASN vs ORDER KSA RDC Report')
asn['MERGE']=asn['H1']+asn['H2']
asn=pd.merge(asn,h5,on='MERGE',how='left')
asn=asn[['BRAND', 'WMS ASN NBR', 'ASN TYPE', 'ASN STATUS', 'RMS PO NBR',
       'CDC INVOICE', 'CUST. INV.', 'REMARKS', 'PHY WH', 'VIR WH', 'ITEM', 'H1_x',
       'H2_x', 'H3', 'H4', 'ASN QTY EXPECTED', 'ASN QTY RECEIVED',
       'ORIG ORD QTY', 'OPEN ORD QTY', 'ALLOC ORD QTY', 'PICKED ORD QTY',
       'SHIPPED ORD QTY', 'ALLOCATION', 'BOM QTY', 'BOM QTY.1',
       'MERGE', 'H5']]
asn.columns=['BRAND', 'WMS ASN NBR', 'ASN TYPE', 'ASN STATUS', 'RMS PO NBR',
       'CDC INVOICE', 'CUS INV', 'REMARK', 'PHY WH', 'VIR WH', 'ITEM', 'H1',
       'H2', 'H3', 'H4', 'ASN QTY EXPECTED', 'ASN QTY RECEIVED',
       'ORIG ORD QTY', 'OPEN ORD QTY', 'ALLOC ORD QTY', 'PICKED ORD QTY',
       'SHIPPED ORD QTY', 'ALLOCATION'  , 'BOM QTY', 'BOM QTY.1',
       'MERGE', 'H5']

wb1 = xw.Book(r"C:\WMS Py\Templates\KSA Distribution Monitoring Week-26.xlsb")
dispsheet = wb1.sheets['02 Dispatch Detail Report - His']
dispsheet.range(1,1).value= disppiv
dispsheet.range(1,8).value= dispbea

wb = xw.Book(r"C:\WMS Py\Templates\Shopwise RDC Dispatch Qty Week 27.xlsx")
dispsheet = wb.sheets['02 Dispatch Detail Report - His']
dispsheet.range(1,1).value= dispshop

pen=pd.read_pickle(r'C:\WMS Py\codes\pending')
pen['BRAND']=np.where((pen['BRANCHCODE']==130211),'eCom',pen['BRAND'])
penpiv=pen[(pen['BRAND']=='REDTAG')].groupby(['DISTRI']).agg({'EA_QTY':'sum'})
penbea=pen[(pen['H1']=='BEA')].groupby(['H5']).agg({'EA_QTY':'sum'}).unstack()
penshop=pen.groupby(['BRANCHCODE']).agg({'EA_QTY':'sum'})


pen1sheet = wb1.sheets['31 KSA RDC Pending For Picking ']
pen1sheet.range(1,1).value= penpiv
pen1sheet.range(1,8).value= penbea

pensheet = wb.sheets['31 KSA RDC Pending For Picking ']
pensheet.range(1,1).value= penshop

rf=pd.read_pickle(r'C:\WMS Py\codes\rfd')
rf['BRAND']=np.where((rf['SHOPCODE']==130211),'eCom',rf['BRAND'])
rfppiv=rf[(rf['BRAND']=='REDTAG')].groupby(['DISTRI','LOCATION']).agg({'COMPONENT QTY':'sum'}).unstack()
rfbea=rf[(rf['H1']=='BEA')].groupby(['H5','LOCATION']).agg({'COMPONENT QTY':'sum'}).unstack()
rfshop=rf.groupby(['SHOPCODE']).agg({'COMPONENT QTY':'sum'})

rf1sheet = wb1.sheets['data rfd']
rf1sheet.range(1,1).value= rfppiv
rf1sheet.range(1,8).value= rfbea


rfsheet = wb.sheets['rfd data']
rfsheet.range(1,1).value= rfshop

allocsheet = wb1.sheets['alloc']
allocsheet.range(1,1).value= alloch1

allocsheet = wb1.sheets['alloc']
allocsheet.range(13,1).value= allocitem

allocsheet = wb1.sheets['alloc']
allocsheet.range(22,1).value= allocregion

asnsheet = wb1.sheets['CROSS']
asnsheet.range(1,1).value= asn.set_index('BRAND')

#ageing= xw.Book(r'C:\WMS Py\Templates\06.Ageing more than 3 days Open order.xlsx')
#ageingpicsheet=ageing.sheets['pick data']
#ageingpicsheet.range(1,1).value=pen.set_index('BRAND')

#ageingrfdsheet=ageing.sheets['rfd data']

#ageingrfdsheet.range(1,1).value=rf.set_index('BRAND')

