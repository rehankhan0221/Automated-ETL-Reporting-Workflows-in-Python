# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 11:31:45 2020

@author: B15400
"""

import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import six
#import pymsteams

os.chdir(r'C:\Python')

#web_hook="https://outlook.office.com/webhook/113d8b6c-902b-4127-af70-10e56ac3d390@0db8f584-2ef2-4402-aae9-3d13e4b08635/IncomingWebhook/c0b9945d206743abbe751f88e4097582/549693ff-84d0-4a64-a132-2083b0ae1c85"



df=pd.read_excel(r'C:\Python\Hourly Sales Report - KSA - 1PM.xlsx',skiprows=32,skipfooter=3)
dcol=pd.read_excel(r'C:\Python\Hourly Sales Report - KSA - 1PM.xlsx',skiprows=1,skipfooter=193)
df=df[~df['Branch Full Name'].isna()]   
df['REGION']=df['REGION'].fillna(method='ffill')
#df=df[df['REGION'].str.contains('Summary')]
df=df.drop(['SUB REGION','BUDGET','BUD VAR%', 'YDAY VAR%', 'LW VAR%', 'GCQ', 'GCV'],axis=1)
df=df[df['TODAY']>0]
df2=df.groupby('REGION').agg({'TODAY':'sum','YESTERDAY':'sum','LW':'sum'}).reset_index()
dk=pd.DataFrame(df.agg({'TODAY':'sum','YESTERDAY':'sum','LW':'sum'})).T
dk.insert(0,'REGION','KSA')
dnum=df[df['TODAY']>0].groupby('REGION')['Branch Full Name'].nunique().reset_index()
dnum.columns=['REGION','#Updated']
ksup=dnum['#Updated'].sum()
dnum.loc[len(dnum)] = ['KSA',ksup]

          
df2=df2.append(dk)
dy=(df2.groupby('REGION')['TODAY'].sum()/df2.groupby('REGION')['YESTERDAY'].sum()-1).mul(100).round(1).astype(str).add('%').reset_index()
dy.columns=['REGION','Var YST']
dl=(df2.groupby('REGION')['TODAY'].sum()/df2.groupby('REGION')['LW'].sum()-1).mul(100).round(1).astype(str).add('%').reset_index()
dl.columns=['REGION','Var LW']
dff=df2.drop(['YESTERDAY','LW'],axis=1)

dff['TODAY']=dff['TODAY'].apply(lambda x : "{:,.0f}".format(x))
dff=pd.merge(dff,dy,on='REGION',how='left')
dff=pd.merge(dff,dl,on='REGION',how='left')
dff=pd.merge(dff,dnum,on='REGION',how='left')

def render_mpl_table(data, col_width=2.5, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')

    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)

    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in six.iteritems(mpl_table._cells):
        cell.set_edgecolor(edge_color)
        if k[0] == 0:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        elif k[0]==6:
            cell.set_text_props(weight='bold')
            cell.set_facecolor('#b1f3b3')
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    plt.savefig('1pm.png')
    return ax

render_mpl_table(dff, header_columns=0, col_width=2.0)
print(dcol.columns[-1])
d=dcol.columns[-1]
#api.whatsapp.com/send?phone=919891269918&text=test
'''
teams_message = pymsteams.connectorcard(web_hook)
teams_message.text(d)
#teams_message.text(dff.set_index('REGION').to_html())
teams_message.send()  

teams_message = pymsteams.connectorcard(web_hook)
#teams_message.text(d)
teams_message.text(dff.set_index('REGION').to_html())
teams_message.send()  
'''
