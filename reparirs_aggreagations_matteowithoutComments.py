# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 18:14:37 2020

@author: dsoumat
"""

import pandas as pd



gate=pd.read_pickle('gate.pickle')

cockpit=pd.read_pickle('cockpit.pickle') 

orders_for_repairs=pd.read_pickle('oracle_orders.pickle')

oke=pd.read_pickle('oke.pickle')

OBP=pd.read_pickle('OBP.pickle')

item_code_for_repairs=pd.read_pickle('item_code_for_repairs.pickle')








gate=pd.read_excel('Gate_reports_hist_from_2012_to_2017 .xlsx')
gate.to_pickle('gate.pickle')


cockpit=pd.read_excel('repair_cockpit.xlsx')
cockpit=cockpit.to_pickle('cockpit.pickle')


orders_for_repairs=pd.read_excel('oracle_orders.xlsx')
orders_for_repairs=pd.to_pickle('oracle_orders.pickle')

oke=pd.read_excel('oke_contracts.xlsx')
oke=oke.to_pickle('oke.pickle')



item_code_for_repairs=pd.read_excel('item_code_for_repairs.xlsx')
item_code_for_repairs.to_pickle('item_code_for_repairs.pickle')


OBP=pd.read_excel('OBP_orders.xlsx')
OBP.to_pickle('OBP.pickle')




columns_cockpit=gate.columns.tolist()
columns_cockpit.append('SOGGE_Cockpit')
colums_to_remove=['CHARACT_FOUND', 'Unnamed: 27', 'Unnamed: 0', 'Unnamed: 28', 'ITEM_CATEGORY_COCKPIT', 'ANNO_RIFERIMENTO', 'serial_number',
                  'Unnamed: 26', 'PROJECT_NUMBER_COCKPIT', 'JOB_NUMBER_ADJUSTED']
columns_cockpit=list(set(columns_cockpit)-(set(colums_to_remove)))

cockpit=cockpit[columns_cockpit]
cockpit['ANNO_RIFERIMENTO']=pd.DatetimeIndex(cockpit['G3_ACTUAL_END_DATE']).year
cockpit.dropna(subset={'ANNO_RIFERIMENTO'},inplace=True)
cockpit.drop_duplicates(subset='JOB_NUMBER',inplace=True)



gate=gate.append(cockpit)

gate['JOB_NUMBER_CLEANED']=gate.JOB_NUMBER.astype(str)

gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split(',')[0])
gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split(' ')[0])
gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split('/')[0])
gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split('.')[0])
gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split('-')[0])
gate['JOB_NUMBER_CLEANED']=gate['JOB_NUMBER_CLEANED'].apply(lambda x : x.split('_')[0])
gate['JOB_NUMBER_CLEANED'].drop_duplicates(inplace=True)


orders_for_repairs['ORDER_NUMBER']=orders_for_repairs['ORDER_NUMBER'].astype(int)
orders_for_repairs['ORDER_NUMBER']=orders_for_repairs['ORDER_NUMBER'].astype(str)

merge1=gate.merge(orders_for_repairs,left_on='JOB_NUMBER_CLEANED', right_on='ORDER_NUMBER', how='left')

merge1.drop(columns='COMPONENT_CATEGORY_y',inplace=True)
merge1.rename(columns={'COMPONENT_CATEGORY_x':'COMPONENT_CATEGORY'}, inplace=True)

oke.drop_duplicates(subset={'SOGGE','MOTHER_PROJECT','SITE_USE_ID'},inplace=True)
merge1['SHIP_TO_ORG_ID']=merge1['SHIP_TO_ORG_ID'].astype(str)



merge1['SHIP_TO_ORG_ID'].fillna('0',inplace=True)
merge1['SHIP_TO_ORG_ID']=merge1['SHIP_TO_ORG_ID'].astype(int)
merge1['SHIP_TO_ORG_ID']=merge1['SHIP_TO_ORG_ID'].astype(str)

oke['MOTHER_PROJECT']=oke['MOTHER_PROJECT'].astype(str)
oke['SITE_USE_ID']=oke['SITE_USE_ID'].astype(str)
oke.drop_duplicates(subset={'MOTHER_PROJECT','SITE_USE_ID'},inplace=True)

merge2=merge1.merge(oke, left_on=['ORDER_NUMBER','SHIP_TO_ORG_ID'],
                    right_on=['MOTHER_PROJECT','SITE_USE_ID'], how='left')




oke_2=oke[['SOGGE', 'MOTHER_PROJECT','SITE_USE_ID']].drop_duplicates()
oke_2.columns=['SOGGE_2', 'MOTHER_PROJECT_2','SITE_USE_ID_2']

orders_for_repairs_2=orders_for_repairs[['ORDER_NUMBER', 'SHIP_TO_ORG_ID','COSTING_PROJECT']].drop_duplicates()
orders_for_repairs_2.columns=['ORDER_NUMBER_2', 'SHIP_TO_ORG_ID_2','COSTING_PROJECT_2']



df1=merge2.merge(orders_for_repairs_2,left_on='JOB_NUMBER_CLEANED', right_on='COSTING_PROJECT_2', how='left')

df1['SHIP_TO_ORG_ID_2'].fillna(value=0,inplace=True)
df1['SHIP_TO_ORG_ID_2']=df1['SHIP_TO_ORG_ID_2'].astype(int)
df1['SHIP_TO_ORG_ID_2']=df1['SHIP_TO_ORG_ID_2'].astype(str)


df2=df1.merge(oke_2, left_on=['COSTING_PROJECT_2','SHIP_TO_ORG_ID_2'], 
              right_on=['MOTHER_PROJECT_2','SITE_USE_ID_2'], how='left')


df2=df2.merge(oke_2, left_on=['ORDER_NUMBER_2','SHIP_TO_ORG_ID_2'], 
              right_on=['MOTHER_PROJECT_2','SITE_USE_ID_2'], how='left')

df2['SOGGE_COSTING_PROJECT']=df2['SOGGE_2_x']
df2['SOGGE_COSTING_PROJECT'].fillna('', inplace=True)
df2['SOGGE_COSTING_PROJECT'][df2['SOGGE_COSTING_PROJECT']=='']=df2['SOGGE_2_y'][df2['SOGGE_COSTING_PROJECT']=='']

#OBP2.drop(columns='CONTRACT_NUM',inplace=True)

OBP.rename(columns={'C_PROJECT':'PROJECT_NUM'}, inplace=True)
OBP.rename(columns={'C_CONTRACT_NO':'CONTRACT_NUM'}, inplace=True)
OBP.rename(columns={'END_USER_ORACLE_OM':'SOGGE_ORACLE_OM'}, inplace=True)
OBP.rename(columns={'SOGGE':'SOGGE_OBP'}, inplace=True)

OBP['PROJECT_NUM']=OBP['PROJECT_NUM'].astype(str)
OBP['CONTRACT_NUM']=OBP['CONTRACT_NUM'].astype(str)

OBP1=OBP.copy()
OBP1['PROJECT_NUM']=OBP1['CONTRACT_NUM']

OBP1.dropna(subset={'PROJECT_NUM'},inplace=True)



OBP1['PROJECT_NUM']=OBP1['PROJECT_NUM'].apply(lambda x : x.split('_')[0])
OBP1['PROJECT_NUM']=OBP1['PROJECT_NUM'].apply(lambda x : x.split(' ')[0])
OBP2=OBP.append(OBP1)
OBP2.dropna(subset={'PROJECT_NUM'},inplace=True)
OBP2.drop_duplicates(inplace=True)

df3=df2.merge(OBP2,left_on='JOB_NUMBER_CLEANED', right_on='PROJECT_NUM', how='left')

item_code_for_repairs.rename(columns={'SOGGE':'SOGGE_ITEM'}, inplace=True)
df4=df3.merge(item_code_for_repairs,left_on='JOB_NUMBER_CLEANED', right_on='ORDERED_ITEM_ADJ', how='left')

relevant_columns=['JOB_NUMBER_CLEANED','JOB_NUMBER_ADJUSTED','ANNO_RIFERIMENTO','SHOP','ORDER_NUMBER',
                  'LINE_NUMBER','LINE_ID','ORDERED_ITEM','MACHINE_TYPE','MACHINE_MODEL','COMPONENT_CATEGORY',
                  'ITEM_CATEGORY','SHIP_TO_ORG_ID']
relevant_columns.append('SITE_ADDRESS')
relevant_columns.append('SITE_COUNTRY')
relevant_columns.append('SOGGE')
relevant_columns.append('SOGGE_Cockpit')
relevant_columns.append('SOGGE_COSTING_PROJECT')
relevant_columns.append('OBP_END_USER_SOGGE_ADJ')
relevant_columns.append('OBP_SIZE_CHECK')
relevant_columns.append('SOGGE_ORACLE_OM')
relevant_columns.append('SOGGE_ITEM')



df4=df4.assign(ORDER_NUMBER="",ORDERED_ITEM="",SHIP_TO_ORG_ID="",SITE_ADDRESS="",SITE_COUNTRY="")#
df4['ORDER_NUMBER_x'].fillna('empty', inplace=True)
df4['SITE_ADDRESS'][df4['ORDER_NUMBER_x']=='empty']=df4['SITE_ADDRESS_y'][df4['ORDER_NUMBER_x']=='empty']
df4['SITE_COUNTRY'][df4['ORDER_NUMBER_x']=='empty']=df4['SITE_COUNTRY_y'][df4['ORDER_NUMBER_x']=='empty']
df4['SHIP_TO_ORG_ID'][df4['ORDER_NUMBER_x']=='empty']=df4['SHIP_TO_ORG_ID_y'][df4['ORDER_NUMBER_x']=='empty']
df4['ORDERED_ITEM'][df4['ORDER_NUMBER_x']=='empty']=df4['ORDERED_ITEM_y'][df4['ORDER_NUMBER_x']=='empty']
df4['ORDER_NUMBER'][df4['ORDER_NUMBER_x']=='empty']=df4['ORDER_NUMBER_y'][df4['ORDER_NUMBER_x']=='empty']

df4=df4[relevant_columns]

'''
df5['SHIP_TO_ORG_ID'].fillna('', inplace=True)
df5['ORDER_NUMBER'].fillna('', inplace=True)


df5['CONCAT_SO_ORGID']=df5['ORDER_NUMBER'].astype(str)+"_"+df5['SHIP_TO_ORG_ID'].astype(str)
df5['CONCAT_SO_ORGID'][df5['SHIP_TO_ORG_ID']=='']="MISSING SO"
df5['CONCAT_SO_ORGID'][df5['ORDER_NUMBER']=='']="MISSING SO"
df5['SO_MATCH'][df5['CONCAT_SO_ORGID']=="MISSING SO"]="MISSING SO"
df5['SO_MATCH'][df5['CONCAT_SO_ORGID']!="MISSING SO"]="SO FOUND"
'''
#code related to  Sheet for Coverage ELAB

df5=df4[['JOB_NUMBER_CLEANED','JOB_NUMBER_ADJUSTED','ANNO_RIFERIMENTO','SHOP','ORDER_NUMBER',
         'SHIP_TO_ORG_ID','SOGGE','SOGGE_Cockpit','SOGGE_COSTING_PROJECT','OBP_END_USER_SOGGE_ADJ',
         'OBP_SIZE_CHECK','SOGGE_ORACLE_OM','SOGGE_ITEM']].copy()
df5.rename(columns={'SOGGE':'SOGGE_OKE'}, inplace=True)
df5=df5.assign(OKE_SOGGE_FOUND="",COCKPIT_SOGGE_FOUND="",
               SOGGE_OBP="",OBP_SOGGE_FOUND="",OM_SOGGE_FOUND="",
               COSTING_PROJECT_SOGGE_FOUND="",ITEM_SOGGE_FOUND="")


df5['SOGGE_OKE'].fillna('', inplace=True)
df5['SOGGE_Cockpit'].fillna('', inplace=True)
df5['SOGGE_COSTING_PROJECT'].fillna('', inplace=True)
df5['SOGGE_ORACLE_OM'].fillna('', inplace=True)
df5['SOGGE_ITEM'].fillna('', inplace=True)
df5['OKE_SOGGE_FOUND'][df5['SOGGE_OKE']=='']="NO SOGGE"
df5['COCKPIT_SOGGE_FOUND'][df5['SOGGE_Cockpit']=='']="NO SOGGE"
df5['COSTING_PROJECT_SOGGE_FOUND'][df5['SOGGE_COSTING_PROJECT']=='']="NO SOGGE"
df5['OM_SOGGE_FOUND'][df5['SOGGE_ORACLE_OM']=='']="NO SOGGE"
df5['ITEM_SOGGE_FOUND'][df5['SOGGE_ITEM']=='']="NO SOGGE"
df5['OKE_SOGGE_FOUND'][df5['SOGGE_OKE']!='']="SOGGE FOUND"
df5['COCKPIT_SOGGE_FOUND'][df5['SOGGE_Cockpit']!='']="SOGGE FOUND"
df5['COSTING_PROJECT_SOGGE_FOUND'][df5['SOGGE_COSTING_PROJECT']!='']="SOGGE FOUND"
df5['OM_SOGGE_FOUND'][df5['SOGGE_ORACLE_OM']!='']="SOGGE FOUND"
df5['ITEM_SOGGE_FOUND'][df5['SOGGE_ITEM']!='']="SOGGE FOUND"

df5['SOGGE_OBP'][df5['OBP_SIZE_CHECK']==8]=df5['OBP_END_USER_SOGGE_ADJ'][df5['OBP_SIZE_CHECK']==8]
df5['OBP_SOGGE_FOUND'][df5['OBP_SIZE_CHECK']==8]="SOGGE FOUND"
df5['OBP_SOGGE_FOUND'][df5['OBP_SIZE_CHECK']!=8]="NO SOGGE"

df5.rename(columns={'ANNO_RIFERIMENTO':'YEAR'}, inplace=True)


gate['JOB_NUMBER_ADJUSTED']=gate['JOB_NUMBER_CLEANED']
gate.drop(columns='JOB_NUMBER_CLEANED',inplace=True)


writer =pd.ExcelWriter("Final_Output_Repairs_File.xlsx")




workbook  = writer.book#oggetto che raccoglie gli sheets

worksheet1=workbook.add_worksheet('Header')#aggiungo uno sheet in Python
writer.sheets['Header']=worksheet1#aggiungo lo sheet  in Excel
gate.to_excel(writer,sheet_name='Header', index=False)#caricamento dei dati del df nello sheet in excel,index=false  rimuove la colonna degli indici delle righe 

worksheet2=workbook.add_worksheet('Lines')
writer.sheets['Lines']=worksheet2
df4.to_excel(writer,sheet_name='Lines', index=False)

worksheet3=workbook.add_worksheet('Sheet for Coverage ELAB')
writer.sheets['Sheet for Coverage ELAB']=worksheet3
df5.to_excel(writer,sheet_name='Sheet for Coverage ELAB', index=False)



writer.save()  #salviamo oggeto  di scrittura
workbook.close()


        




