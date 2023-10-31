#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import os
import xlwings as xw
import openpyxl
from datetime import date

today_date = date.today().strftime("%Y-%m-%d")



# In[2]:


#-------------------  Reading The Excel Spreadsheets   -------------------#
file_name = 'C:/Users/Alexander Ravazzoni/Documents/WC Lookback/WC Lookback Test Template.xlsx'

#This Weeks Data
new_df = pd.read_excel(file_name, sheet_name='2022 New Data').replace(['Film - Streaming Service','TV - Streaming Service'], 'TV/Film - Streaming')
#Last Weeks Data
old_df = pd.read_excel(file_name, sheet_name='2022 Old Data').replace(['Film - Streaming Service','TV - Streaming Service'], 'TV/Film - Streaming')
#Last Years Data
yoy_df = pd.read_excel(file_name, sheet_name='2021 Data').replace(['Film - Streaming Service','TV - Streaming Service'], 'TV/Film - Streaming')


# In[3]:


#Assigning Column Names to Given DFs
new_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category', 'Account Subcategory', 
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'Q1 Total','Q1 Local',
                  'Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q2 Total', 'Q2 Local', 'Q2 Multi', 'Q2 CTV', 'Q2 PG','Q3 Total',
                  'Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Q4 Total', 'Q4 Local', 'Q4 Multi', 'Q4 CTV','Q4 PG',
                  'Total Amount Net', 'Total Local', 'Total Multi', 'Total CTV', 'Total PG']

old_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category','Account Subcategory',
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'LW Q1 Total','LW Q1 Local',
                  'LW Q1 Multi', 'LW Q1 CTV', 'LW Q1 PG', 'LW Q2 Total', 'LW Q2 Local', 'LW Q2 Multi', 'LW Q2 CTV',
                  'LW Q2 PG','LW Q3 Total', 'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV', 'LW Q3 PG', 'LW Q4 Total',
                  'LW Q4 Local', 'LW Q4 Multi', 'LW Q4 CTV','LW Q4 PG', 'LW Total Amount Net', 'LW Total Local',
                  'LW Total Multi', 'LW Total CTV', 'LW Total PG']

last_year_cols = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category','Account Subcategory',
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent','2021 Q1','2021 Q1 Local',
                  '2021 Q1 Multi', '2021 Q1 CTV', '2021 Q1 PG', '2021 Q2', '2021 Q2 Local', '2021 Q2 Multi',
                  '2021 Q2 CTV', '2021 Q2 PG', '2021 Q3', '2021 Q3 Local', '2021 Q3 Multi', '2021 Q3 CTV', '2021 Q3 PG',
                  '2021 Q4', '2021 Q4 Local', '2021 Q4 Multi', '2021 Q4 CTV', '2021 Q4 PG',
                  '2021 Total',	'2021 Total Local', '2021 Total Multi', '2021 Total CTV', '2021 Total PG']

yoy_df.columns = last_year_cols


# In[ ]:





# In[4]:


#Appending the Agency Indy Column
agency_dict = {
    'Publicis':'Agency',
    'Dentsu':'Agency',
    'Independent Agencies':'Indy',
    'WPP':'Agency',
    'Omnicom':'Agency',
    'Interpublic':'Agency',
    'Horizon':'Agency',
    'Client Direct':'Indy',
    'Havas':'Indy'}

new_df['Agency_Indy'] = new_df['Agency Holding Co'].map(agency_dict)
old_df['Agency_Indy'] = old_df['Agency Holding Co'].map(agency_dict)
yoy_df['Agency_Indy'] = yoy_df['Agency Holding Co'].map(agency_dict)


# In[5]:


#Fill In Zeros for Finances

new_rev_cols = ['Q1 Total', 'Q1 Local', 'Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q2 Total', 'Q2 Local', 'Q2 Multi',
                'Q2 CTV','Q2 PG', 'Q3 Total', 'Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Q4 Total', 'Q4 Local',
                'Q4 Multi', 'Q4 CTV', 'Q4 PG']
for cols in new_rev_cols:
    new_df[cols].fillna(value=0, inplace=True)

old_rev_cols = ['LW Q1 Total', 'LW Q1 Local', 'LW Q1 Multi', 'LW Q1 CTV', 'LW Q1 PG', 'LW Q2 Total', 'LW Q2 Local',
                'LW Q2 Multi', 'LW Q2 CTV','LW Q2 PG', 'LW Q3 Total', 'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV',
                'LW Q3 PG', 'LW Q4 Total', 'LW Q4 Local', 'LW Q4 Multi','LW Q4 CTV', 'LW Q4 PG']
for cols in old_rev_cols:
    old_df[cols].fillna(value=0, inplace=True)

yoy_rev_cols = ['2021 Q1','2021 Q1 Local','2021 Q1 Multi', '2021 Q1 CTV', '2021 Q1 PG', '2021 Q2',
                '2021 Q2 Local', '2021 Q2 Multi','2021 Q2 CTV', '2021 Q2 PG', '2021 Q3', '2021 Q3 Local',
                '2021 Q3 Multi', '2021 Q3 CTV', '2021 Q3 PG', '2021 Q4', '2021 Q4 Local', '2021 Q4 Multi',
                '2021 Q4 CTV', '2021 Q4 PG', '2021 Total', '2021 Total Local', '2021 Total Multi',
                '2021 Total CTV', '2021 Total PG']
for cols in yoy_rev_cols:
    yoy_df[cols].fillna(value=0, inplace=True)
    


# In[6]:


#Removing those unwanted characters
unwanted_chars = ['~', ' - Global', ' - GLOBAL', ' - US', 'Australia', '(blank)']

for char in unwanted_chars:
    new_df = new_df.replace(char, '', regex=True)
    old_df = old_df.replace(char, '', regex=True)
    yoy_df = yoy_df.replace(char, '', regex=True)


# In[7]:


#GROUPING THIS WEEKS DATA
new_child_df = new_df.groupby(['Advertiser']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_parent_df = new_df.groupby(['Ultimate Parent']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_child_q1_df = new_df.groupby(['Advertiser']).agg({'Q1 Total':'sum'}).sort_values(by=['Q1 Total'],ascending=False).reset_index()
new_child_q2_df = new_df.groupby(['Advertiser']).agg({'Q2 Total':'sum'}).sort_values(by=['Q2 Total'],ascending=False).reset_index()
new_child_q3_df = new_df.groupby(['Advertiser']).agg({'Q3 Total':'sum'}).sort_values(by=['Q3 Total'],ascending=False).reset_index()
new_child_q4_df = new_df.groupby(['Advertiser']).agg({'Q4 Total':'sum'}).sort_values(by=['Q4 Total'],ascending=False).reset_index()
new_industry_df = new_df.groupby(['Account Category']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_industry_child_df = new_df.groupby(['Account Category','Advertiser']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_parent_ctv_df = new_df.groupby(['Advertiser']).agg({'Q1 CTV':'sum', 'Q2 CTV':'sum', 'Q3 CTV':'sum', 'Q4 CTV':'sum', 'Total CTV':'sum'}).sort_values(by=['Total CTV'],ascending=False).reset_index()
new_parent_local_df = new_df.groupby(['Advertiser']).agg({'Q1 Local':'sum', 'Q2 Local':'sum', 'Q3 Local':'sum', 'Q4 Local':'sum', 'Total Local':'sum'}).sort_values(by=['Total Local'],ascending=False).reset_index()
new_parent_multi_df = new_df.groupby(['Advertiser']).agg({'Q1 Multi':'sum', 'Q2 Multi':'sum', 'Q3 Multi':'sum', 'Q4 Multi':'sum', 'Total Multi':'sum'}).sort_values(by=['Total Multi'],ascending=False).reset_index()
new_parent_pg_df = new_df.groupby(['Advertiser']).agg({'Q1 PG':'sum', 'Q2 PG':'sum', 'Q3 PG':'sum', 'Q4 PG':'sum', 'Total PG':'sum'}).sort_values(by=['Total PG'],ascending=False).reset_index()
new_agency_indy_df = new_df.groupby(['Agency_Indy']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
new_indy_df = new_df.loc[new_df['Agency_Indy'] == 'Indy']
new_indy_df = new_indy_df.groupby(['Agency']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_agency_df = new_df.loc[new_df['Agency_Indy'] == 'Agency']
new_seller_agency_df = new_agency_df.groupby(['Teammember Name','Agency Holding Co']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Teammember Name'],ascending=True).reset_index()
new_agency_holding_df = new_agency_df.groupby(['Agency Holding Co']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_agency_df = new_agency_df.groupby(['Agency Holding Co','Agency']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Agency Holding Co','Total Amount Net'],ascending=False).reset_index()
new_seller_df = new_df.groupby(['Teammember Name']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()

#DF specific to PSW
new_theater_df = new_df.loc[new_df['Account Subcategory'] == 'Film - Theatrical']
new_theatrical_df = new_theater_df.groupby(['Advertiser']).agg({'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()

new_sub_df = new_df.loc[new_df['Account Category'] == 'Entertainment']
new_ent_sub_df = new_sub_df.groupby(['Account Subcategory']).agg({'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()

new_ent_df = new_sub_df.groupby(['Account Subcategory','Advertiser']).agg({'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()


# In[8]:


#GROUPING LAST WEEKS DATA
old_child_df = old_df.groupby(['Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_parent_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_child_q1_df = old_df.groupby(['Advertiser']).agg({'LW Q1 Total':'sum'}).sort_values(by=['LW Q1 Total'],ascending=False).reset_index()
old_child_q2_df = old_df.groupby(['Advertiser']).agg({'LW Q2 Total':'sum'}).sort_values(by=['LW Q2 Total'],ascending=False).reset_index()
old_child_q3_df = old_df.groupby(['Advertiser']).agg({'LW Q3 Total':'sum'}).sort_values(by=['LW Q3 Total'],ascending=False).reset_index()
old_child_q4_df = old_df.groupby(['Advertiser']).agg({'LW Q4 Total':'sum'}).sort_values(by=['LW Q4 Total'],ascending=False).reset_index()
old_industry_df = old_df.groupby(['Account Category']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_industry_child_df = old_df.groupby(['Account Category','Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_parent_ctv_df = old_df.groupby(['Advertiser']).agg({'LW Total CTV':'sum'}).sort_values(by=['LW Total CTV'],ascending=False).reset_index()
old_parent_local_df = old_df.groupby(['Advertiser']).agg({'LW Total Local':'sum'}).sort_values(by=['LW Total Local'],ascending=False).reset_index()
old_parent_multi_df = old_df.groupby(['Advertiser']).agg({'LW Total Multi':'sum'}).sort_values(by=['LW Total Multi'],ascending=False).reset_index()
old_parent_pg_df = old_df.groupby(['Advertiser']).agg({'LW Total PG':'sum'}).sort_values(by=['LW Total PG'],ascending=False).reset_index()
old_agency_holding_df = old_df.groupby(['Agency Holding Co']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_agency_indy_df = old_df.groupby(['Agency_Indy']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
old_indy_df = old_df.loc[old_df['Agency_Indy'] == 'Indy']
old_indy_df = old_indy_df.groupby(['Agency']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
old_agency_df = old_df.loc[old_df['Agency_Indy'] == 'Agency']
old_seller_agency_df = old_agency_df.groupby(['Teammember Name','Agency Holding Co']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['Teammember Name'],ascending=True).reset_index()
old_agency_df = old_agency_df.groupby(['Agency Holding Co','Agency']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['Agency Holding Co','LW Total Amount Net'],ascending=False).reset_index()
old_seller_df = old_df.groupby(['Teammember Name']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()


#DF specific to PSW
old_theater_df = old_df.loc[old_df['Account Subcategory'] == 'Film - Theatrical']
old_theatrical_df = old_theater_df.groupby(['Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()

old_sub_df = old_df.loc[old_df['Account Category'] == 'Entertainment']
old_ent_sub_df = old_sub_df.groupby(['Account Subcategory']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()

old_ent_df = old_sub_df.groupby(['Account Subcategory','Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()


# In[9]:


#GROUPING LAST YEARS DATA
yoy_child_df = yoy_df.groupby(['Advertiser']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_parent_df = yoy_df.groupby(['Ultimate Parent']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_q1_df = yoy_df.groupby(['Advertiser']).agg({'2021 Q1':'sum'}).sort_values(by=['2021 Q1'],ascending=False).reset_index()
yoy_q2_df = yoy_df.groupby(['Advertiser']).agg({'2021 Q2':'sum'}).sort_values(by=['2021 Q2'],ascending=False).reset_index()
yoy_q3_df = yoy_df.groupby(['Advertiser']).agg({'2021 Q3':'sum'}).sort_values(by=['2021 Q3'],ascending=False).reset_index()
yoy_q4_df = yoy_df.groupby(['Advertiser']).agg({'2021 Q4':'sum'}).sort_values(by=['2021 Q4'],ascending=False).reset_index()
yoy_industry_df = yoy_df.groupby(['Account Category']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_industry_child_df = yoy_df.groupby(['Account Category','Advertiser']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_ctv_df = yoy_df.groupby(['Advertiser']).agg({'2021 Total CTV':'sum'}).sort_values(by=['2021 Total CTV'],ascending=False).reset_index()
yoy_local_df = yoy_df.groupby(['Advertiser']).agg({'2021 Total Local':'sum'}).sort_values(by=['2021 Total Local'],ascending=False).reset_index()
yoy_multi_df = yoy_df.groupby(['Advertiser']).agg({'2021 Total Multi':'sum'}).sort_values(by=['2021 Total Multi'],ascending=False).reset_index()
yoy_pg_df = yoy_df.groupby(['Advertiser']).agg({'2021 Total PG':'sum'}).sort_values(by=['2021 Total PG'],ascending=False).reset_index()
yoy_agency_holding_df = yoy_df.groupby(['Agency Holding Co']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_agency_indy_df = yoy_df.groupby(['Agency_Indy']).agg({'2021 Total':'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
yoy_indy_df = yoy_df.loc[yoy_df['Agency_Indy'] == 'Indy']
yoy_indy_df = yoy_indy_df.groupby(['Agency']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()
yoy_agency_df = yoy_df.loc[yoy_df['Agency_Indy'] == 'Agency']
yoy_agency_df = yoy_agency_df.groupby(['Agency Holding Co','Agency']).agg({'2021 Total':'sum'}).sort_values(by=['Agency Holding Co','2021 Total'],ascending=False).reset_index()

#DF specific to PSW
yoy_theater_df = yoy_df.loc[yoy_df['Account Subcategory'] == 'Film - Theatrical']
yoy_theatrical_df = yoy_theater_df.groupby(['Advertiser']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()

yoy_sub_df = yoy_df.loc[yoy_df['Account Category'] == 'Entertainment']
yoy_ent_sub_df = yoy_sub_df.groupby(['Account Subcategory']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()

yoy_ent_df = yoy_sub_df.groupby(['Advertiser']).agg({'2021 Total':'sum'}).sort_values(by=['2021 Total'],ascending=False).reset_index()


# In[10]:


#----------- Merging This and Last Week Dataframes ------------#
child_result = new_child_df.merge(old_child_df,how='outer', on='Advertiser').merge(yoy_child_df,how='outer', on='Advertiser')
child_result.insert(7,'WoW Change','')
child_result['WoW Change'] = child_result['Total Amount Net'] - child_result['LW Total Amount Net']
child_result.insert(8,'WoW Change %','')
child_result['WoW Change %'] = (child_result['Total Amount Net'] - child_result['LW Total Amount Net'])/child_result['LW Total Amount Net']
child_result['YoY Change %'] = (child_result['Total Amount Net'] - child_result['2021 Total'])/child_result['2021 Total']

parent_result = new_parent_df.merge(old_parent_df,how='outer', on='Ultimate Parent').merge(yoy_parent_df,how='outer', on='Ultimate Parent')
parent_result.insert(7,'WoW Change','')
parent_result['WoW Change'] = parent_result['Total Amount Net'] - parent_result['LW Total Amount Net']
parent_result.insert(8,'WoW Change %','')
parent_result['WoW Change %'] = (parent_result['Total Amount Net'] - parent_result['LW Total Amount Net'])/parent_result['LW Total Amount Net']
parent_result['YoY Change %'] = (parent_result['Total Amount Net'] - parent_result['2021 Total'])/parent_result['2021 Total']

child_q1_result = new_child_q1_df.merge(old_child_q1_df,how='outer', on='Advertiser').merge(yoy_q1_df,how='outer', on='Advertiser')
child_q1_result.insert(3,'WoW Change','')
child_q1_result['WoW Change'] = child_q1_result['Q1 Total'] - child_q1_result['LW Q1 Total']
child_q1_result.insert(4,'WoW Change %','')
child_q1_result['WoW Change %'] = (child_q1_result['Q1 Total'] - child_q1_result['LW Q1 Total'])/child_q1_result['LW Q1 Total']
child_q1_result['YoY Change %'] = (child_q1_result['Q1 Total'] - child_q1_result['2021 Q1'])/child_q1_result['2021 Q1']

child_q2_result = new_child_q2_df.merge(old_child_q2_df,how='outer', on='Advertiser').merge(yoy_q2_df,how='outer', on='Advertiser')
child_q2_result.insert(3,'WoW Change','')
child_q2_result['WoW Change'] = child_q2_result['Q2 Total'] - child_q2_result['LW Q2 Total']
child_q2_result.insert(4,'WoW Change %','')
child_q2_result['WoW Change %'] = (child_q2_result['Q2 Total'] - child_q2_result['LW Q2 Total'])/child_q2_result['LW Q2 Total']
child_q2_result['YoY Change %'] = (child_q2_result['Q2 Total'] - child_q2_result['2021 Q2'])/child_q2_result['2021 Q2']

child_q3_result = new_child_q3_df.merge(old_child_q3_df,how='outer', on='Advertiser').merge(yoy_q3_df,how='outer', on='Advertiser')
child_q3_result.insert(3,'WoW Change','')
child_q3_result['WoW Change'] = child_q3_result['Q3 Total'] - child_q3_result['LW Q3 Total']
child_q3_result.insert(4,'WoW Change %','')
child_q3_result['WoW Change %'] = (child_q3_result['Q3 Total'] - child_q3_result['LW Q3 Total'])/child_q3_result['LW Q3 Total']
child_q3_result['YoY Change %'] = (child_q3_result['Q3 Total'] - child_q3_result['2021 Q3'])/child_q3_result['2021 Q3']

child_q4_result = new_child_q4_df.merge(old_child_q4_df,how='outer', on='Advertiser').merge(yoy_q4_df,how='outer', on='Advertiser')
child_q4_result.insert(3,'WoW Change','')
child_q4_result['WoW Change'] = child_q4_result['Q4 Total'] - child_q4_result['LW Q4 Total']
child_q4_result.insert(4,'WoW Change %','')
child_q4_result['WoW Change %'] = (child_q4_result['Q4 Total'] - child_q4_result['LW Q4 Total'])/child_q4_result['LW Q4 Total']
child_q4_result['YoY Change %'] = (child_q4_result['Q4 Total'] - child_q4_result['2021 Q4'])/child_q4_result['2021 Q4']

industry_result = new_industry_df.merge(old_industry_df,how='outer', on='Account Category').merge(yoy_industry_df,how='outer', on='Account Category')
industry_result.insert(7,'WoW Change','')
industry_result['WoW Change'] = industry_result['Total Amount Net'] - industry_result['LW Total Amount Net']
industry_result.insert(8,'WoW Change %','')
industry_result['WoW Change %'] = (industry_result['Total Amount Net'] - industry_result['LW Total Amount Net'])/industry_result['LW Total Amount Net']
industry_result['YoY Change %'] = (industry_result['Total Amount Net'] - industry_result['2021 Total'])/industry_result['2021 Total']

industry_child_result = new_industry_child_df.merge(old_industry_child_df,how='outer', on=['Account Category','Advertiser']).merge(yoy_industry_child_df,how='outer', on=['Account Category','Advertiser'])
industry_child_result.insert(7,'WoW Change','')
industry_child_result['WoW Change'] = industry_child_result['Total Amount Net'] - industry_child_result['LW Total Amount Net']
industry_child_result.insert(8,'WoW Change %','')
industry_child_result['WoW Change %'] = (industry_child_result['Total Amount Net'] - industry_child_result['LW Total Amount Net'])/industry_child_result['LW Total Amount Net']
industry_child_result['YoY Change %'] = (industry_child_result['Total Amount Net'] - industry_child_result['2021 Total'])/industry_child_result['2021 Total']

ctv_result = new_parent_ctv_df.merge(old_parent_ctv_df,how='outer', on='Advertiser').merge(yoy_ctv_df,how='outer', on='Advertiser')
ctv_result.insert(7,'WoW Change','')
ctv_result['WoW Change'] = ctv_result['Total CTV'] - ctv_result['LW Total CTV']
ctv_result.insert(8,'WoW Change %','')
ctv_result['WoW Change %'] = (ctv_result['Total CTV'] - ctv_result['LW Total CTV'])/ctv_result['LW Total CTV']
ctv_result['YoY Change %'] = (ctv_result['Total CTV'] - ctv_result['2021 Total CTV'])/ctv_result['2021 Total CTV']

local_result = new_parent_local_df.merge(old_parent_local_df,how='outer', on='Advertiser').merge(yoy_local_df,how='outer', on='Advertiser')
local_result.insert(7,'WoW Change','')
local_result['WoW Change'] = local_result['Total Local'] - local_result['LW Total Local']
local_result.insert(8,'WoW Change %','')
local_result['WoW Change %'] = (local_result['Total Local'] - local_result['LW Total Local'])/local_result['LW Total Local']
local_result['YoY Change %'] = (local_result['Total Local'] - local_result['2021 Total Local'])/local_result['2021 Total Local']

multi_result = new_parent_multi_df.merge(old_parent_multi_df,how='outer', on='Advertiser').merge(yoy_multi_df,how='outer', on='Advertiser')
multi_result.insert(7,'WoW Change','')
multi_result['WoW Change'] = multi_result['Total Multi'] - multi_result['LW Total Multi']
multi_result.insert(8,'WoW Change %','')
multi_result['WoW Change %'] = (multi_result['Total Multi'] - multi_result['LW Total Multi'])/multi_result['LW Total Multi']
multi_result['YoY Change %'] = (multi_result['Total Multi'] - multi_result['2021 Total Multi'])/multi_result['2021 Total Multi']

pg_result = new_parent_pg_df.merge(old_parent_pg_df,how='outer', on='Advertiser').merge(yoy_pg_df,how='outer', on='Advertiser')
pg_result.insert(7,'WoW Change','')
pg_result['WoW Change'] = pg_result['Total PG'] - pg_result['LW Total PG']
pg_result.insert(8,'WoW Change %','')
pg_result['WoW Change %'] = (pg_result['Total PG'] - pg_result['LW Total PG'])/pg_result['LW Total PG']
pg_result['YoY Change %'] = (pg_result['Total PG'] - pg_result['2021 Total PG'])/pg_result['2021 Total PG']

agency_holding_result = new_agency_holding_df.merge(old_agency_holding_df,how='outer', on='Agency Holding Co').merge(yoy_agency_holding_df,how='outer', on='Agency Holding Co')
agency_holding_result.insert(7,'WoW Change','')
agency_holding_result['WoW Change'] = agency_holding_result['Total Amount Net'] - agency_holding_result['LW Total Amount Net']
agency_holding_result.insert(8,'WoW Change %','')
agency_holding_result['WoW Change %'] = (agency_holding_result['Total Amount Net'] - agency_holding_result['LW Total Amount Net'])/agency_holding_result['LW Total Amount Net']
agency_holding_result['YoY Change %'] = (agency_holding_result['Total Amount Net'] - agency_holding_result['2021 Total'])/agency_holding_result['2021 Total']

agency_indy_result = new_agency_indy_df.merge(old_agency_indy_df,how='outer', on='Agency_Indy').merge(yoy_agency_indy_df,how='outer', on='Agency_Indy')
agency_indy_result.insert(7,'WoW Change','')
agency_indy_result['WoW Change'] = agency_indy_result['Total Amount Net'] - agency_indy_result['LW Total Amount Net']
agency_indy_result.insert(8,'WoW Change %','')
agency_indy_result['WoW Change %'] = (agency_indy_result['Total Amount Net'] - agency_indy_result['LW Total Amount Net'])/agency_indy_result['LW Total Amount Net']
agency_indy_result['YoY Change %'] = (agency_indy_result['Total Amount Net'] - agency_indy_result['2021 Total'])/agency_indy_result['2021 Total']

indy_result = new_indy_df.merge(old_indy_df,how='outer', on='Agency').merge(yoy_indy_df,how='outer', on='Agency')
indy_result.insert(7,'WoW Change','')
indy_result['WoW Change'] = indy_result['Total Amount Net'] - indy_result['LW Total Amount Net']
indy_result.insert(8,'WoW Change %','')
indy_result['WoW Change %'] = (indy_result['Total Amount Net'] - indy_result['LW Total Amount Net'])/indy_result['LW Total Amount Net']
indy_result['YoY Change %'] = (indy_result['Total Amount Net'] - indy_result['2021 Total'])/indy_result['2021 Total']

agency_result = new_agency_df.merge(old_agency_df,how='outer', on=['Agency Holding Co','Agency']).merge(yoy_agency_df,how='outer', on=['Agency Holding Co','Agency'])
agency_result.insert(8,'WoW Change','')
agency_result['WoW Change'] = agency_result['Total Amount Net'] - agency_result['LW Total Amount Net']
agency_result.insert(9,'WoW Change %','')
agency_result['WoW Change %'] = (agency_result['Total Amount Net'] - agency_result['LW Total Amount Net'])/agency_result['LW Total Amount Net']
agency_result['YoY Change %'] = (agency_result['Total Amount Net'] - agency_result['2021 Total'])/agency_result['2021 Total']

seller_result = new_seller_df.merge(old_seller_df,how='outer', on='Teammember Name')
seller_result.insert(7,'WoW Change','')
seller_result['WoW Change'] = seller_result['Total Amount Net'] - seller_result['LW Total Amount Net']
seller_result.insert(8,'WoW Change %','')
seller_result['WoW Change %'] = (seller_result['Total Amount Net'] - seller_result['LW Total Amount Net'])/seller_result['LW Total Amount Net']

seller_agency_result = new_seller_agency_df.merge(old_seller_agency_df,how='outer', on=['Teammember Name','Agency Holding Co'])
seller_agency_result.insert(8,'WoW Change','')
seller_agency_result['WoW Change'] = seller_agency_result['Total Amount Net'] - seller_agency_result['LW Total Amount Net']
seller_agency_result.insert(9,'WoW Change %','')
seller_agency_result['WoW Change %'] = (seller_agency_result['Total Amount Net'] - seller_agency_result['LW Total Amount Net'])/seller_agency_result['LW Total Amount Net']


#DF specific to PSW
theatrical_child_result = new_theatrical_df.merge(old_theatrical_df,how='outer', on='Advertiser').merge(yoy_theatrical_df,how='outer', on='Advertiser')
theatrical_child_result.insert(3,'WoW Change','')
theatrical_child_result['WoW Change'] = theatrical_child_result['Total Amount Net'] - theatrical_child_result['LW Total Amount Net']
theatrical_child_result.insert(4,'WoW Change %','')
theatrical_child_result['WoW Change %'] = (theatrical_child_result['Total Amount Net'] - theatrical_child_result['LW Total Amount Net'])/theatrical_child_result['LW Total Amount Net']
theatrical_child_result['YoY Change %'] = (theatrical_child_result['Total Amount Net'] - theatrical_child_result['2021 Total'])/theatrical_child_result['2021 Total']

ent_sub_result = new_ent_sub_df.merge(old_ent_sub_df,how='outer', on='Account Subcategory').merge(yoy_ent_sub_df,how='outer', on='Account Subcategory')
ent_sub_result.insert(3,'WoW Change','')
ent_sub_result['WoW Change'] = ent_sub_result['Total Amount Net'] - ent_sub_result['LW Total Amount Net']
ent_sub_result.insert(4,'WoW Change %','')
ent_sub_result['WoW Change %'] = (ent_sub_result['Total Amount Net'] - ent_sub_result['LW Total Amount Net'])/ent_sub_result['LW Total Amount Net']
ent_sub_result['YoY Change %'] = (ent_sub_result['Total Amount Net'] - ent_sub_result['2021 Total'])/ent_sub_result['2021 Total']

ent_result = new_ent_df.merge(old_ent_df,how='outer', on=['Account Subcategory','Advertiser']).merge(yoy_ent_df,how='outer', on='Advertiser')
ent_result.insert(4,'WoW Change','')
ent_result['WoW Change'] = ent_result['Total Amount Net'] - ent_result['LW Total Amount Net']
ent_result.insert(5,'WoW Change %','')
ent_result['WoW Change %'] = (ent_result['Total Amount Net'] - ent_result['LW Total Amount Net'])/ent_result['LW Total Amount Net']
ent_result['YoY Change %'] = (ent_result['Total Amount Net'] - ent_result['2021 Total'])/ent_result['2021 Total']


# In[11]:


#Identifying Lost
lost_dfs = [child_result,parent_result, industry_result,industry_child_result, theatrical_child_result,
                 ent_sub_result, ent_result]

for k in lost_dfs:
    mask1 = ( ((k['Total Amount Net'].isna()) & (k['2021 Total'] > 0)) | ((k['Total Amount Net']==0) & (k['2021 Total'] > 0)) )
    k.loc[mask1, 'WoW Change'] = 'Lost'


# In[12]:


#Identifying New 
child_result['2021 Total'].fillna(value='New This Year', inplace=True)
child_result.loc[child_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

parent_result['2021 Total'].fillna(value='New This Year', inplace=True)
parent_result.loc[parent_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

industry_result['2021 Total'].fillna(value='New This Year', inplace=True)
industry_result.loc[industry_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

industry_child_result['2021 Total'].fillna(value='New This Year', inplace=True)
industry_child_result.loc[industry_child_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

lost_mask = ( ((ctv_result['Total CTV'].isna()) & (ctv_result['2021 Total CTV'] > 0)) | ((ctv_result['Total CTV']==0) & (ctv_result['2021 Total CTV'] > 0)) )
ctv_result.loc[lost_mask, 'WoW Change'] = 'Lost'
ctv_result['2021 Total CTV'].fillna(value='New This Year', inplace=True)
ctv_result.loc[ctv_result['YoY Change %'] == np.inf, '2021 Total CTV'] = 'New This Year'

lost_mask = ( ((local_result['Total Local'].isna()) & (local_result['2021 Total Local'] > 0)) | ((local_result['Total Local']==0) & (local_result['2021 Total Local'] > 0)) )
local_result.loc[lost_mask, 'WoW Change'] = 'Lost'
local_result['2021 Total Local'].fillna(value='New This Year', inplace=True)
local_result.loc[local_result['YoY Change %'] == np.inf, '2021 Total Local'] = 'New This Year'

lost_mask = ( ((multi_result['Total Multi'].isna()) & (multi_result['2021 Total Multi'] > 0)) | ((multi_result['Total Multi'] ==0 ) & (multi_result['2021 Total Multi'] > 0)))
multi_result.loc[lost_mask, 'WoW Change'] = 'Lost'
multi_result['2021 Total Multi'].fillna(value='New This Year', inplace=True)
multi_result.loc[multi_result['YoY Change %'] == np.inf, '2021 Total Multi'] = 'New This Year'

lost_mask = ( ((pg_result['Total PG'].isna()) & (pg_result['2021 Total PG'] > 0)) | ((pg_result['Total PG']==0) & (pg_result['2021 Total PG'] > 0)) )
pg_result.loc[lost_mask, 'WoW Change'] = 'Lost'
pg_result['2021 Total PG'].fillna(value='New This Year', inplace=True)
pg_result.loc[pg_result['YoY Change %'] == np.inf, '2021 Total PG'] = 'New This Year'

agency_holding_result['2021 Total'].fillna(value='New This Year', inplace=True)
agency_holding_result.loc[agency_holding_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

agency_indy_result['2021 Total'].fillna(value='New This Year', inplace=True)
indy_result['2021 Total'].fillna(value='New This Year', inplace=True)
indy_result.loc[indy_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

agency_result['2021 Total'].fillna(value='New This Year', inplace=True)
agency_result.loc[agency_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

theatrical_child_result['2021 Total'].fillna(value='New This Year', inplace=True)
theatrical_child_result.loc[theatrical_child_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

ent_sub_result['2021 Total'].fillna(value='New This Year', inplace=True)
ent_sub_result.loc[ent_sub_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

ent_result['2021 Total'].fillna(value='New This Year', inplace=True)
ent_result.loc[ent_result['YoY Change %'] == np.inf, '2021 Total'] = 'New This Year'

lost_mask = ( ((child_q1_result['Q1 Total'].isna()) & (child_q1_result['2021 Q1'] > 0)) | ((child_q1_result['Q1 Total']==0) & (child_q1_result['2021 Q1'] > 0)) )
child_q1_result.loc[lost_mask, 'WoW Change'] = 'Lost'
child_q1_result['2021 Q1'].fillna(value='New This Year', inplace=True)
child_q1_result.loc[child_q1_result['YoY Change %'] == np.inf, '2021 Q1'] = 'New This Year'

lost_mask = ( ((child_q2_result['Q2 Total'].isna()) & (child_q2_result['2021 Q2'] > 0)) | ((child_q2_result['Q2 Total']==0) & (child_q2_result['2021 Q2'] > 0)) )
child_q2_result.loc[lost_mask, 'WoW Change'] = 'Lost'
child_q2_result['2021 Q2'].fillna(value='New This Year', inplace=True)
child_q2_result.loc[child_q2_result['YoY Change %'] == np.inf, '2021 Q2'] = 'New This Year'

lost_mask = ( ((child_q3_result['Q3 Total'].isna()) & (child_q3_result['2021 Q3'] > 0)) | ((child_q3_result['Q3 Total']==0) & (child_q3_result['2021 Q3'] > 0)) )
child_q3_result.loc[lost_mask, 'WoW Change'] = 'Lost'
child_q3_result['2021 Q3'].fillna(value='New This Year', inplace=True)
child_q3_result.loc[child_q3_result['YoY Change %'] == np.inf, '2021 Q3'] = 'New This Year'

lost_mask = ( ((child_q4_result['Q4 Total'].isna()) & (child_q4_result['2021 Q4'] > 0)) | ((child_q4_result['Q4 Total']==0) & (child_q4_result['2021 Q4'] > 0)) )
child_q4_result.loc[lost_mask, 'WoW Change'] = 'Lost'
child_q4_result['2021 Q4'].fillna(value='New This Year', inplace=True)
child_q4_result.loc[child_q4_result['YoY Change %'] == np.inf, '2021 Q4'] = 'New This Year'


# In[13]:


new_week_cols = [child_result,parent_result,child_q1_result,child_q2_result,child_q3_result,child_q4_result,
                industry_result,industry_child_result,ctv_result,multi_result,local_result,pg_result,
                 theatrical_child_result,ent_sub_result,ent_result]
for j in new_week_cols:
    j.loc[j['WoW Change %'] == np.inf, 'WoW Change'] = 'New This Week'
    j['WoW Change'].fillna(value='New This Week', inplace=True)


# In[14]:


result_df = [child_result, parent_result, child_q1_result, child_q2_result, child_q3_result, child_q4_result, 
             industry_result, industry_child_result, ctv_result, local_result, multi_result, pg_result, 
             agency_holding_result, agency_indy_result, agency_result, seller_result, theatrical_child_result, 
             ent_sub_result]

for dfs in result_df:
    dfs.replace([np.inf, -np.inf], np.nan, inplace=True)


# In[15]:


#DROP THE LAST WEEK TOTAL COLUMN
child_final = child_result.drop('LW Total Amount Net', axis=1)
parent_final = parent_result.drop('LW Total Amount Net', axis=1)
q1_final = child_q1_result.drop('LW Q1 Total', axis=1)
q2_final = child_q2_result.drop('LW Q2 Total', axis=1)
q3_final = child_q3_result.drop('LW Q3 Total', axis=1)
q4_final = child_q4_result.drop('LW Q4 Total', axis=1)
industry_final = industry_result.drop('LW Total Amount Net', axis=1)
industry_child_final = industry_child_result.drop('LW Total Amount Net', axis=1)
ctv_final = ctv_result.drop('LW Total CTV', axis=1)
local_final = local_result.drop('LW Total Local', axis=1)
multi_final = multi_result.drop('LW Total Multi', axis=1)
pg_final = pg_result.drop('LW Total PG', axis=1)
agency_holding_final = agency_holding_result.drop('LW Total Amount Net', axis=1)
agency_indy_final = agency_indy_result.drop('LW Total Amount Net', axis=1)
indy_final = indy_result.drop('LW Total Amount Net', axis=1)
agency_final = agency_result.drop('LW Total Amount Net', axis=1)
seller_final = seller_result.drop('LW Total Amount Net', axis=1)
seller_agency_final = seller_agency_result.drop('LW Total Amount Net', axis=1)
theatrical_child_final = theatrical_child_result.drop('LW Total Amount Net', axis=1)
ent_sub_final = ent_sub_result.drop('LW Total Amount Net', axis=1)
ent_final = ent_result.drop('LW Total Amount Net', axis=1)

no_apple_result = child_result[~child_final['Advertiser'].str.contains('Apple')]
no_apple_final = no_apple_result.drop('LW Total Amount Net', axis=1)


# In[16]:


q1_final = q1_final[(q1_final['Q1 Total'] > 0) | (q1_final['WoW Change'] == 'Lost')] 
q2_final = q2_final[(q2_final['Q2 Total'] > 0) | (q2_final['WoW Change'] == 'Lost')] 
q3_final = q3_final[(q3_final['Q3 Total'] > 0) | (q3_final['WoW Change'] == 'Lost')] 
q4_final = q4_final[(q4_final['Q4 Total'] > 0) | (q4_final['WoW Change'] == 'Lost')] 
ctv_final = ctv_final[(ctv_final['Total CTV'] > 0) | (ctv_final['WoW Change'] == 'Lost')] 
local_final = local_final[(local_final['Total Local'] > 0) | (local_final['WoW Change'] == 'Lost')] 
multi_final = multi_final[(multi_final['Total Multi'] > 0) | (multi_final['WoW Change'] == 'Lost')] 
pg_final = pg_final[(pg_final['Total PG'] > 0) | (pg_final['WoW Change'] == 'Lost')] 


# In[17]:


#EXPORT SECTION
src = file_name

wb = xw.Book(src)

wb.sheets['Summary'].range('L9').options(index=False,header=False).value = child_final
wb.sheets['Summary'].range('W9').options(index=False,header=False).value = parent_final
wb.sheets['Summary'].range('AH9').options(index=False,header=False).value = no_apple_final
wb.sheets['Summary'].range('AS9').options(index=False,header=False).value = q1_final
wb.sheets['Summary'].range('AZ9').options(index=False,header=False).value = q2_final
wb.sheets['Summary'].range('BG9').options(index=False,header=False).value = q3_final
wb.sheets['Summary'].range('BN9').options(index=False,header=False).value = q4_final
wb.sheets['Summary'].range('BU9').options(index=False,header=False).value = industry_final
wb.sheets['Summary'].range('CF9').options(index=False,header=False).value = industry_child_final
wb.sheets['Summary'].range('CR9').options(index=False,header=False).value = theatrical_child_final
wb.sheets['Summary'].range('CY9').options(index=False,header=False).value = ent_sub_final
wb.sheets['Summary'].range('DF9').options(index=False,header=False).value = ent_final
wb.sheets['Summary'].range('DN9').options(index=False,header=False).value = seller_final

wb.sheets['LOB Summary'].range('I14').options(index=False,header=False).value = ctv_final
wb.sheets['LOB Summary'].range('AC14').options(index=False,header=False).value = local_final
wb.sheets['LOB Summary'].range('AW14').options(index=False,header=False).value = multi_final
wb.sheets['LOB Summary'].range('BQ14').options(index=False,header=False).value = pg_final

wb.sheets['Agency Summary'].range('C14').options(index=False,header=False).value = agency_holding_final
wb.sheets['Agency Summary'].range('C24').options(index=False,header=False).value = agency_indy_final
wb.sheets['Agency Summary'].range('O8').options(index=False,header=False).value = indy_final
wb.sheets['Agency Summary'].range('Z8').options(index=False,header=False).value = agency_final
wb.sheets['Agency Summary'].range('AL8').options(index=False,header=False).value = seller_agency_final

#ADD THE RAW DATAFRAMES AS WELL
wb.sheets['Raw Summary'].range('A2').options(index=False,header=True).value = child_result
wb.sheets['Raw Summary'].range('M2').options(index=False,header=True).value = parent_result
wb.sheets['Raw Summary'].range('Y2').options(index=False,header=True).value = no_apple_result
wb.sheets['Raw Summary'].range('AK2').options(index=False,header=True).value = child_q1_result
wb.sheets['Raw Summary'].range('AS2').options(index=False,header=True).value = child_q2_result
wb.sheets['Raw Summary'].range('BA2').options(index=False,header=True).value = child_q3_result
wb.sheets['Raw Summary'].range('BI2').options(index=False,header=True).value = child_q4_result
wb.sheets['Raw Summary'].range('BQ2').options(index=False,header=True).value = industry_result
wb.sheets['Raw Summary'].range('CC2').options(index=False,header=True).value = industry_child_result
wb.sheets['Raw Summary'].range('CP2').options(index=False,header=True).value = theatrical_child_result
wb.sheets['Raw Summary'].range('CX2').options(index=False,header=True).value = ent_sub_result
wb.sheets['Raw Summary'].range('DF2').options(index=False,header=True).value = ent_result
wb.sheets['Raw Summary'].range('DO2').options(index=False,header=True).value = seller_result

wb.sheets['Raw LOB'].range('A2').options(index=False,header=True).value = ctv_result
wb.sheets['Raw LOB'].range('M2').options(index=False,header=True).value = local_result
wb.sheets['Raw LOB'].range('Y2').options(index=False,header=True).value = multi_result
wb.sheets['Raw LOB'].range('AK2').options(index=False,header=True).value = pg_result

wb.sheets['Raw Agency'].range('A2').options(index=False,header=True).value = agency_holding_result
wb.sheets['Raw Agency'].range('A26').options(index=False,header=True).value = agency_indy_result
wb.sheets['Raw Agency'].range('M2').options(index=False,header=True).value = indy_result
wb.sheets['Raw Agency'].range('Y2').options(index=False,header=True).value = agency_result
wb.sheets['Raw Agency'].range('AL2').options(index=False,header=True).value = seller_agency_result

dest = f"C:/Users/Alexander Ravazzoni/Documents/WC Lookback/Lookback {today_date} Template WC.xlsx"

wb.save(dest)

print("Your report is located here -> "+dest)


# In[18]:


Curr_Total = new_df['Total Amount Net'].sum()
LW_Total = old_df['LW Total Amount Net'].sum()
LY_Total = yoy_df['2021 Total'].sum()

q1_total_percent = '{:.1%}'.format(new_df['Q1 Total'].sum()/new_df['Total Amount Net'].sum())
q2_total_percent = '{:.1%}'.format(new_df['Q2 Total'].sum()/new_df['Total Amount Net'].sum())
q3_total_percent = '{:.1%}'.format(new_df['Q3 Total'].sum()/new_df['Total Amount Net'].sum())
q4_total_percent = '{:.1%}'.format(new_df['Q4 Total'].sum()/new_df['Total Amount Net'].sum())

ctv_total = new_df['Total CTV'].sum()
local_total = new_df['Total Local'].sum()
multi_total = new_df['Total Multi'].sum()
pg_total = new_df['Total PG'].sum()

ctv_percent = '{:.1%}'.format(ctv_total/Curr_Total)
local_percent = '{:.1%}'.format(local_total/Curr_Total)
multi_percent = '{:.1%}'.format(multi_total/Curr_Total)
pg_percent = '{:.1%}'.format(pg_total/Curr_Total)

this_week_diff = "${:,.0f}".format(Curr_Total - LW_Total)
yoy_diff = "${:,.0f}".format(Curr_Total - LY_Total)

if Curr_Total - LW_Total > 0:
    wow_rev_dir = 'up'
else:
    wow_rev_dir = 'down'
    
if Curr_Total - LY_Total > 0:
    yoy_rev_dir = 'up'
else:
    yoy_rev_dir = 'down'

top_cats = industry_final.nlargest(3,['Total Amount Net'])
top_cats_list = top_cats['Account Category'].tolist()
top_categories = ' '.join([str(item+', ') for item in top_cats_list])

top_cat_sum = top_cats['Total Amount Net'].sum()
industry_percent = '{:.1%}'.format(top_cat_sum / Curr_Total)

wow_top_cat = industry_final.nlargest(1,['WoW Change %'])
yoy_top_cat = industry_final.nlargest(3,['YoY Change %'])
yoy_top_indy = yoy_top_cat['Account Category'].tolist()

new_category = new_df.groupby(['Account Category']).agg({'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
new_subcategory = new_df.groupby(['Account Category','Account Subcategory']).agg({'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'],ascending=False).reset_index()
cat_vs_sub = new_subcategory.merge(new_category,how='left', on='Account Category')
cat_vs_sub['Sub Weight'] = cat_vs_sub['Total Amount Net_x']/cat_vs_sub['Total Amount Net_y']
subcategory_df = cat_vs_sub[cat_vs_sub['Account Category'].isin(top_cats_list)]
sub_df = subcategory_df.nlargest(2,['Sub Weight']).reset_index()


# In[19]:


sub1_percent = '{:.1%}'.format(sub_df.loc[0][4])
sub2_percent = '{:.1%}'.format(sub_df.loc[1][4])

wow_top_deals_percent = child_final.nlargest(3,['WoW Change %'])
top_advertisers_wow = wow_top_deals_percent['Advertiser'].tolist()
top_accounts = ' '.join([str(acc+', ') for acc in top_advertisers_wow])
wow_top_accounts_percent = '{:.1%}'.format(wow_top_deals_percent['WoW Change'].sum()/(wow_top_deals_percent['Total Amount Net'].sum() - wow_top_deals_percent['WoW Change'].sum()))

added_deals = child_final.loc[(child_final['WoW Change'] != 'New This Week') & (child_final['WoW Change %'] > 0)]
added_deals_sum = "${:,.0f}".format(added_deals['WoW Change'].sum())

less_deals = child_final.loc[(child_final['WoW Change'] != 'New This Week') & (child_final['WoW Change %'] < 0)]
less_deals_sum = "${:,.0f}".format(less_deals['WoW Change'].sum())

retained_sum = "${:,.0f}".format(added_deals['WoW Change'].sum() + less_deals['WoW Change'].sum())

new_deals = child_final.loc[child_final['WoW Change'] == 'New This Week']
#new_deals_sum = "${:,.0f}".format(new_deals['Total Amount Net'].sum())
#new_advertisers = list(new_deals['Advertiser'])
#new_account_count = len(new_advertisers)
#new_advertisers = new_deals['Advertiser'].tolist()
#new_accounts = ' '.join([str(item+', ') for item in new_advertisers])

insights_prompt = f'''
 - Revenue is currently {wow_rev_dir} {this_week_diff} this week and {yoy_rev_dir} {yoy_diff} compared to last year
 - Q1 Booked is {q1_total_percent}, Q2 - {q2_total_percent}, Q3 - {q3_total_percent}, Q4 - {q4_total_percent}
 - Retained accounts contribute {retained_sum}, with {added_deals_sum} in incremental spend, and {less_deals_sum} in decremental spend

 - Some healthy growth from {top_accounts} this week with an average WoW Change % of {wow_top_accounts_percent}
 - Top three industries are {top_categories} making up {industry_percent} of whats booked this year
 - {sub_df.loc[0][1]} is {sub1_percent} of {sub_df.loc[0][0]} spend, and {sub_df.loc[1][1]} takes up {sub2_percent} of {sub_df.loc[1][0]}
 - The most improved industry against last year is {yoy_top_indy[0]}, a lot of this is driven from 
 - CTV makes up {ctv_percent} of booked revenue, Local - {local_percent}, Multi - {multi_percent}, and PG - {pg_percent}
'''


# In[20]:


print(insights_prompt)
print("Your report is located here -> "+dest)


# In[ ]:





# In[ ]:





# In[ ]:




