#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import xlwings as xw
import openpyxl
from datetime import date
import numpy as np

today_date = '10-27-2023'


# In[2]:


#-------------------  Reading The Excel Spreadsheets   -------------------#
file_name = 'C:/Users/Alexander Ravazzoni/Documents/Agency Lookback/Agency Lookback Template.xlsx'

#This Weeks Data
current_df = pd.read_excel(file_name, sheet_name='2023 New Data')
#Last Weeks Data
last_week_df = pd.read_excel(file_name, sheet_name='2023 Old Data')
#Last Years Data
last_year_df = pd.read_excel(file_name, sheet_name='2022 Data')


# In[3]:


#Assigning Column Names to Given DFs
current_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category', 'Account Subcategory', 
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent','Q4 Total', 'Q4 Local', 'Q4 Multi',
                  'Q4 CTV','Q4 PG', 'Q1 Total','Q1 Local','Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q2 Total', 'Q2 Local',
                  'Q2 Multi', 'Q2 CTV', 'Q2 PG','Q3 Total','Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Total Amount Net',
                  'Total Local', 'Total Multi', 'Total CTV', 'Total PG']

last_week_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category','Account Subcategory',
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'LW Q4 Total','LW Q4 Local', 
                  'LW Q4 Multi', 'LW Q4 CTV','LW Q4 PG','LW Q1 Total','LW Q1 Local', 'LW Q1 Multi', 'LW Q1 CTV',
                  'LW Q1 PG', 'LW Q2 Total', 'LW Q2 Local', 'LW Q2 Multi', 'LW Q2 CTV','LW Q2 PG','LW Q3 Total',
                  'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV', 'LW Q3 PG', 'LW Total Amount Net', 'LW Total Local',
                  'LW Total Multi', 'LW Total CTV', 'LW Total PG']

last_year_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category','Account Subcategory',
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent','2022 Q4','2022 Q4 Local', 
                  '2022 Q4 Multi', '2022 Q4 CTV', '2022 Q4 PG','2022 Q1','2022 Q1 Local','2022 Q1 Multi', '2022 Q1 CTV',
                  '2022 Q1 PG', '2022 Q2', '2022 Q2 Local', '2022 Q2 Multi','2022 Q2 CTV', '2022 Q2 PG', '2022 Q3',
                  '2022 Q3 Local', '2022 Q3 Multi', '2022 Q3 CTV', '2022 Q3 PG','2022 Total','2022 Total Local',
                  '2022 Total Multi', '2022 Total CTV', '2022 Total PG']


# In[4]:


#Appending the Agency Indy Column
agency_dict = {
    'Publicis':'Publicis',
    'Dentsu':'Dentsu',
    'Independent Agencies':'Indy Client',
    'WPP':'WPP',
    'Omnicom':'Omnicom',
    'Interpublic':'Interpublic',
    'Horizon':'Horizon',
    'Client Direct':'Indy Client',
    'Havas':'Indy Client'}

current_df['Agency_Indy'] = current_df['Agency Holding Co'].map(agency_dict)
last_week_df['Agency_Indy'] = last_week_df['Agency Holding Co'].map(agency_dict)
last_year_df['Agency_Indy'] = last_year_df['Agency Holding Co'].map(agency_dict)


# In[5]:


#Fill In Zeros for Finances

new_rev_cols = ['Q1 Total', 'Q1 Local', 'Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q2 Total', 'Q2 Local', 'Q2 Multi',
                'Q2 CTV','Q2 PG', 'Q3 Total', 'Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Q4 Total', 'Q4 Local',
                'Q4 Multi', 'Q4 CTV', 'Q4 PG']
for cols in new_rev_cols:
    current_df[cols].fillna(value=0, inplace=True)

old_rev_cols = ['LW Q1 Total', 'LW Q1 Local', 'LW Q1 Multi', 'LW Q1 CTV', 'LW Q1 PG', 'LW Q2 Total', 'LW Q2 Local',
                'LW Q2 Multi', 'LW Q2 CTV','LW Q2 PG', 'LW Q3 Total', 'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV',
                'LW Q3 PG', 'LW Q4 Total', 'LW Q4 Local', 'LW Q4 Multi','LW Q4 CTV', 'LW Q4 PG']
for cols in old_rev_cols:
    last_week_df[cols].fillna(value=0, inplace=True)

yoy_rev_cols = ['2022 Q1','2022 Q1 Local','2022 Q1 Multi', '2022 Q1 CTV', '2022 Q1 PG', '2022 Q2',
                '2022 Q2 Local', '2022 Q2 Multi','2022 Q2 CTV', '2022 Q2 PG', '2022 Q3', '2022 Q3 Local',
                '2022 Q3 Multi', '2022 Q3 CTV', '2022 Q3 PG', '2022 Q4', '2022 Q4 Local', '2022 Q4 Multi',
                '2022 Q4 CTV', '2022 Q4 PG', '2022 Total', '2022 Total Local', '2022 Total Multi',
                '2022 Total CTV', '2022 Total PG']
for cols in yoy_rev_cols:
    last_year_df[cols].fillna(value=0, inplace=True)
    


# In[6]:


#Removing those unwanted characters
unwanted_chars = ['~', ' - Global', ' - GLOBAL', ' - US', 'Australia', '(blank)']

for char in unwanted_chars:
    current_df = current_df.replace(char, '', regex=True)
    last_week_df = last_week_df.replace(char, '', regex=True)
    last_year_df = last_year_df.replace(char, '', regex=True)


# In[7]:


#Source Excel Template
src = file_name

wb = xw.Book(src)


# In[8]:


agency_list = ['Publicis','Omnicom','WPP','Dentsu','Interpublic','Horizon','Indy Client']


# In[9]:


for holding_co in agency_list:
    new_df = current_df.loc[current_df['Agency_Indy'] == holding_co]
    old_df = last_week_df.loc[last_week_df['Agency_Indy'] == holding_co]
    yoy_df = last_year_df.loc[last_year_df['Agency_Indy'] == holding_co]

    #GROUPING THIS WEEKS DATA
    new_child_df = new_df.groupby(['Advertiser']).agg({'Q4 Total':'sum', 'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Advertiser'],ascending=True).reset_index()
    new_parent_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 Total':'sum', 'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_q1_df = new_df.groupby(['Ultimate Parent']).agg({'Q1 Total':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_q2_df = new_df.groupby(['Ultimate Parent']).agg({'Q2 Total':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_q3_df = new_df.groupby(['Ultimate Parent']).agg({'Q3 Total':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_q4_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 Total':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_industry_df = new_df.groupby(['Account Category']).agg({'Q4 Total':'sum', 'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Account Category'],ascending=True).reset_index()
    new_parent_ctv_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 CTV':'sum', 'Q1 CTV':'sum', 'Q2 CTV':'sum', 'Q3 CTV':'sum', 'Total CTV':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_local_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 Local':'sum', 'Q1 Local':'sum', 'Q2 Local':'sum', 'Q3 Local':'sum', 'Total Local':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_multi_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 Multi':'sum', 'Q1 Multi':'sum', 'Q2 Multi':'sum', 'Q3 Multi':'sum', 'Total Multi':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_parent_pg_df = new_df.groupby(['Ultimate Parent']).agg({'Q4 PG':'sum', 'Q1 PG':'sum', 'Q2 PG':'sum', 'Q3 PG':'sum', 'Total PG':'sum'}).sort_values(by=['Ultimate Parent'],ascending=True).reset_index()
    new_agency_df = new_df.groupby(['Agency']).agg({'Q4 Total':'sum', 'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Agency'],ascending=True).reset_index()
    new_seller_df = new_df.groupby(['Teammember Name']).agg({'Q4 Total':'sum', 'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Teammember Name'],ascending=True).reset_index()


    #GROUPING LAST WEEKS DATA
    old_child_df = old_df.groupby(['Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
    old_parent_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
    old_parent_q1_df = old_df.groupby(['Ultimate Parent']).agg({'LW Q1 Total':'sum'}).sort_values(by=['LW Q1 Total'],ascending=False).reset_index()
    old_parent_q2_df = old_df.groupby(['Ultimate Parent']).agg({'LW Q2 Total':'sum'}).sort_values(by=['LW Q2 Total'],ascending=False).reset_index()
    old_parent_q3_df = old_df.groupby(['Ultimate Parent']).agg({'LW Q3 Total':'sum'}).sort_values(by=['LW Q3 Total'],ascending=False).reset_index()
    old_parent_q4_df = old_df.groupby(['Ultimate Parent']).agg({'LW Q4 Total':'sum'}).sort_values(by=['LW Q4 Total'],ascending=False).reset_index()
    old_industry_df = old_df.groupby(['Account Category']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
    old_parent_ctv_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total CTV':'sum'}).sort_values(by=['LW Total CTV'],ascending=False).reset_index()
    old_parent_local_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total Local':'sum'}).sort_values(by=['LW Total Local'],ascending=False).reset_index()
    old_parent_multi_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total Multi':'sum'}).sort_values(by=['LW Total Multi'],ascending=False).reset_index()
    old_parent_pg_df = old_df.groupby(['Ultimate Parent']).agg({'LW Total PG':'sum'}).sort_values(by=['LW Total PG'],ascending=False).reset_index()
    old_agency_df = old_df.groupby(['Agency']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()
    old_seller_df = old_df.groupby(['Teammember Name']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'],ascending=False).reset_index()


    #GROUPING LAST YEARS DATA
    yoy_child_df = yoy_df.groupby(['Advertiser']).agg({'2022 Total':'sum'}).sort_values(by=['2022 Total'],ascending=False).reset_index()
    yoy_parent_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Total':'sum'}).sort_values(by=['2022 Total'],ascending=False).reset_index()
    yoy_q1_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Q1':'sum'}).sort_values(by=['2022 Q1'],ascending=False).reset_index()
    yoy_q2_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Q2':'sum'}).sort_values(by=['2022 Q2'],ascending=False).reset_index()
    yoy_q3_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Q3':'sum'}).sort_values(by=['2022 Q3'],ascending=False).reset_index()
    yoy_q4_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Q4':'sum'}).sort_values(by=['2022 Q4'],ascending=False).reset_index()
    yoy_industry_df = yoy_df.groupby(['Account Category']).agg({'2022 Total':'sum'}).sort_values(by=['2022 Total'],ascending=False).reset_index()
    yoy_ctv_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Total CTV':'sum'}).sort_values(by=['2022 Total CTV'],ascending=False).reset_index()
    yoy_local_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Total Local':'sum'}).sort_values(by=['2022 Total Local'],ascending=False).reset_index()
    yoy_multi_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Total Multi':'sum'}).sort_values(by=['2022 Total Multi'],ascending=False).reset_index()
    yoy_pg_df = yoy_df.groupby(['Ultimate Parent']).agg({'2022 Total PG':'sum'}).sort_values(by=['2022 Total PG'],ascending=False).reset_index()
    yoy_agency_df = yoy_df.groupby(['Agency']).agg({'2022 Total':'sum'}).sort_values(by=['2022 Total'],ascending=False).reset_index()


    #----------- Merging This and Last Week Dataframes ------------#
    child_result = new_child_df.merge(old_child_df,how='outer', on='Advertiser').merge(yoy_child_df,how='outer', on='Advertiser')
    child_result.insert(7,'WoW Change','')
    child_result['WoW Change'] = child_result['Total Amount Net'] - child_result['LW Total Amount Net']
    child_result.insert(8,'WoW Change %','')
    child_result['WoW Change %'] = (child_result['Total Amount Net'] - child_result['LW Total Amount Net'])/child_result['LW Total Amount Net']
    child_result['YoY Change %'] = (child_result['Total Amount Net'] - child_result['2022 Total'])/child_result['2022 Total']

    parent_result = new_parent_df.merge(old_parent_df,how='outer', on='Ultimate Parent').merge(yoy_parent_df,how='outer', on='Ultimate Parent')
    parent_result.insert(7,'WoW Change','')
    parent_result['WoW Change'] = parent_result['Total Amount Net'] - parent_result['LW Total Amount Net']
    parent_result.insert(8,'WoW Change %','')
    parent_result['WoW Change %'] = (parent_result['Total Amount Net'] - parent_result['LW Total Amount Net'])/parent_result['LW Total Amount Net']
    parent_result['YoY Change %'] = (parent_result['Total Amount Net'] - parent_result['2022 Total'])/parent_result['2022 Total']

    parent_q1_result = new_parent_q1_df.merge(old_parent_q1_df,how='outer', on='Ultimate Parent').merge(yoy_q1_df,how='outer', on='Ultimate Parent')
    parent_q1_result.insert(3,'WoW Change','')
    parent_q1_result['WoW Change'] = parent_q1_result['Q1 Total'] - parent_q1_result['LW Q1 Total']
    parent_q1_result.insert(4,'WoW Change %','')
    parent_q1_result['WoW Change %'] = (parent_q1_result['Q1 Total'] - parent_q1_result['LW Q1 Total'])/parent_q1_result['LW Q1 Total']
    parent_q1_result['YoY Change %'] = (parent_q1_result['Q1 Total'] - parent_q1_result['2022 Q1'])/parent_q1_result['2022 Q1']

    parent_q2_result = new_parent_q2_df.merge(old_parent_q2_df,how='outer', on='Ultimate Parent').merge(yoy_q2_df,how='outer', on='Ultimate Parent')
    parent_q2_result.insert(3,'WoW Change','')
    parent_q2_result['WoW Change'] = parent_q2_result['Q2 Total'] - parent_q2_result['LW Q2 Total']
    parent_q2_result.insert(4,'WoW Change %','')
    parent_q2_result['WoW Change %'] = (parent_q2_result['Q2 Total'] - parent_q2_result['LW Q2 Total'])/parent_q2_result['LW Q2 Total']
    parent_q2_result['YoY Change %'] = (parent_q2_result['Q2 Total'] - parent_q2_result['2022 Q2'])/parent_q2_result['2022 Q2']

    parent_q3_result = new_parent_q3_df.merge(old_parent_q3_df,how='outer', on='Ultimate Parent').merge(yoy_q3_df,how='outer', on='Ultimate Parent')
    parent_q3_result.insert(3,'WoW Change','')
    parent_q3_result['WoW Change'] = parent_q3_result['Q3 Total'] - parent_q3_result['LW Q3 Total']
    parent_q3_result.insert(4,'WoW Change %','')
    parent_q3_result['WoW Change %'] = (parent_q3_result['Q3 Total'] - parent_q3_result['LW Q3 Total'])/parent_q3_result['LW Q3 Total']
    parent_q3_result['YoY Change %'] = (parent_q3_result['Q3 Total'] - parent_q3_result['2022 Q3'])/parent_q3_result['2022 Q3']

    parent_q4_result = new_parent_q4_df.merge(old_parent_q4_df,how='outer', on='Ultimate Parent').merge(yoy_q4_df,how='outer', on='Ultimate Parent')
    parent_q4_result.insert(3,'WoW Change','')
    parent_q4_result['WoW Change'] = parent_q4_result['Q4 Total'] - parent_q4_result['LW Q4 Total']
    parent_q4_result.insert(4,'WoW Change %','')
    parent_q4_result['WoW Change %'] = (parent_q4_result['Q4 Total'] - parent_q4_result['LW Q4 Total'])/parent_q4_result['LW Q4 Total']
    parent_q4_result['YoY Change %'] = (parent_q4_result['Q4 Total'] - parent_q4_result['2022 Q4'])/parent_q4_result['2022 Q4']

    industry_result = new_industry_df.merge(old_industry_df,how='outer', on='Account Category').merge(yoy_industry_df,how='outer', on='Account Category')
    industry_result.insert(7,'WoW Change','')
    industry_result['WoW Change'] = industry_result['Total Amount Net'] - industry_result['LW Total Amount Net']
    industry_result.insert(8,'WoW Change %','')
    industry_result['WoW Change %'] = (industry_result['Total Amount Net'] - industry_result['LW Total Amount Net'])/industry_result['LW Total Amount Net']
    industry_result['YoY Change %'] = (industry_result['Total Amount Net'] - industry_result['2022 Total'])/industry_result['2022 Total']

    agency_result = new_agency_df.merge(old_agency_df,how='outer', on='Agency').merge(yoy_agency_df,how='outer', on='Agency')
    agency_result.insert(7,'WoW Change','')
    agency_result['WoW Change'] = agency_result['Total Amount Net'] - agency_result['LW Total Amount Net']
    agency_result.insert(8,'WoW Change %','')
    agency_result['WoW Change %'] = (agency_result['Total Amount Net'] - agency_result['LW Total Amount Net'])/agency_result['LW Total Amount Net']
    agency_result['YoY Change %'] = (agency_result['Total Amount Net'] - agency_result['2022 Total'])/agency_result['2022 Total']

    ctv_result = new_parent_ctv_df.merge(old_parent_ctv_df,how='outer', on='Ultimate Parent').merge(yoy_ctv_df,how='outer', on='Ultimate Parent')
    ctv_result.insert(7,'WoW Change','')
    ctv_result['WoW Change'] = ctv_result['Total CTV'] - ctv_result['LW Total CTV']
    ctv_result.insert(8,'WoW Change %','')
    ctv_result['WoW Change %'] = (ctv_result['Total CTV'] - ctv_result['LW Total CTV'])/ctv_result['LW Total CTV']
    ctv_result['YoY Change %'] = (ctv_result['Total CTV'] - ctv_result['2022 Total CTV'])/ctv_result['2022 Total CTV']

    local_result = new_parent_local_df.merge(old_parent_local_df,how='outer', on='Ultimate Parent').merge(yoy_local_df,how='outer', on='Ultimate Parent')
    local_result.insert(7,'WoW Change','')
    local_result['WoW Change'] = local_result['Total Local'] - local_result['LW Total Local']
    local_result.insert(8,'WoW Change %','')
    local_result['WoW Change %'] = (local_result['Total Local'] - local_result['LW Total Local'])/local_result['LW Total Local']
    local_result['YoY Change %'] = (local_result['Total Local'] - local_result['2022 Total Local'])/local_result['2022 Total Local']

    multi_result = new_parent_multi_df.merge(old_parent_multi_df,how='outer', on='Ultimate Parent').merge(yoy_multi_df,how='outer', on='Ultimate Parent')
    multi_result.insert(7,'WoW Change','')
    multi_result['WoW Change'] = multi_result['Total Multi'] - multi_result['LW Total Multi']
    multi_result.insert(8,'WoW Change %','')
    multi_result['WoW Change %'] = (multi_result['Total Multi'] - multi_result['LW Total Multi'])/multi_result['LW Total Multi']
    multi_result['YoY Change %'] = (multi_result['Total Multi'] - multi_result['2022 Total Multi'])/multi_result['2022 Total Multi']

    pg_result = new_parent_pg_df.merge(old_parent_pg_df,how='outer', on='Ultimate Parent').merge(yoy_pg_df,how='outer', on='Ultimate Parent')
    pg_result.insert(7,'WoW Change','')
    pg_result['WoW Change'] = pg_result['Total PG'] - pg_result['LW Total PG']
    pg_result.insert(8,'WoW Change %','')
    pg_result['WoW Change %'] = (pg_result['Total PG'] - pg_result['LW Total PG'])/pg_result['LW Total PG']
    pg_result['YoY Change %'] = (pg_result['Total PG'] - pg_result['2022 Total PG'])/pg_result['2022 Total PG']

    seller_result = new_seller_df.merge(old_seller_df,how='outer', on='Teammember Name')
    seller_result.insert(7,'WoW Change','')
    seller_result['WoW Change'] = seller_result['Total Amount Net'] - seller_result['LW Total Amount Net']
    seller_result.insert(8,'WoW Change %','')
    seller_result['WoW Change %'] = (seller_result['Total Amount Net'] - seller_result['LW Total Amount Net'])/seller_result['LW Total Amount Net']

    #Identifying Lost
    lost_dfs = [child_result,parent_result, industry_result,agency_result]

    for k in lost_dfs:
        mask1 = ( ((k['Total Amount Net'].isna()) & (k['2022 Total'] > 0)) | ((k['Total Amount Net']==0) & (k['2022 Total'] > 0)) )
        k.loc[mask1, 'WoW Change'] = 'Lost'

    #Fill In New and Lost Labels
    child_result['2022 Total'].fillna(value='New This Year', inplace=True)
    child_result.loc[child_result['YoY Change %'] == np.inf, '2022 Total'] = 'New This Year'

    parent_result['2022 Total'].fillna(value='New This Year', inplace=True)
    parent_result.loc[parent_result['YoY Change %'] == np.inf, '2022 Total'] = 'New This Year'

    industry_result['2022 Total'].fillna(value='New This Year', inplace=True)
    industry_result.loc[industry_result['YoY Change %'] == np.inf, '2022 Total'] = 'New This Year'

    agency_result['2022 Total'].fillna(value='New This Year', inplace=True)
    agency_result.loc[agency_result['YoY Change %'] == np.inf, '2022 Total'] = 'New This Year'

    lost_mask = ( ((ctv_result['Total CTV'].isna()) & (ctv_result['2022 Total CTV'] > 0)) | ((ctv_result['Total CTV']==0) & (ctv_result['2022 Total CTV'] > 0)) )
    ctv_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    ctv_result['2022 Total CTV'].fillna(value='New This Year', inplace=True)
    ctv_result.loc[ctv_result['YoY Change %'] == np.inf, '2022 Total CTV'] = 'New This Year'

    lost_mask = ( ((local_result['Total Local'].isna()) & (local_result['2022 Total Local'] > 0)) | ((local_result['Total Local']==0) & (local_result['2022 Total Local'] > 0)) )
    local_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    local_result['2022 Total Local'].fillna(value='New This Year', inplace=True)
    local_result.loc[local_result['YoY Change %'] == np.inf, '2022 Total Local'] = 'New This Year'

    lost_mask = ( ((multi_result['Total Multi'].isna()) & (multi_result['2022 Total Multi'] > 0)) | ((multi_result['Total Multi'] ==0 ) & (multi_result['2022 Total Multi'] > 0)))
    multi_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    multi_result['2022 Total Multi'].fillna(value='New This Year', inplace=True)
    multi_result.loc[multi_result['YoY Change %'] == np.inf, '2022 Total Multi'] = 'New This Year'

    lost_mask = ( ((pg_result['Total PG'].isna()) & (pg_result['2022 Total PG'] > 0)) | ((pg_result['Total PG']==0) & (pg_result['2022 Total PG'] > 0)) )
    pg_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    pg_result['2022 Total PG'].fillna(value='New This Year', inplace=True)
    pg_result.loc[pg_result['YoY Change %'] == np.inf, '2022 Total PG'] = 'New This Year'

    lost_mask = ( ((parent_q1_result['Q1 Total'].isna()) & (parent_q1_result['2022 Q1'] > 0)) | ((parent_q1_result['Q1 Total']==0) & (parent_q1_result['2022 Q1'] > 0)) )
    parent_q1_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q1_result['2022 Q1'].fillna(value='New This Year', inplace=True)
    parent_q1_result.loc[parent_q1_result['YoY Change %'] == np.inf, '2022 Q1'] = 'New This Year'

    lost_mask = ( ((parent_q2_result['Q2 Total'].isna()) & (parent_q2_result['2022 Q2'] > 0)) | ((parent_q2_result['Q2 Total']==0) & (parent_q2_result['2022 Q2'] > 0)) )
    parent_q2_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q2_result['2022 Q2'].fillna(value='New This Year', inplace=True)
    parent_q2_result.loc[parent_q2_result['YoY Change %'] == np.inf, '2022 Q2'] = 'New This Year'

    lost_mask = ( ((parent_q3_result['Q3 Total'].isna()) & (parent_q3_result['2022 Q3'] > 0)) | ((parent_q3_result['Q3 Total']==0) & (parent_q3_result['2022 Q3'] > 0)) )
    parent_q3_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q3_result['2022 Q3'].fillna(value='New This Year', inplace=True)
    parent_q3_result.loc[parent_q3_result['YoY Change %'] == np.inf, '2022 Q3'] = 'New This Year'

    lost_mask = ( ((parent_q4_result['Q4 Total'].isna()) & (parent_q4_result['2022 Q4'] > 0)) | ((parent_q4_result['Q4 Total']==0) & (parent_q4_result['2022 Q4'] > 0)) )
    parent_q4_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q4_result['2022 Q4'].fillna(value='New This Year', inplace=True)
    parent_q4_result.loc[parent_q4_result['YoY Change %'] == np.inf, '2022 Q4'] = 'New This Year'

    new_week_cols = [child_result,parent_result,parent_q1_result,parent_q2_result,parent_q3_result,parent_q4_result,
                    industry_result,agency_result,ctv_result,multi_result,local_result,pg_result]
    for j in new_week_cols:
        j.loc[j['WoW Change %'] == np.inf, 'WoW Change'] = 'New This Week'
        j['WoW Change'].fillna(value='New This Week', inplace=True)

    result_df = [child_result, parent_result, parent_q1_result, parent_q2_result, parent_q3_result, parent_q4_result, 
                 industry_result, ctv_result, local_result, multi_result, pg_result,agency_result, seller_result]

    for dfs in result_df:
        dfs.replace([np.inf, -np.inf], np.nan, inplace=True)
        dfs.loc[dfs['WoW Change'] == 'Lost', 'YoY Change %'] = -1

    #DROP THE LAST WEEK TOTAL COLUMN
    child_final = child_result.drop('LW Total Amount Net', axis=1)
    parent_final = parent_result.drop('LW Total Amount Net', axis=1)
    q1_final = parent_q1_result.drop('LW Q1 Total', axis=1)
    q2_final = parent_q2_result.drop('LW Q2 Total', axis=1)
    q3_final = parent_q3_result.drop('LW Q3 Total', axis=1)
    q4_final = parent_q4_result.drop('LW Q4 Total', axis=1)
    industry_final = industry_result.drop('LW Total Amount Net', axis=1)
    agency_final = agency_result.drop('LW Total Amount Net', axis=1)
    ctv_final = ctv_result.drop('LW Total CTV', axis=1)
    local_final = local_result.drop('LW Total Local', axis=1)
    multi_final = multi_result.drop('LW Total Multi', axis=1)
    pg_final = pg_result.drop('LW Total PG', axis=1)
    seller_final = seller_result.drop('LW Total Amount Net', axis=1)
    
    child_final = child_final[(child_final['Total Amount Net'] > 0) | (child_final['WoW Change'] == 'Lost')] 
    parent_final = parent_final[(parent_final['Total Amount Net'] > 0) | (parent_final['WoW Change'] == 'Lost')] 
    industry_final = industry_final[(industry_final['Total Amount Net'] > 0) | (industry_final['WoW Change'] == 'Lost')]
    agency_final = agency_final[(agency_final['Total Amount Net'] > 0) | (agency_final['WoW Change'] == 'Lost')] 
    q1_final = q1_final[(q1_final['Q1 Total'] > 0) | (q1_final['WoW Change'] == 'Lost')] 
    q2_final = q2_final[(q2_final['Q2 Total'] > 0) | (q2_final['WoW Change'] == 'Lost')] 
    q3_final = q3_final[(q3_final['Q3 Total'] > 0) | (q3_final['WoW Change'] == 'Lost')] 
    q4_final = q4_final[(q4_final['Q4 Total'] > 0) | (q4_final['WoW Change'] == 'Lost')]
    ctv_final = ctv_final[(ctv_final['Total CTV'] > 0) | (ctv_final['WoW Change'] == 'Lost')] 
    local_final = local_final[(local_final['Total Local'] > 0) | (local_final['WoW Change'] == 'Lost')] 
    multi_final = multi_final[(multi_final['Total Multi'] > 0) | (multi_final['WoW Change'] == 'Lost')] 
    pg_final = pg_final[(pg_final['Total PG'] > 0) | (pg_final['WoW Change'] == 'Lost')] 

    
    final_df_list = [child_final,parent_final,industry_final,q1_final,q2_final,q3_final,q4_final,ctv_final,local_final,multi_final,pg_final]
    for finals in final_df_list:
        finals.sort_values(by=[finals.columns[0]],ascending=True).reset_index()
    
    
    #EXPORT SECTION
    #Front End Tab Export
    wb.sheets[holding_co].range('L9').options(index=False,header=False).value = agency_final
    wb.sheets[holding_co].range('W9').options(index=False,header=False).value = child_final
    wb.sheets[holding_co].range('AH9').options(index=False,header=False).value = parent_final
    wb.sheets[holding_co].range('AS9').options(index=False,header=False).value = q4_final
    wb.sheets[holding_co].range('AZ9').options(index=False,header=False).value = q1_final
    wb.sheets[holding_co].range('BG9').options(index=False,header=False).value = q2_final
    wb.sheets[holding_co].range('BN9').options(index=False,header=False).value = q3_final
    wb.sheets[holding_co].range('BU9').options(index=False,header=False).value = industry_final
    wb.sheets[holding_co].range('CF9').options(index=False,header=False).value = seller_final
    wb.sheets[holding_co].range('CO9').options(index=False,header=False).value = ctv_final
    wb.sheets[holding_co].range('CZ9').options(index=False,header=False).value = local_final
    wb.sheets[holding_co].range('DK9').options(index=False,header=False).value = multi_final
    wb.sheets[holding_co].range('DV9').options(index=False,header=False).value = pg_final


    
    #ADD THE RAW DATAFRAMES AS WELL
    wb.sheets[f'{holding_co} Raw'].range('A2').options(index=False,header=True).value = agency_result
    wb.sheets[f'{holding_co} Raw'].range('M2').options(index=False,header=True).value = child_result
    wb.sheets[f'{holding_co} Raw'].range('Y2').options(index=False,header=True).value = parent_result
    wb.sheets[f'{holding_co} Raw'].range('AK2').options(index=False,header=True).value = parent_q4_result
    wb.sheets[f'{holding_co} Raw'].range('AS2').options(index=False,header=True).value = parent_q1_result
    wb.sheets[f'{holding_co} Raw'].range('BA2').options(index=False,header=True).value = parent_q2_result
    wb.sheets[f'{holding_co} Raw'].range('BI2').options(index=False,header=True).value = parent_q3_result
    wb.sheets[f'{holding_co} Raw'].range('BQ2').options(index=False,header=True).value = industry_result
    wb.sheets[f'{holding_co} Raw'].range('CC2').options(index=False,header=True).value = seller_result
    wb.sheets[f'{holding_co} Raw'].range('CM2').options(index=False,header=True).value = ctv_result
    wb.sheets[f'{holding_co} Raw'].range('CY2').options(index=False,header=True).value = local_result
    wb.sheets[f'{holding_co} Raw'].range('DK2').options(index=False,header=True).value = multi_result
    wb.sheets[f'{holding_co} Raw'].range('DW2').options(index=False,header=True).value = pg_result



    #new_df.drop(new_df.index, inplace=True)
    #old_df.drop(old_df.index, inplace=True)


# In[10]:


dest = 'C:/Users/Alexander Ravazzoni/Documents/Agency Lookback/Agency Lookback Template.xlsx'

wb.save(dest)


# In[11]:


print("Your report is located here -> "+dest)


# In[12]:


child_final.shape[1]


# In[13]:


child_final.columns[0]


# In[14]:


new_df.dtypes


# In[15]:


print(pd. __version__)


# In[ ]:




