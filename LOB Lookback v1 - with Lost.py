#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import xlwings as xw
import openpyxl
from datetime import date
import numpy as np

today_date = '2023-09-01'


# In[2]:


#-------------------  Reading The Excel Spreadsheets   -------------------#
file_name = 'C:/Users/Alexander Ravazzoni/Documents/LOB Lookback/LOB Lookback Template.xlsx'

#This Weeks Data
new_df = pd.read_excel(file_name, sheet_name='2023 New Data')
#Last Weeks Data
old_df = pd.read_excel(file_name, sheet_name='2023 Old Data')
#Last Years Data
yoy_df = pd.read_excel(file_name, sheet_name='2022 Data')


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

last_year_cols = ['Team Name', 'Deal Name', 'Advertiser', 'Account Category','Account Subcategory',
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent','2022 Q1','2022 Q1 Local',
                  '2022 Q1 Multi', '2022 Q1 CTV', '2022 Q1 PG', '2022 Q2', '2022 Q2 Local', '2022 Q2 Multi',
                  '2022 Q2 CTV', '2022 Q2 PG', '2022 Q3', '2022 Q3 Local', '2022 Q3 Multi', '2022 Q3 CTV', '2022 Q3 PG',
                  '2022 Q4', '2022 Q4 Local', '2022 Q4 Multi', '2022 Q4 CTV', '2022 Q4 PG',
                  '2022 Total',	'2022 Total Local', '2022 Total Multi', '2022 Total CTV', '2022 Total PG']

yoy_df.columns = last_year_cols


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

yoy_rev_cols = ['2022 Q1','2022 Q1 Local','2022 Q1 Multi', '2022 Q1 CTV', '2022 Q1 PG', '2022 Q2',
                '2022 Q2 Local', '2022 Q2 Multi','2022 Q2 CTV', '2022 Q2 PG', '2022 Q3', '2022 Q3 Local',
                '2022 Q3 Multi', '2022 Q3 CTV', '2022 Q3 PG', '2022 Q4', '2022 Q4 Local', '2022 Q4 Multi',
                '2022 Q4 CTV', '2022 Q4 PG', '2022 Total', '2022 Total Local', '2022 Total Multi',
                '2022 Total CTV', '2022 Total PG']
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


#EXPORT SECTION
src = file_name
wb = xw.Book(src)


# In[8]:


LOB_List = ['CTV','Local','Multi','PG']

for i in LOB_List:
    Q1_LOB = f'Q1 {i}'
    Q2_LOB = f'Q2 {i}'
    Q3_LOB = f'Q3 {i}'
    Q4_LOB = f'Q4 {i}'
    Total_LOB = f'Total {i}'
    LW_Q1_LOB = f'LW Q1 {i}'
    LW_Q2_LOB = f'LW Q2 {i}'
    LW_Q3_LOB = f'LW Q3 {i}'
    LW_Q4_LOB = f'LW Q4 {i}'
    LW_Total_LOB = f'LW Total {i}'
    YoY_Q1_LOB = f'2022 Q1 {i}'
    YoY_Q2_LOB = f'2022 Q2 {i}'
    YoY_Q3_LOB = f'2022 Q3 {i}'
    YoY_Q4_LOB = f'2022 Q4 {i}'
    YoY_Total_LOB = f'2022 Total {i}'

    #GROUPING THIS WEEKS DATA
    new_child_df = new_df.groupby(['Advertiser']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=[Total_LOB],ascending=False).reset_index()
    new_parent_df = new_df.groupby(['Ultimate Parent']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=[Total_LOB],ascending=False).reset_index()
    new_parent_q1_df = new_df.groupby(['Ultimate Parent']).agg({Q1_LOB:'sum'}).sort_values(by=[Q1_LOB],ascending=False).reset_index()
    new_parent_q2_df = new_df.groupby(['Ultimate Parent']).agg({Q2_LOB:'sum'}).sort_values(by=[Q2_LOB],ascending=False).reset_index()
    new_parent_q3_df = new_df.groupby(['Ultimate Parent']).agg({Q3_LOB:'sum'}).sort_values(by=[Q3_LOB],ascending=False).reset_index()
    new_parent_q4_df = new_df.groupby(['Ultimate Parent']).agg({Q4_LOB:'sum'}).sort_values(by=[Q4_LOB],ascending=False).reset_index()
    new_industry_df = new_df.groupby(['Account Category']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=[Total_LOB],ascending=False).reset_index()
    new_agency_indy_df = new_df.groupby(['Agency_Indy']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
    new_indy_df = new_df.loc[new_df['Agency_Indy'] == 'Indy']
    new_indy_df = new_indy_df.groupby(['Agency']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=[Total_LOB],ascending=False).reset_index()
    new_agency_df = new_df.loc[new_df['Agency_Indy'] == 'Agency']
    new_agency_holding_df = new_agency_df.groupby(['Agency Holding Co']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=[Total_LOB],ascending=False).reset_index()
    new_agency_df = new_agency_df.groupby(['Agency Holding Co','Agency']).agg({Q1_LOB:'sum', Q2_LOB:'sum', Q3_LOB:'sum', Q4_LOB:'sum', Total_LOB:'sum'}).sort_values(by=['Agency Holding Co',Total_LOB],ascending=False).reset_index()


    #GROUPING LAST WEEKS DATA
    old_child_df = old_df.groupby(['Advertiser']).agg({LW_Total_LOB:'sum'}).sort_values(by=[LW_Total_LOB],ascending=False).reset_index()
    old_parent_df = old_df.groupby(['Ultimate Parent']).agg({LW_Total_LOB:'sum'}).sort_values(by=[LW_Total_LOB],ascending=False).reset_index()
    old_parent_q1_df = old_df.groupby(['Ultimate Parent']).agg({LW_Q1_LOB:'sum'}).sort_values(by=[LW_Q1_LOB],ascending=False).reset_index()
    old_parent_q2_df = old_df.groupby(['Ultimate Parent']).agg({LW_Q2_LOB:'sum'}).sort_values(by=[LW_Q2_LOB],ascending=False).reset_index()
    old_parent_q3_df = old_df.groupby(['Ultimate Parent']).agg({LW_Q3_LOB:'sum'}).sort_values(by=[LW_Q3_LOB],ascending=False).reset_index()
    old_parent_q4_df = old_df.groupby(['Ultimate Parent']).agg({LW_Q4_LOB:'sum'}).sort_values(by=[LW_Q4_LOB],ascending=False).reset_index()
    old_industry_df = old_df.groupby(['Account Category']).agg({LW_Total_LOB:'sum'}).sort_values(by=[LW_Total_LOB],ascending=False).reset_index()
    old_agency_holding_df = old_df.groupby(['Agency Holding Co']).agg({LW_Total_LOB:'sum'}).sort_values(by=[LW_Total_LOB],ascending=False).reset_index()
    old_agency_indy_df = old_df.groupby(['Agency_Indy']).agg({LW_Total_LOB:'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
    old_indy_df = old_df.loc[old_df['Agency_Indy'] == 'Indy']
    old_indy_df = old_indy_df.groupby(['Agency']).agg({LW_Total_LOB:'sum'}).sort_values(by=[LW_Total_LOB],ascending=False).reset_index()
    old_agency_df = old_df.loc[old_df['Agency_Indy'] == 'Agency']
    old_agency_df = old_agency_df.groupby(['Agency Holding Co','Agency']).agg({LW_Total_LOB:'sum'}).sort_values(by=['Agency Holding Co',LW_Total_LOB],ascending=False).reset_index()


    #GROUPING LAST YEARS DATA
    yoy_child_df = yoy_df.groupby(['Advertiser']).agg({YoY_Total_LOB:'sum'}).sort_values(by=[YoY_Total_LOB],ascending=False).reset_index()
    yoy_parent_df = yoy_df.groupby(['Ultimate Parent']).agg({YoY_Total_LOB:'sum'}).sort_values(by=[YoY_Total_LOB],ascending=False).reset_index()
    yoy_q1_df = yoy_df.groupby(['Ultimate Parent']).agg({YoY_Q1_LOB:'sum'}).sort_values(by=[YoY_Q1_LOB],ascending=False).reset_index()
    yoy_q2_df = yoy_df.groupby(['Ultimate Parent']).agg({YoY_Q2_LOB:'sum'}).sort_values(by=[YoY_Q2_LOB],ascending=False).reset_index()
    yoy_q3_df = yoy_df.groupby(['Ultimate Parent']).agg({YoY_Q3_LOB:'sum'}).sort_values(by=[YoY_Q3_LOB],ascending=False).reset_index()
    yoy_q4_df = yoy_df.groupby(['Ultimate Parent']).agg({YoY_Q4_LOB:'sum'}).sort_values(by=[YoY_Q4_LOB],ascending=False).reset_index()
    yoy_industry_df = yoy_df.groupby(['Account Category']).agg({YoY_Total_LOB:'sum'}).sort_values(by=[YoY_Total_LOB],ascending=False).reset_index()
    yoy_agency_holding_df = yoy_df.groupby(['Agency Holding Co']).agg({YoY_Total_LOB:'sum'}).sort_values(by=[YoY_Total_LOB],ascending=False).reset_index()
    yoy_agency_indy_df = yoy_df.groupby(['Agency_Indy']).agg({YoY_Total_LOB:'sum'}).sort_values(by=['Agency_Indy'],ascending=True).reset_index()
    yoy_indy_df = yoy_df.loc[yoy_df['Agency_Indy'] == 'Indy']
    yoy_indy_df = yoy_indy_df.groupby(['Agency']).agg({YoY_Total_LOB:'sum'}).sort_values(by=[YoY_Total_LOB],ascending=False).reset_index()
    yoy_agency_df = yoy_df.loc[yoy_df['Agency_Indy'] == 'Agency']
    yoy_agency_df = yoy_agency_df.groupby(['Agency Holding Co','Agency']).agg({YoY_Total_LOB:'sum'}).sort_values(by=['Agency Holding Co',YoY_Total_LOB],ascending=False).reset_index()

    #----------- Merging This and Last Week Dataframes ------------#
    child_result = new_child_df.merge(old_child_df,how='outer', on='Advertiser').merge(yoy_child_df,how='outer', on='Advertiser')
    child_result.insert(7,'WoW Change','')
    child_result['WoW Change'] = child_result[Total_LOB] - child_result[LW_Total_LOB]
    child_result.insert(8,'WoW Change %','')
    child_result['WoW Change %'] = (child_result[Total_LOB] - child_result[LW_Total_LOB])/child_result[LW_Total_LOB]
    child_result['YoY Change %'] = (child_result[Total_LOB] - child_result[YoY_Total_LOB])/child_result[YoY_Total_LOB]

    parent_result = new_parent_df.merge(old_parent_df,how='outer', on='Ultimate Parent').merge(yoy_parent_df,how='outer', on='Ultimate Parent')
    parent_result.insert(7,'WoW Change','')
    parent_result['WoW Change'] = parent_result[Total_LOB] - parent_result[LW_Total_LOB]
    parent_result.insert(8,'WoW Change %','')
    parent_result['WoW Change %'] = (parent_result[Total_LOB] - parent_result[LW_Total_LOB])/parent_result[LW_Total_LOB]
    parent_result['YoY Change %'] = (parent_result[Total_LOB] - parent_result[YoY_Total_LOB])/parent_result[YoY_Total_LOB]

    parent_q1_result = new_parent_q1_df.merge(old_parent_q1_df,how='outer', on='Ultimate Parent').merge(yoy_q1_df,how='outer', on='Ultimate Parent')
    parent_q1_result.insert(3,'WoW Change','')
    parent_q1_result['WoW Change'] = parent_q1_result[Q1_LOB] - parent_q1_result[LW_Q1_LOB]
    parent_q1_result.insert(4,'WoW Change %','')
    parent_q1_result['WoW Change %'] = (parent_q1_result[Q1_LOB] - parent_q1_result[LW_Q1_LOB])/parent_q1_result[LW_Q1_LOB]
    parent_q1_result['YoY Change %'] = (parent_q1_result[Q1_LOB] - parent_q1_result[YoY_Q1_LOB])/parent_q1_result[YoY_Q1_LOB]

    parent_q2_result = new_parent_q2_df.merge(old_parent_q2_df,how='outer', on='Ultimate Parent').merge(yoy_q2_df,how='outer', on='Ultimate Parent')
    parent_q2_result.insert(3,'WoW Change','')
    parent_q2_result['WoW Change'] = parent_q2_result[Q2_LOB] - parent_q2_result[LW_Q2_LOB]
    parent_q2_result.insert(4,'WoW Change %','')
    parent_q2_result['WoW Change %'] = (parent_q2_result[Q2_LOB] - parent_q2_result[LW_Q2_LOB])/parent_q2_result[LW_Q2_LOB]
    parent_q2_result['YoY Change %'] = (parent_q2_result[Q2_LOB] - parent_q2_result[YoY_Q2_LOB])/parent_q2_result[YoY_Q2_LOB]

    parent_q3_result = new_parent_q3_df.merge(old_parent_q3_df,how='outer', on='Ultimate Parent').merge(yoy_q3_df,how='outer', on='Ultimate Parent')
    parent_q3_result.insert(3,'WoW Change','')
    parent_q3_result['WoW Change'] = parent_q3_result[Q3_LOB] - parent_q3_result[LW_Q3_LOB]
    parent_q3_result.insert(4,'WoW Change %','')
    parent_q3_result['WoW Change %'] = (parent_q3_result[Q3_LOB] - parent_q3_result[LW_Q3_LOB])/parent_q3_result[LW_Q3_LOB]
    parent_q3_result['YoY Change %'] = (parent_q3_result[Q3_LOB] - parent_q3_result[YoY_Q3_LOB])/parent_q3_result[YoY_Q3_LOB]

    parent_q4_result = new_parent_q4_df.merge(old_parent_q4_df,how='outer', on='Ultimate Parent').merge(yoy_q4_df,how='outer', on='Ultimate Parent')
    parent_q4_result.insert(3,'WoW Change','')
    parent_q4_result['WoW Change'] = parent_q4_result[Q4_LOB] - parent_q4_result[LW_Q4_LOB]
    parent_q4_result.insert(4,'WoW Change %','')
    parent_q4_result['WoW Change %'] = (parent_q4_result[Q4_LOB] - parent_q4_result[LW_Q4_LOB])/parent_q4_result[LW_Q4_LOB]
    parent_q4_result['YoY Change %'] = (parent_q4_result[Q4_LOB] - parent_q4_result[YoY_Q4_LOB])/parent_q4_result[YoY_Q4_LOB]

    industry_result = new_industry_df.merge(old_industry_df,how='outer', on='Account Category').merge(yoy_industry_df,how='outer', on='Account Category')
    industry_result.insert(7,'WoW Change','')
    industry_result['WoW Change'] = industry_result[Total_LOB] - industry_result[LW_Total_LOB]
    industry_result.insert(8,'WoW Change %','')
    industry_result['WoW Change %'] = (industry_result[Total_LOB] - industry_result[LW_Total_LOB])/industry_result[LW_Total_LOB]
    industry_result['YoY Change %'] = (industry_result[Total_LOB] - industry_result[YoY_Total_LOB])/industry_result[YoY_Total_LOB]

    agency_holding_result = new_agency_holding_df.merge(old_agency_holding_df,how='outer', on='Agency Holding Co').merge(yoy_agency_holding_df,how='outer', on='Agency Holding Co')
    agency_holding_result.insert(7,'WoW Change','')
    agency_holding_result['WoW Change'] = agency_holding_result[Total_LOB] - agency_holding_result[LW_Total_LOB]
    agency_holding_result.insert(8,'WoW Change %','')
    agency_holding_result['WoW Change %'] = (agency_holding_result[Total_LOB] - agency_holding_result[LW_Total_LOB])/agency_holding_result[LW_Total_LOB]
    agency_holding_result['YoY Change %'] = (agency_holding_result[Total_LOB] - agency_holding_result[YoY_Total_LOB])/agency_holding_result[YoY_Total_LOB]

    agency_indy_result = new_agency_indy_df.merge(old_agency_indy_df,how='outer', on='Agency_Indy').merge(yoy_agency_indy_df,how='outer', on='Agency_Indy')
    agency_indy_result.insert(7,'WoW Change','')
    agency_indy_result['WoW Change'] = agency_indy_result[Total_LOB] - agency_indy_result[LW_Total_LOB]
    agency_indy_result.insert(8,'WoW Change %','')
    agency_indy_result['WoW Change %'] = (agency_indy_result[Total_LOB] - agency_indy_result[LW_Total_LOB])/agency_indy_result[LW_Total_LOB]
    agency_indy_result['YoY Change %'] = (agency_indy_result[Total_LOB] - agency_indy_result[YoY_Total_LOB])/agency_indy_result[YoY_Total_LOB]

    indy_result = new_indy_df.merge(old_indy_df,how='outer', on='Agency').merge(yoy_indy_df,how='outer', on='Agency')
    indy_result.insert(7,'WoW Change','')
    indy_result['WoW Change'] = indy_result[Total_LOB] - indy_result[LW_Total_LOB]
    indy_result.insert(8,'WoW Change %','')
    indy_result['WoW Change %'] = (indy_result[Total_LOB] - indy_result[LW_Total_LOB])/indy_result[LW_Total_LOB]
    indy_result['YoY Change %'] = (indy_result[Total_LOB] - indy_result[YoY_Total_LOB])/indy_result[YoY_Total_LOB]

    agency_result = new_agency_df.merge(old_agency_df,how='outer', on=['Agency Holding Co','Agency']).merge(yoy_agency_df,how='outer', on=['Agency Holding Co','Agency'])
    agency_result.insert(8,'WoW Change','')
    agency_result['WoW Change'] = agency_result[Total_LOB] - agency_result[LW_Total_LOB]
    agency_result.insert(9,'WoW Change %','')
    agency_result['WoW Change %'] = (agency_result[Total_LOB] - agency_result[LW_Total_LOB])/agency_result[LW_Total_LOB]
    agency_result['YoY Change %'] = (agency_result[Total_LOB] - agency_result[YoY_Total_LOB])/agency_result[YoY_Total_LOB]

    
    #Idenitfying New and Lost
    lost_mask = ( ((child_result[Total_LOB].isna()) & (child_result[YoY_Total_LOB] > 0)) | ((child_result[Total_LOB]==0) & (child_result[YoY_Total_LOB] > 0)) )
    child_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    child_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    child_result.loc[child_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'
    
    lost_mask = ( ((parent_result[Total_LOB].isna()) & (parent_result[YoY_Total_LOB] > 0)) | ((parent_result[Total_LOB]==0) & (parent_result[YoY_Total_LOB] > 0)) )
    parent_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    parent_result.loc[parent_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'

    lost_mask = ( ((industry_result[Total_LOB].isna()) & (industry_result[YoY_Total_LOB] > 0)) | ((industry_result[Total_LOB]==0) & (industry_result[YoY_Total_LOB] > 0)) )
    industry_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    industry_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    industry_result.loc[industry_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'

    lost_mask = ( ((agency_holding_result[Total_LOB].isna()) & (agency_holding_result[YoY_Total_LOB] > 0)) | ((agency_holding_result[Total_LOB]==0) & (agency_holding_result[YoY_Total_LOB] > 0)) )
    agency_holding_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    agency_holding_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    agency_holding_result.loc[agency_holding_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'

    agency_indy_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    
    lost_mask = ( ((indy_result[Total_LOB].isna()) & (indy_result[YoY_Total_LOB] > 0)) | ((indy_result[Total_LOB]==0) & (indy_result[YoY_Total_LOB] > 0)) )
    indy_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    indy_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    indy_result.loc[indy_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'

    lost_mask = ( ((agency_result[Total_LOB].isna()) & (agency_result[YoY_Total_LOB] > 0)) | ((agency_result[Total_LOB]==0) & (agency_result[YoY_Total_LOB] > 0)) )
    agency_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    agency_result[YoY_Total_LOB].fillna(value='New This Year', inplace=True)
    agency_result.loc[agency_result['YoY Change %'] == np.inf, YoY_Total_LOB] = 'New This Year'
    
    lost_mask = ( ((parent_q1_result[Q1_LOB].isna()) & (parent_q1_result[YoY_Q1_LOB] > 0)) | ((parent_q1_result[Q1_LOB]==0) & (parent_q1_result[YoY_Q1_LOB] > 0)) )
    parent_q1_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q1_result[YoY_Q1_LOB].fillna(value='New This Year', inplace=True)
    parent_q1_result.loc[parent_q1_result['YoY Change %'] == np.inf, YoY_Q1_LOB] = 'New This Year'
    
    lost_mask = ( ((parent_q2_result[Q2_LOB].isna()) & (parent_q2_result[YoY_Q2_LOB] > 0)) | ((parent_q2_result[Q2_LOB]==0) & (parent_q2_result[YoY_Q2_LOB] > 0)) )
    parent_q2_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q2_result[YoY_Q2_LOB].fillna(value='New This Year', inplace=True)
    parent_q2_result.loc[parent_q2_result['YoY Change %'] == np.inf, YoY_Q2_LOB] = 'New This Year'

    lost_mask = ( ((parent_q3_result[Q3_LOB].isna()) & (parent_q3_result[YoY_Q3_LOB] > 0)) | ((parent_q3_result[Q3_LOB]==0) & (parent_q3_result[YoY_Q3_LOB] > 0)) )
    parent_q3_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q3_result[YoY_Q3_LOB].fillna(value='New This Year', inplace=True)
    parent_q3_result.loc[parent_q3_result['YoY Change %'] == np.inf, YoY_Q3_LOB] = 'New This Year'
    
    lost_mask = ( ((parent_q4_result[Q4_LOB].isna()) & (parent_q4_result[YoY_Q4_LOB] > 0)) | ((parent_q4_result[Q4_LOB]==0) & (parent_q4_result[YoY_Q4_LOB] > 0)) )
    parent_q4_result.loc[lost_mask, 'WoW Change'] = 'Lost'
    parent_q4_result[YoY_Q4_LOB].fillna(value='New This Year', inplace=True)
    parent_q4_result.loc[parent_q4_result['YoY Change %'] == np.inf, YoY_Q4_LOB] = 'New This Year'
    
    #Label New This Week
    new_week_cols = [child_result,parent_result,parent_q1_result,parent_q2_result,parent_q3_result,parent_q4_result,
                industry_result,agency_holding_result,agency_indy_result,indy_result,agency_result]
    for j in new_week_cols:
        j.loc[j['WoW Change %'] == np.inf, 'WoW Change'] = 'New This Week'
        j['WoW Change'].fillna(value='New This Week', inplace=True)

    #--------- Create the Final Dataframes for the Summary Tab

    #DROP THE LAST WEEK TOTAL COLUMN
    child_final = child_result.drop(LW_Total_LOB, axis=1)
    parent_final = parent_result.drop(LW_Total_LOB, axis=1)
    q1_final = parent_q1_result.drop(LW_Q1_LOB, axis=1)
    q2_final = parent_q2_result.drop(LW_Q2_LOB, axis=1)
    q3_final = parent_q3_result.drop(LW_Q3_LOB, axis=1)
    q4_final = parent_q4_result.drop(LW_Q4_LOB, axis=1)
    industry_final = industry_result.drop(LW_Total_LOB, axis=1)
    agency_holding_final = agency_holding_result.drop(LW_Total_LOB, axis=1)
    agency_indy_final = agency_indy_result.drop(LW_Total_LOB, axis=1)
    indy_final = indy_result.drop(LW_Total_LOB, axis=1)
    agency_final = agency_result.drop(LW_Total_LOB, axis=1)

    #Remove NP Inf Values
    final_df = [child_final, parent_final, q1_final, q2_final, q3_final, q4_final, industry_final,
               agency_holding_final, indy_final, agency_indy_final, agency_final]

    for dfs in final_df:
        dfs.replace([np.inf, -np.inf], 'N/A', inplace=True)
        dfs['YoY Change %'].fillna(value='N/A', inplace=True)
        dfs['WoW Change %'].fillna(value='N/A', inplace=True)
    
    child_final = child_final[(child_final[Total_LOB] > 0) | (child_final['WoW Change'] == 'Lost')] 
    parent_final = parent_final[(parent_final[Total_LOB] > 0) | (parent_final['WoW Change'] == 'Lost')] 
    q1_final = q1_final[(q1_final[Q1_LOB] > 0) | (q1_final['WoW Change'] == 'Lost')] 
    q2_final = q2_final[(q2_final[Q2_LOB] > 0) | (q2_final['WoW Change'] == 'Lost')] 
    q3_final = q3_final[(q3_final[Q3_LOB] > 0) | (q3_final['WoW Change'] == 'Lost')] 
    q4_final = q4_final[(q4_final[Q4_LOB] > 0) | (q4_final['WoW Change'] == 'Lost')]
    industry_final = industry_final[(industry_final[Total_LOB] > 0) | (industry_final['WoW Change'] == 'Lost')] 
    agency_holding_final = agency_holding_final[(agency_holding_final[Total_LOB] > 0) | (agency_holding_final['WoW Change'] == 'Lost')] 
    agency_indy_final = agency_indy_final[(agency_indy_final[Total_LOB] > 0) | (agency_indy_final['WoW Change'] == 'Lost')] 
    indy_final = indy_final[(indy_final[Total_LOB] > 0) | (indy_final['WoW Change'] == 'Lost')]
    agency_final = agency_final[(agency_final[Total_LOB] > 0) | (agency_final['WoW Change'] == 'Lost')] 
    
    
    #EXPORT SECTION

    wb.sheets[f'{i}'].range('B8').options(index=False,header=False).value = child_final
    wb.sheets[f'{i}'].range('M8').options(index=False,header=False).value = parent_final
    wb.sheets[f'{i}'].range('X8').options(index=False,header=False).value = q1_final
    wb.sheets[f'{i}'].range('AE8').options(index=False,header=False).value = q2_final
    wb.sheets[f'{i}'].range('AL8').options(index=False,header=False).value = q3_final
    wb.sheets[f'{i}'].range('AS8').options(index=False,header=False).value = q4_final
    wb.sheets[f'{i}'].range('AZ8').options(index=False,header=False).value = industry_final
    wb.sheets[f'{i}'].range('BK8').options(index=False,header=False).value = indy_final
    wb.sheets[f'{i}'].range('BV8').options(index=False,header=False).value = agency_final


    #ADD THE RAW DATAFRAMES AS WELL
    wb.sheets[f'{i} Raw'].range('A2').options(index=False,header=True).value = child_result
    wb.sheets[f'{i} Raw'].range('M2').options(index=False,header=True).value = parent_result
    wb.sheets[f'{i} Raw'].range('Y2').options(index=False,header=True).value = parent_q1_result
    wb.sheets[f'{i} Raw'].range('AG2').options(index=False,header=True).value = parent_q2_result
    wb.sheets[f'{i} Raw'].range('AO2').options(index=False,header=True).value = parent_q3_result
    wb.sheets[f'{i} Raw'].range('AW2').options(index=False,header=True).value = parent_q4_result
    wb.sheets[f'{i} Raw'].range('BE2').options(index=False,header=True).value = industry_result
    wb.sheets[f'{i} Raw'].range('BQ2').options(index=False,header=True).value = indy_result
    wb.sheets[f'{i} Raw'].range('CC2').options(index=False,header=True).value = agency_result


# In[9]:


dest = f"C:/Users/Alexander Ravazzoni/Documents/LOB Lookback/LOB Lookback {today_date} Template.xlsx"

wb.save(dest)


# In[10]:


print(dest)

