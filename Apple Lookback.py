#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import xlwings as xw
import openpyxl
from datetime import date
import numpy as np

today_date = '2023-10-06'


# In[2]:


# reading excel spreadsheets
file_name = 'C:/Users/Alexander Ravazzoni/Documents/Apple Lookback/Apple Lookback Test Template.xlsx'

#This week's data
new_df = pd.read_excel(file_name, sheet_name = 'This Year New')
#Last week's data
old_df = pd.read_excel(file_name, sheet_name = 'This Year Old')
#Last year's data
yoy_df = pd.read_excel(file_name, sheet_name = 'Last Year Data')


# In[3]:


#Assigning Column Names to Dataframes
new_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category', 'Account Subcategory', 
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'Territory', 'Q1 Total','Q1 Local',
                  'Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q1 Prime', 'Q2 Total', 'Q2 Local', 'Q2 Multi', 'Q2 CTV', 'Q2 PG', 'Q2 Prime',
                  'Q3 Total', 'Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Q3 Prime', 'Q4 Total', 'Q4 Local', 'Q4 Multi', 'Q4 CTV',
                  'Q4 PG', 'Q4 Prime', 'Total Amount Net', 'CY Local', 'CY Multi', 'CY CTV', 'CY PG', 'CY Prime']

old_df.columns = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category', 'Account Subcategory', 
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'Territory','LW Q1 Total','LW Q1 Local',
                  'LW Q1 Multi', 'LW Q1 CTV', 'LW Q1 PG', 'LW Q1 Prime', 'LW Q2 Total', 'LW Q2 Local', 'LW Q2 Multi', 'LW Q2 CTV',
                  'LW Q2 PG', 'LW Q2 Prime', 'LW Q3 Total', 'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV', 'LW Q3 PG', 'LW Q3 Prime', 'LW Q4 Total',
                  'LW Q4 Local', 'LW Q4 Multi', 'LW Q4 CTV','LW Q4 PG', 'LW Q4 Prime', 'LW Total Amount Net', 'LW CY Local',
                  'LW CY Multi', 'LW CY CTV', 'LW CY PG', 'LW CY Prime']

last_year_cols = ['Order ID', 'Team Name', 'Deal Name', 'Advertiser', 'Account Category', 'Account Subcategory', 
                  'Agency', 'Agency Holding Co', 'Teammember Name', 'Ultimate Parent', 'Territory','LY Q1 Total','LY Q1 Local',
                  'LY Q1 Multi', 'LY Q1 CTV', 'LY Q1 PG', 'LY Q1 Prime', 'LY Q2 Total', 'LY Q2 Local', 'LY Q2 Multi', 'LY Q2 CTV',
                  'LY Q2 PG', 'LY Q2 Prime', 'LY Q3 Total', 'LY Q3 Local', 'LY Q3 Multi', 'LY Q3 CTV', 'LY Q3 PG', 'LY Q3 Prime', 'LY Q4 Total',
                  'LY Q4 Local', 'LY Q4 Multi', 'LY Q4 CTV','LY Q4 PG', 'LY Q4 Prime', 'LY Total Amount Net', 'LY CY Local',
                  'LY CY Multi', 'LY CY CTV', 'LY CY PG', 'LY CY Prime']

yoy_df.columns = last_year_cols


# In[ ]:





# In[4]:


#Fill in 0s in blank dollar spots

new_rev_cols = ['Q1 Total', 'Q1 Local', 'Q1 Multi', 'Q1 CTV', 'Q1 PG', 'Q1 Prime', 'Q2 Total', 'Q2 Local', 'Q2 Multi',
                'Q2 CTV','Q2 PG', 'Q2 Prime', 'Q3 Total', 'Q3 Local', 'Q3 Multi', 'Q3 CTV', 'Q3 PG', 'Q3 Prime', 'Q4 Total', 'Q4 Local',
                'Q4 Multi', 'Q4 CTV', 'Q4 PG', 'Q4 Prime']

for cols in new_rev_cols:
    new_df[cols].fillna(value=0, inplace = True)

old_rev_cols = ['LW Q1 Total', 'LW Q1 Local', 'LW Q1 Multi', 'LW Q1 CTV', 'LW Q1 PG', 'LW Q1 Prime', 'LW Q2 Total', 'LW Q2 Local',
                'LW Q2 Multi', 'LW Q2 CTV','LW Q2 PG', 'LW Q2 Prime', 'LW Q3 Total', 'LW Q3 Local', 'LW Q3 Multi', 'LW Q3 CTV',
                'LW Q3 PG', 'LW Q3 Prime', 'LW Q4 Total', 'LW Q4 Local', 'LW Q4 Multi','LW Q4 CTV', 'LW Q4 PG', 'LW Q4 Prime',]

for cols in old_rev_cols:
    old_df[cols].fillna(value=0, inplace = True)

yoy_rev_cols = ['LY Q1 Total', 'LY Q1 Local', 'LY Q1 Multi', 'LY Q1 CTV', 'LY Q1 PG', 'LY Q1 Prime',
               'LY Q2 Total', 'LY Q2 Local', 'LY Q2 Multi', 'LY Q2 CTV', 'LY Q2 PG', 'LY Q2 Prime',
               'LY Q3 Total', 'LY Q3 Local', 'LY Q3 Multi', 'LY Q3 CTV', 'LY Q3 PG', 'LY Q3 Prime',
               'LY Q4 Total', 'LY Q4 Local', 'LY Q4 Multi', 'LY Q4 CTV', 'LY Q4 PG', 'LY Q4 Prime',]

for cols in yoy_rev_cols:
    yoy_df[cols].fillna(value=0, inplace = True)


# In[5]:


#removing unwanted characters
unwanted_chars = ['~', ' - Global']

for char in unwanted_chars:
    new_df = new_df.replace(char, '', regex=True)
    old_df = old_df.replace(char, '', regex=True)
    yoy_df = yoy_df.replace(char, '', regex=True)


# In[6]:


#Grouping this week's data

## -- this is the apple summary tab
new_territory_df = new_df.groupby(['Territory']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'], ascending=False).reset_index()
new_territory_brand_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'], ascending=False).reset_index()
new_q1_terr_brand_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 Total':'sum'}).sort_values(by=['Q1 Total'], ascending=False).reset_index()
new_q2_terr_brand_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q2 Total':'sum'}).sort_values(by=['Q2 Total'], ascending=False).reset_index()
new_q3_terr_brand_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q3 Total':'sum'}).sort_values(by=['Q3 Total'], ascending=False).reset_index()
new_q4_terr_brand_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q4 Total':'sum'}).sort_values(by=['Q4 Total'], ascending=False).reset_index()
new_territory_subcat_df = new_df.groupby(['Territory', 'Account Subcategory']).agg({'Q1 Total':'sum', 'Q2 Total':'sum', 'Q3 Total':'sum', 'Q4 Total':'sum', 'Total Amount Net':'sum'}).sort_values(by=['Total Amount Net'], ascending=False).reset_index()

## -- this is the apple LOB tab
new_ctv_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 CTV':'sum', 'Q2 CTV':'sum', 'Q3 CTV':'sum', 'Q4 CTV':'sum', 'CY CTV':'sum'}).sort_values(by=['CY CTV'], ascending=False).reset_index()
new_multi_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 Multi':'sum', 'Q2 Multi':'sum', 'Q3 Multi':'sum', 'Q4 Multi':'sum', 'CY Multi':'sum'}).sort_values(by=['CY Multi'], ascending=False).reset_index()
new_local_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 Local':'sum', 'Q2 Local':'sum', 'Q3 Local':'sum', 'Q4 Local':'sum', 'CY Local':'sum'}).sort_values(by=['CY Local'], ascending=False).reset_index()
new_pg_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 PG':'sum', 'Q2 PG':'sum', 'Q3 PG':'sum', 'Q4 PG':'sum', 'CY PG':'sum'}).sort_values(by=['CY PG'], ascending=False).reset_index()
new_prime_df = new_df.groupby(['Territory', 'Advertiser']).agg({'Q1 Prime':'sum', 'Q2 Prime':'sum', 'Q3 Prime':'sum', 'Q4 Prime':'sum', 'CY Prime':'sum'}).sort_values(by=['CY Prime'], ascending=False).reset_index()


# In[7]:


#Grouping last week's data

## -- this is the apple summary tab
old_territory_df = old_df.groupby(['Territory']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'], ascending=False).reset_index()
old_territory_brand_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'], ascending=False).reset_index()
old_q1_terr_brand_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW Q1 Total':'sum'}).sort_values(by=['LW Q1 Total'], ascending=False).reset_index()
old_q2_terr_brand_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW Q2 Total':'sum'}).sort_values(by=['LW Q2 Total'], ascending=False).reset_index()
old_q3_terr_brand_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW Q3 Total':'sum'}).sort_values(by=['LW Q3 Total'], ascending=False).reset_index()
old_q4_terr_brand_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW Q4 Total':'sum'}).sort_values(by=['LW Q4 Total'], ascending=False).reset_index()
old_territory_subcat_df = old_df.groupby(['Territory', 'Account Subcategory']).agg({'LW Total Amount Net':'sum'}).sort_values(by=['LW Total Amount Net'], ascending=False).reset_index()

## -- this is the apple LOB tab
old_ctv_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW CY CTV':'sum'}).sort_values(by=['LW CY CTV'], ascending=False).reset_index()
old_multi_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW CY Multi':'sum'}).sort_values(by=['LW CY Multi'], ascending=False).reset_index()
old_local_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW CY Local':'sum'}).sort_values(by=['LW CY Local'], ascending=False).reset_index()
old_pg_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW CY PG':'sum'}).sort_values(by=['LW CY PG'], ascending=False).reset_index()
old_prime_df = old_df.groupby(['Territory', 'Advertiser']).agg({'LW CY Prime':'sum'}).sort_values(by=['LW CY Prime'], ascending=False).reset_index()


# In[8]:


#Grouping last year's data

## -- this is the apple summary tab
yoy_territory_df = yoy_df.groupby(['Territory']).agg({'LY Total Amount Net':'sum'}).sort_values(by=['LY Total Amount Net'], ascending=False).reset_index()
yoy_territory_brand_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY Total Amount Net':'sum'}).sort_values(by=['LY Total Amount Net'], ascending=False).reset_index()
yoy_q1_terr_brand_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY Q1 Total':'sum'}).sort_values(by=['LY Q1 Total'], ascending=False).reset_index()
yoy_q2_terr_brand_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY Q2 Total':'sum'}).sort_values(by=['LY Q2 Total'], ascending=False).reset_index()
yoy_q3_terr_brand_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY Q3 Total':'sum'}).sort_values(by=['LY Q3 Total'], ascending=False).reset_index()
yoy_q4_terr_brand_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY Q4 Total':'sum'}).sort_values(by=['LY Q4 Total'], ascending=False).reset_index()
yoy_territory_subcat_df = yoy_df.groupby(['Territory', 'Account Subcategory']).agg({'LY Total Amount Net':'sum'}).sort_values(by=['LY Total Amount Net'], ascending=False).reset_index()

## -- this is the apple LOB tab
yoy_ctv_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY CY CTV':'sum'}).sort_values(by=['LY CY CTV'], ascending=False).reset_index()
yoy_multi_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY CY Multi':'sum'}).sort_values(by=['LY CY Multi'], ascending=False).reset_index()
yoy_local_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY CY Local':'sum'}).sort_values(by=['LY CY Local'], ascending=False).reset_index()
yoy_pg_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY CY PG':'sum'}).sort_values(by=['LY CY PG'], ascending=False).reset_index()
yoy_prime_df = yoy_df.groupby(['Territory', 'Advertiser']).agg({'LY CY Prime':'sum'}).sort_values(by=['LY CY Prime'], ascending=False).reset_index()


# In[9]:


#Merging this & last week dataframes

## --- Apple Summary Tab DFs
territory_result = new_territory_df.merge(old_territory_df,how='outer', on='Territory').merge(yoy_territory_df,how='outer', on='Territory')
territory_result.insert(7,'WoW Change','')
territory_result['WoW Change'] = territory_result['Total Amount Net'] - territory_result['LW Total Amount Net']
territory_result.insert(8, 'WoW Change %', '')
territory_result['WoW Change %'] = (territory_result['Total Amount Net'] - territory_result['LW Total Amount Net'])/territory_result['LW Total Amount Net']
territory_result['YoY Change %'] = (territory_result['Total Amount Net'] - territory_result['LY Total Amount Net'])/territory_result['LY Total Amount Net']

terr_brand_result = new_territory_brand_df.merge(old_territory_brand_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_territory_brand_df, how='outer', on = ['Territory', 'Advertiser'])
terr_brand_result.insert(8, 'WoW Change', '')
terr_brand_result['WoW Change'] = terr_brand_result['Total Amount Net'] - terr_brand_result['LW Total Amount Net']
terr_brand_result.insert(9, 'WoW Change %', '')
terr_brand_result['WoW Change %'] = (terr_brand_result['Total Amount Net'] - terr_brand_result['LW Total Amount Net'])/terr_brand_result['LW Total Amount Net']
terr_brand_result['YoY Change %'] = (terr_brand_result['Total Amount Net'] - terr_brand_result['LY Total Amount Net'])/terr_brand_result['LY Total Amount Net']

q1_terr_brand_result = new_q1_terr_brand_df.merge(old_q1_terr_brand_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_q1_terr_brand_df, how='outer', on = ['Territory', 'Advertiser'])
q1_terr_brand_result.insert(4, 'WoW Change', '')
q1_terr_brand_result['WoW Change'] = q1_terr_brand_result['Q1 Total'] - q1_terr_brand_result['LW Q1 Total']
q1_terr_brand_result.insert(5, 'WoW Change %', '')
q1_terr_brand_result['WoW Change %'] = (q1_terr_brand_result['Q1 Total'] - q1_terr_brand_result['LW Q1 Total'])/q1_terr_brand_result['LW Q1 Total']
q1_terr_brand_result['YoY Change %'] = (q1_terr_brand_result['Q1 Total'] - q1_terr_brand_result['LY Q1 Total'])/q1_terr_brand_result['LY Q1 Total']

q2_terr_brand_result = new_q2_terr_brand_df.merge(old_q2_terr_brand_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_q2_terr_brand_df, how='outer', on = ['Territory', 'Advertiser'])
q2_terr_brand_result.insert(4, 'WoW Change', '')
q2_terr_brand_result['WoW Change'] = q2_terr_brand_result['Q2 Total'] - q2_terr_brand_result['LW Q2 Total']
q2_terr_brand_result.insert(5, 'WoW Change %', '')
q2_terr_brand_result['WoW Change %'] = (q2_terr_brand_result['Q2 Total'] - q2_terr_brand_result['LW Q2 Total'])/q2_terr_brand_result['LW Q2 Total']
q2_terr_brand_result['YoY Change %'] = (q2_terr_brand_result['Q2 Total'] - q2_terr_brand_result['LY Q2 Total'])/q2_terr_brand_result['LY Q2 Total']

q3_terr_brand_result = new_q3_terr_brand_df.merge(old_q3_terr_brand_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_q3_terr_brand_df, how='outer', on = ['Territory', 'Advertiser'])
q3_terr_brand_result.insert(4, 'WoW Change', '')
q3_terr_brand_result['WoW Change'] = q3_terr_brand_result['Q3 Total'] - q3_terr_brand_result['LW Q3 Total']
q3_terr_brand_result.insert(5, 'WoW Change %', '')
q3_terr_brand_result['WoW Change %'] = (q3_terr_brand_result['Q3 Total'] - q3_terr_brand_result['LW Q3 Total'])/q3_terr_brand_result['LW Q3 Total']
q3_terr_brand_result['YoY Change %'] = (q3_terr_brand_result['Q3 Total'] - q3_terr_brand_result['LY Q3 Total'])/q3_terr_brand_result['LY Q3 Total']

q4_terr_brand_result = new_q4_terr_brand_df.merge(old_q4_terr_brand_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_q4_terr_brand_df, how='outer', on = ['Territory', 'Advertiser'])
q4_terr_brand_result.insert(4, 'WoW Change', '')
q4_terr_brand_result['WoW Change'] = q4_terr_brand_result['Q4 Total'] - q4_terr_brand_result['LW Q4 Total']
q4_terr_brand_result.insert(5, 'WoW Change %', '')
q4_terr_brand_result['WoW Change %'] = (q4_terr_brand_result['Q4 Total'] - q4_terr_brand_result['LW Q4 Total'])/q4_terr_brand_result['LW Q4 Total']
q4_terr_brand_result['YoY Change %'] = (q4_terr_brand_result['Q4 Total'] - q4_terr_brand_result['LY Q4 Total'])/q4_terr_brand_result['LY Q4 Total']

terr_subcat_result = new_territory_subcat_df.merge(old_territory_subcat_df,how='outer', on=['Territory', 'Account Subcategory']).merge(yoy_territory_subcat_df,how='outer', on=['Territory', 'Account Subcategory'])
terr_subcat_result.insert(8,'WoW Change','')
terr_subcat_result['WoW Change'] = terr_subcat_result['Total Amount Net'] - terr_subcat_result['LW Total Amount Net']
terr_subcat_result.insert(9, 'WoW Change %', '')
terr_subcat_result['WoW Change %'] = (terr_subcat_result['Total Amount Net'] - terr_subcat_result['LW Total Amount Net'])/terr_subcat_result['LW Total Amount Net']
terr_subcat_result['YoY Change %'] = (terr_subcat_result['Total Amount Net'] - terr_subcat_result['LY Total Amount Net'])/terr_subcat_result['LY Total Amount Net']

## ---- LOB Summary Tab DFs
ctv_result = new_ctv_df.merge(old_ctv_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_ctv_df, how='outer', on = ['Territory', 'Advertiser'])
ctv_result.insert(8, 'WoW Change', '')
ctv_result['WoW Change'] = ctv_result['CY CTV'] - ctv_result['LW CY CTV']
ctv_result.insert(9, 'WoW Change %', '')
ctv_result['WoW Change %'] = (ctv_result['CY CTV'] - ctv_result['LW CY CTV'])/ctv_result['LW CY CTV']
ctv_result['YoY Change %'] = (ctv_result['CY CTV'] - ctv_result['LY CY CTV'])/ctv_result['LY CY CTV']

multi_result = new_multi_df.merge(old_multi_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_multi_df, how='outer', on = ['Territory', 'Advertiser'])
multi_result.insert(8, 'WoW Change', '')
multi_result['WoW Change'] = multi_result['CY Multi'] - multi_result['LW CY Multi']
multi_result.insert(9, 'WoW Change %', '')
multi_result['WoW Change %'] = (multi_result['CY Multi'] - multi_result['LW CY Multi'])/multi_result['LW CY Multi']
multi_result['YoY Change %'] = (multi_result['CY Multi'] - multi_result['LY CY Multi'])/multi_result['LY CY Multi']

local_result = new_local_df.merge(old_local_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_local_df, how='outer', on = ['Territory', 'Advertiser'])
local_result.insert(8, 'WoW Change', '')
local_result['WoW Change'] = local_result['CY Local'] - local_result['LW CY Local']
local_result.insert(9, 'WoW Change %', '')
local_result['WoW Change %'] = (local_result['CY Local'] - local_result['LW CY Local'])/local_result['LW CY Local']
local_result['YoY Change %'] = (local_result['CY Local'] - local_result['LY CY Local'])/local_result['LY CY Local']

pg_result = new_pg_df.merge(old_pg_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_pg_df, how='outer', on = ['Territory', 'Advertiser'])
pg_result.insert(8, 'WoW Change', '')
pg_result['WoW Change'] = pg_result['CY PG'] - pg_result['CY PG']
pg_result.insert(9, 'WoW Change %', '')
pg_result['WoW Change %'] = (pg_result['CY PG'] - pg_result['LW CY PG'])/pg_result['LW CY PG']
pg_result['YoY Change %'] = (pg_result['CY PG'] - pg_result['LY CY PG'])/pg_result['LY CY PG']

prime_result = new_prime_df.merge(old_prime_df,how='outer', on=['Territory', 'Advertiser']).merge(yoy_prime_df, how='outer', on = ['Territory', 'Advertiser'])
prime_result.insert(8, 'WoW Change', '')
prime_result['WoW Change'] = prime_result['CY Prime'] - prime_result['LW CY Prime']
prime_result.insert(9, 'WoW Change %', '')
prime_result['WoW Change %'] = (prime_result['CY Prime'] - prime_result['LW CY Prime'])/prime_result['LW CY Prime']
prime_result['YoY Change %'] = (prime_result['CY Prime'] - prime_result['LY CY Prime'])/prime_result['LY CY Prime']


# In[10]:


#q3_terr_brand_result


# In[11]:


# Identifying Lost

lost_dfs = [territory_result, terr_brand_result, terr_subcat_result]

for k in lost_dfs:
    mask1 = ( ((k['Total Amount Net'].isna()) & (k['LY Total Amount Net'] > 0)) | ((k['Total Amount Net'] == 0) & (k['LY Total Amount Net'] >0)) )
    k.loc[mask1, 'WoW Change'] = 'Lost'


# In[12]:


territory_result['LY Total Amount Net'].fillna(value='New This Year', inplace = True)
territory_result.loc[territory_result['YoY Change %'] == np.inf, 'LY Total Amount Net'] = 'New This Year'

terr_brand_result['LY Total Amount Net'].fillna(value='New This Year', inplace = True)
terr_brand_result.loc[terr_brand_result['YoY Change %'] == np.inf, 'LY Total Amount Net'] = 'New This Year'

terr_subcat_result['LY Total Amount Net'].fillna(value='New This Year', inplace = True)
terr_subcat_result.loc[terr_subcat_result['YoY Change %'] == np.inf, 'LY Total Amount Net'] = 'New This Year'

lost_mask = ( ((ctv_result['CY CTV'].isna()) & (ctv_result['LY CY CTV'] > 0)) | ((ctv_result['CY CTV']==0) & (ctv_result['LY CY CTV'] > 0)) )
ctv_result.loc[lost_mask, 'WoW Change'] = 'Lost'
ctv_result['LY CY CTV'].fillna(value='New This Year', inplace=True)
ctv_result.loc[ctv_result['YoY Change %'] == np.inf, 'LY CY CTV'] = 'New This Year'

lost_mask = ( ((multi_result['CY Multi'].isna()) & (multi_result['LY CY Multi'] > 0)) | ((multi_result['CY Multi']==0) & (multi_result['LY CY Multi'] > 0)) )
multi_result.loc[lost_mask, 'WoW Change'] = 'Lost'
multi_result['LY CY Multi'].fillna(value='New This Year', inplace=True)
multi_result.loc[multi_result['YoY Change %'] == np.inf, 'LY CY Multi'] = 'New This Year'

lost_mask = ( ((local_result['CY Local'].isna()) & (local_result['LY CY Local'] > 0)) | ((local_result['CY Local']==0) & (local_result['LY CY Local'] > 0)) )
local_result.loc[lost_mask, 'WoW Change'] = 'Lost'
local_result['LY CY Local'].fillna(value='New This Year', inplace=True)
local_result.loc[local_result['YoY Change %'] == np.inf, 'LY CY Local'] = 'New This Year'

lost_mask = ( ((pg_result['CY PG'].isna()) & (pg_result['LY CY PG'] > 0)) | ((pg_result['CY PG']==0) & (pg_result['LY CY PG'] > 0)) )
pg_result.loc[lost_mask, 'WoW Change'] = 'Lost'
pg_result['LY CY PG'].fillna(value='New This Year', inplace=True)
pg_result.loc[pg_result['YoY Change %'] == np.inf, 'LY CY PG'] = 'New This Year'

lost_mask = ( ((prime_result['CY Prime'].isna()) & (prime_result['LY CY Prime'] > 0)) | ((prime_result['CY Prime']==0) & (prime_result['LY CY Prime'] > 0)) )
prime_result.loc[lost_mask, 'WoW Change'] = 'Lost'
prime_result['LY CY Prime'].fillna(value='New This Year', inplace=True)
prime_result.loc[prime_result['YoY Change %'] == np.inf, 'LY CY Prime'] = 'New This Year'

lost_mask = ( ((q1_terr_brand_result['Q1 Total'].isna()) & (q1_terr_brand_result['LY Q1 Total'] > 0)) | ((q1_terr_brand_result['Q1 Total']==0) & (q1_terr_brand_result['LY Q1 Total'] > 0)) )
q1_terr_brand_result.loc[lost_mask, 'WoW Change'] = 'Lost'
q1_terr_brand_result['LY Q1 Total'].fillna(value='New This Year', inplace=True)
q1_terr_brand_result.loc[q1_terr_brand_result['YoY Change %'] == np.inf, 'LY Q1 Total'] = 'New This Year'

lost_mask = ( ((q2_terr_brand_result['Q2 Total'].isna()) & (q2_terr_brand_result['LY Q2 Total'] > 0)) | ((q2_terr_brand_result['Q2 Total']==0) & (q2_terr_brand_result['LY Q2 Total'] > 0)) )
q2_terr_brand_result.loc[lost_mask, 'WoW Change'] = 'Lost'
q2_terr_brand_result['LY Q2 Total'].fillna(value='New This Year', inplace=True)
q2_terr_brand_result.loc[q2_terr_brand_result['YoY Change %'] == np.inf, 'LY Q2 Total'] = 'New This Year'

lost_mask = ( ((q3_terr_brand_result['Q3 Total'].isna()) & (q3_terr_brand_result['LY Q3 Total'] > 0)) | ((q3_terr_brand_result['Q3 Total']==0) & (q3_terr_brand_result['LY Q3 Total'] > 0)) )
q3_terr_brand_result.loc[lost_mask, 'WoW Change'] = 'Lost'
q3_terr_brand_result['LY Q3 Total'].fillna(value='New This Year', inplace=True)
q3_terr_brand_result.loc[q3_terr_brand_result['YoY Change %'] == np.inf, 'LY Q3 Total'] = 'New This Year'

lost_mask = ( ((q4_terr_brand_result['Q4 Total'].isna()) & (q4_terr_brand_result['LY Q4 Total'] > 0)) | ((q4_terr_brand_result['Q4 Total']==0) & (q4_terr_brand_result['LY Q4 Total'] > 0)) )
q4_terr_brand_result.loc[lost_mask, 'WoW Change'] = 'Lost'
q4_terr_brand_result['LY Q4 Total'].fillna(value='New This Year', inplace=True)
q4_terr_brand_result.loc[q4_terr_brand_result['YoY Change %'] == np.inf, 'LY Q4 Total'] = 'New This Year'


# In[13]:


#q2_terr_brand_result


# In[14]:


new_week_cols = [territory_result, terr_brand_result, terr_subcat_result, ctv_result, multi_result, local_result, pg_result, 
                prime_result, q1_terr_brand_result, q2_terr_brand_result, q3_terr_brand_result, q4_terr_brand_result]

for j in new_week_cols:
    j.loc[j['WoW Change %'] == np.inf, 'WoW Change'] = 'New This Week'
    j['WoW Change'].fillna(value='New This Week', inplace=True)


# In[15]:


result_df = [territory_result, terr_brand_result, terr_subcat_result, ctv_result, multi_result, local_result, pg_result, 
                prime_result, q1_terr_brand_result, q2_terr_brand_result, q3_terr_brand_result, q4_terr_brand_result]

for dfs in result_df:
    dfs.replace([np.inf, -np.inf], np.nan, inplace=True)


# In[16]:


#Drop Last Week Total Column
terr_final = territory_result.drop('LW Total Amount Net', axis = 1)
terr__brand_final = terr_brand_result.drop('LW Total Amount Net', axis = 1)
q1_final = q1_terr_brand_result.drop('LW Q1 Total', axis = 1)
q2_final = q2_terr_brand_result.drop('LW Q2 Total', axis = 1)
q3_final = q3_terr_brand_result.drop('LW Q3 Total', axis = 1)
q4_final = q4_terr_brand_result.drop('LW Q4 Total', axis = 1)
terr_subcat_final = terr_subcat_result.drop('LW Total Amount Net', axis = 1)
ctv_final = ctv_result.drop('LW CY CTV', axis = 1)
multi_final = multi_result.drop('LW CY Multi', axis = 1)
local_final = local_result.drop('LW CY Local', axis = 1)
pg_final = pg_result.drop('LW CY PG', axis = 1)
prime_final = prime_result.drop('LW CY Prime', axis = 1)


# In[17]:


q1_final = q1_final[(q1_final['Q1 Total'] > 0) | (q1_final['WoW Change'] == 'Lost')] 
q2_final = q2_final[(q2_final['Q2 Total'] > 0) | (q2_final['WoW Change'] == 'Lost')] 
q3_final = q3_final[(q3_final['Q3 Total'] > 0) | (q3_final['WoW Change'] == 'Lost')] 
q4_final = q4_final[(q4_final['Q4 Total'] > 0) | (q4_final['WoW Change'] == 'Lost')] 
ctv_final = ctv_final[(ctv_final['CY CTV'] > 0) | (ctv_final['WoW Change'] == 'Lost')] 
local_final = local_final[(local_final['CY Local'] > 0) | (local_final['WoW Change'] == 'Lost')] 
multi_final = multi_final[(multi_final['CY Multi'] > 0) | (multi_final['WoW Change'] == 'Lost')] 
pg_final = pg_final[(pg_final['CY PG'] > 0) | (pg_final['WoW Change'] == 'Lost')] 
prime_final = prime_final[(prime_final['CY Prime'] > 0) | (prime_final['WoW Change'] == 'Lost')] 


# In[18]:


# Export Section

src = file_name

wb = xw.Book(src)

wb.sheets['Apple Summary'].range('K6').options(index=False, header=False).value = terr_final
wb.sheets['Apple Summary'].range('V6').options(index=False, header=False).value = terr__brand_final
wb.sheets['Apple Summary'].range('AH6').options(index=False, header=False).value = q1_final
wb.sheets['Apple Summary'].range('AP6').options(index=False, header=False).value = q2_final
wb.sheets['Apple Summary'].range('AX6').options(index=False, header=False).value = q3_final
wb.sheets['Apple Summary'].range('BF6').options(index=False, header=False).value = q4_final
wb.sheets['Apple Summary'].range('BN6').options(index=False, header=False).value = terr_subcat_final

wb.sheets['Apple LOB Summary'].range('I12').options(index=False, header=False).value = ctv_final
wb.sheets['Apple LOB Summary'].range('AD12').options(index=False, header=False).value = multi_final
wb.sheets['Apple LOB Summary'].range('AY12').options(index=False, header=False).value = local_final
wb.sheets['Apple LOB Summary'].range('BT12').options(index=False, header=False).value = pg_final
wb.sheets['Apple LOB Summary'].range('CO12').options(index=False, header=False).value = prime_final

## Adding Raw Dataframes

wb.sheets['Raw Apple Summary'].range('A2').options(index=False,header=True).value = territory_result
wb.sheets['Raw Apple Summary'].range('M2').options(index=False,header=True).value = terr_brand_result
wb.sheets['Raw Apple Summary'].range('Z2').options(index=False,header=True).value = q1_terr_brand_result
wb.sheets['Raw Apple Summary'].range('AI2').options(index=False,header=True).value = q2_terr_brand_result
wb.sheets['Raw Apple Summary'].range('AR2').options(index=False,header=True).value = q3_terr_brand_result
wb.sheets['Raw Apple Summary'].range('BA2').options(index=False,header=True).value = q4_terr_brand_result
wb.sheets['Raw Apple Summary'].range('BJ2').options(index=False,header=True).value = terr_subcat_result

wb.sheets['Raw LOB Summary'].range('A2').options(index=False, header=True).value = ctv_result
wb.sheets['Raw LOB Summary'].range('N2').options(index=False, header=True).value = multi_result
wb.sheets['Raw LOB Summary'].range('AA2').options(index=False, header=True).value = local_result
wb.sheets['Raw LOB Summary'].range('AN2').options(index=False, header=True).value = pg_result
wb.sheets['Raw LOB Summary'].range('BA2').options(index=False, header=True).value = prime_result

dest = f"C:/Users/Alexander Ravazzoni/Documents/Apple Lookback/Apple Lookback {today_date} Template.xlsx"

wb.save(dest)


# In[19]:


## Insights (beta)
### insights I want to do revolve around overall booked by country, any new deal(s)/ advs and a basic LOB Breakout
## seems like the Apple LB is simpler

## this/last week & last year totals 
td_total = new_df['Total Amount Net'].sum()
lw_total = old_df['LW Total Amount Net'].sum()
ly_total = yoy_df['LY Total Amount Net'].sum()

# LOB total
ctv_total = new_df['CY CTV'].sum()
local_total = new_df['CY Local'].sum()
multi_total = new_df['CY Multi'].sum()
pg_total = new_df['CY PG'].sum()
prime_total = new_df['CY Prime'].sum()

# LOB % Weight
ctv_percent = '{:.1%}'.format(ctv_total/td_total)
local_percent = '{:.1%}'.format(local_total/td_total)
multi_percent = '{:.1%}'.format(multi_total/td_total)
pg_percent = '{:.1%}'.format(pg_total/td_total)

# WoW/ YoY changes
wow_diff = "${:,0f}".format(td_total - lw_total)
yoy_diff = "${:,.0f}".format(td_total - ly_total)

if td_total - lw_total > 0:
    wow_rev_dir = 'up'
else:
    wow_rev_dir = 'down'
    
if td_total - ly_total > 0:
    yoy_rev_dir = 'up'
else:
    yoy_rev_dir = 'down'


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




