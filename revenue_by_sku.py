#!/usr/bin/env python
# coding: utf-8
################################################################################################################################################################################################################################################################################################################################################
# Purpose: Analyze revenue by each SKU
#
# Input: 1 csv files 
#        Data: Sales Analysis Table in EVP, at least for 90 days
#        Note1: Here We use 1 csv file for October 2020 - January 2021, 90 days
# 
# Required Columns: [SalesSku, SalesOrderNumber, SalesOrderDate, FulfillmentOrderNumber, FulfillmentChannelName, FulfillmentChannelType, Quantity, Sku, TotalSales, TotalCost, Commission, InventoryCost, EstimatedShippingCost, ShippingCost]
#        Note1: To reduce redundant resources, it's better to remove other columns. Still can run if you do not remove them.
#        Note2: Before running this program, check column names in Sales Analysis Table, especially empty spaces.
#
# Output: 1 xlsx file
#        xlsx: Top 10 sales of SKUs in 7/30/60/90 days
# 
# Customized configuration - Only need to change variables below: 
# * vendor_list    <- Add/Drop vendors
# * data           <- csv file of Sales Analysis Table in EVP, at least 90 days
# * pd.ExcelWriter <- Path of excel output file
#
# Optional Comments: If you want to create a text file for output, please refer to the comment (# Write in a text file)
#
################################################################################################################################################################################################################################################################################################################################################

import pandas as pd
import numpy as np
import datetime

# List of vendors, add here when cooperating new vendors
vendor_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
# Creating Excel Writer Object from Pandas  
writer = pd.ExcelWriter('output/analysis_result/Sales_Analysis_by_SKU.xlsx',engine='xlsxwriter', mode='w')   
workbook = writer.book

# Read data
data = pd.read_csv('data/sale_data_90days_trim.CSV') 

# Filter to EVP data only
#data = data[data['SalesSku'].str.contains('EVP')]

# Filter Dropship vendors
data = data[data['FulfillmentChannelName'].isin(vendor_list)]

# Prepare date range with 7/30/60/90 days
data['SalesOrderDate']= pd.to_datetime(data['SalesOrderDate'])
data_7days = data[data['SalesOrderDate'] > datetime.datetime.now() - pd.to_timedelta("7day")] #1929
data_30days = data[data['SalesOrderDate'] > datetime.datetime.now() - pd.to_timedelta("30day")] #9721
data_60days = data[data['SalesOrderDate'] > datetime.datetime.now() - pd.to_timedelta("60day")] #19562
data_90days = data[data['SalesOrderDate'] > datetime.datetime.now() - pd.to_timedelta("90day")] #25890

# Top sales in 7 days
sales_7days = data_7days.groupby('Sku')['TotalSales'].agg(['sum','count'])
sort_sales_7days = sales_7days.sort_values(by='sum', ascending=False)
result_7days = sort_sales_7days[:10].rename(columns={'sum':'Sales'})
result_7days.reset_index(inplace=True)
result_7days.index += 1

# Top sales in 30 days
sales_30days = data_30days.groupby('Sku')['TotalSales'].agg(['sum','count'])
sort_sales_30days = sales_30days.sort_values(by='sum', ascending=False)
result_30days = sort_sales_30days[:10].rename(columns={'sum':'Sales'})
result_30days.reset_index(inplace=True)
result_30days.index += 1

# Top sales in 60 days
sales_60days = data_60days.groupby('Sku')['TotalSales'].agg(['sum','count'])
sort_sales_60days = sales_60days.sort_values(by='sum', ascending=False)
result_60days = sort_sales_60days[:10].rename(columns={'sum':'Sales'})
result_60days.reset_index(inplace=True)
result_60days.index += 1

# Top sales in 90 days
sales_90days = data_90days.groupby('Sku')['TotalSales'].agg(['sum','count'])
sort_sales_90days = sales_90days.sort_values(by='sum', ascending=False)
result_90days = sort_sales_90days[:10].rename(columns={'sum':'Sales'})
result_90days.reset_index(inplace=True)
result_90days.index += 1
    
# Write data in Excel file
v = 'Top20_recent_3Months'
worksheet=workbook.add_worksheet(v)
writer.sheets[v] = worksheet
worksheet.write_string(0, 0, 'Top 10 sales in 7 days: ')
result_7days.to_excel(writer,sheet_name=v,startrow=1 , startcol=0)   
worksheet.write_string(13, 0, 'Top 10 sales in 30 days: ')
result_30days.to_excel(writer,sheet_name=v,startrow=14, startcol=0)
worksheet.write_string(26, 0, 'Top 10 sales in 60 days: ')
result_60days.to_excel(writer,sheet_name=v,startrow=27, startcol=0)
worksheet.write_string(39, 0, 'Top 10 sales in 90 days: ')
result_90days.to_excel(writer,sheet_name=v,startrow=40, startcol=0)

# Save file for xlsx
writer.save()

################################ For Reference ################################
# Write in a text file
#with open('analysis_result/Sales_Analysis_bySKU.txt', 'w') as f:
#    f.write('Top 10 sales in 7 days: \n')
#    f.write(result_7days.__repr__())
#    f.write('\n\nTop 10 sales in 30 days: \n')
#    f.write(result_30days.__repr__())
#    f.write('\n\nTop 10 sales in 60 days: \n')
#    f.write(result_60days.__repr__())
#    f.write('\n\nTop 10 sales in 90 days: \n')
#    f.write(result_90days.__repr__())