#!/usr/bin/env python
# coding: utf-8
################################################################################################################################################################################################################################################################################################################################################
# Purpose: Analyze revenue in multiple time periods, per vendor
#
# Input: 3 csv files 
#        Data: Sales Analysis Table in EVP, 3 different time periods
#        Note1: Here We use 3 csv files for November 2020 , December 2020, and January 1-14, 2021 
#        Note2: Add files in the code if you need to analyze more than 3 periods of time 
# 
# Required Columns: [Sales Sku, Sales Order Number, Sales Order Date, Sales Channel Name, Fulfillment Item Id, Fulfillment Sku, Fulfillment Order Number, Fulfillment Channel Name, Fulfillment Channel Type, Quantity, Sku, Total Sales, Total Cost, Commission, Inventory Cost, Estimated Shipping Cost, Shipping Cost]
#        Note1: To reduce redundant resources, it's better to remove other columns. Still can run if you do not remove them.
#        Note2: Before running this program, check column names in Sales Analysis Table, especially empty spaces.
#
# Output: 1 csv file
#        csv: Business information in multiple time periods, per vendor
# 
# Customized configuration - Only need to change variables below: 
# * vendor_list    <- Add/Drop vendors
# * nov_data       <- csv file of Sales Analysis Table in EVP (1st time period you would like to analyze)
# * dec_data       <- csv file of Sales Analysis Table in EVP (2nd time period you would like to analyze)
# * jan_data       <- csv file of Sales Analysis Table in EVP (3rd time period you would like to analyze)
# * result.to_csv  <- Path of csv output file
#
# Optional Comments: If you want to create a text file for output, please refer to the comment # (Optional: write data in a text file)
#
################################################################################################################################################################################################################################################################################################################################################
import pandas as pd
import numpy as np

# List of vendors, add here when cooperating new vendors
vendor_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']

# import two files of previous and current month
nov_data = pd.read_csv('data/sale_data_nov_trim.CSV')  
dec_data = pd.read_csv('data/sale_data_dec_trim.CSV')
jan_data = pd.read_csv('data/sale_data_jan0114_trim.csv')

# Filter to EVP data only
#nov_data = nov_data[nov_data['Sales Sku'].str.contains('EVP')] 
#dec_data = dec_data[dec_data['Sales Sku'].str.contains('EVP')]
#jan_data = jan_data[jan_data['Sales Sku'].str.contains('EVP')]

# Filter Dropship vendors
nov_data = nov_data[nov_data['Fulfillment Channel Name'].isin(vendor_list)]
dec_data = dec_data[dec_data['Fulfillment Channel Name'].isin(vendor_list)]
jan_data = jan_data[jan_data['Fulfillment Channel Name'].isin(vendor_list)]


# Create total sales for current two months, group by each vendor
nov_sale_data = nov_data[['Fulfillment Channel Name', 'Total Sales']]
dec_sale_data = dec_data[['Fulfillment Channel Name', 'Total Sales']]
jan_sale_data = jan_data[['Fulfillment Channel Name', 'Total Sales']]
nov_agg_data = nov_sale_data.groupby('Fulfillment Channel Name')['Total Sales'].agg(['sum','count'])
nov_agg_data = nov_agg_data.rename(columns={'sum':'Sales_Nov', 'count':'count_Nov'})
dec_agg_data = dec_sale_data.groupby('Fulfillment Channel Name')['Total Sales'].agg(['sum','count'])
dec_agg_data = dec_agg_data.rename(columns={'sum':'Sales_Dec', 'count':'count_Dec'})
jan_agg_data = jan_sale_data.groupby('Fulfillment Channel Name')['Total Sales'].agg(['sum','count'])
jan_agg_data = jan_agg_data.rename(columns={'sum':'Sales_Jan0114', 'count':'count_Jan0114'})

# Combine two monthly data
result = pd.concat([nov_agg_data, dec_agg_data, jan_agg_data], axis=1)
result['Value per order Nov'] = result['Sales_Nov']/result['count_Nov']
result['Value per order Dec'] = result['Sales_Dec']/result['count_Dec']
result['Value per order Jan0114'] = result['Sales_Jan0114']/result['count_Jan0114']

result = result[['Sales_Nov', 'Sales_Dec', 'Sales_Jan0114', 'count_Nov', 'count_Dec', 'count_Jan0114', 'Value per order Nov',                  'Value per order Dec', 'Value per order Jan0114']]
result = result.append(result.sum().rename('Total'))

# Write data in a file
result.to_csv('output/analysis_result/Sales_by_Vendors/Monthly_Info_by_vendors.csv')

################################# For Reference #################################
# (Optional: write data in a text file)
#####with open('analysis_result/Monthly_comparation_by_vendors.txt', 'w') as f:
        #f.write('[Fundamental Statistics] \n')
        #f.write(result)
        #f.write('Top 10 overlapped items with negative effect: \n')
    #####f.write(result.__repr__())