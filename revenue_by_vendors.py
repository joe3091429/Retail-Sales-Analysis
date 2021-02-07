#!/usr/bin/env python
# coding: utf-8
################################################################################################################################################################################################################################################################################################################################################
# Purpose: Analyze revenue by each vendor in two months
#
# Input: 2 csv files 
#        Data: Sales Analysis Table in EVP, two monthly csv files
#        Note1: Here We use 1 csv file for November 2020 & 1 csv file for December 2020)
# 
# Required Columns: [Sales Sku, Sales Order Number, Sales Order Date, Sales Channel Name, Fulfillment Item Id, Fulfillment Sku, Fulfillment Order Number, Fulfillment Channel Name, Fulfillment Channel Type, Quantity, Sku, Total Sales, Total Cost, Commission, Inventory Cost, Estimated Shipping Cost, Shipping Cost]
#        Note1: To reduce redundant resources, it's better to remove other columns. Still can run if you do not remove them.
#        Note2: Before running this program, check column names in Sales Analysis Table, especially empty spaces.
#
# Output: 1 xlsx, 1 csv file
#        xlsx: Top 10 / Bottom 10 SKUs in each vendor, one sheet per vendor.
#        csv: Business information by vendors
# 
# Customized configuration - Only need to change variables below: 
# * vendor_list    <- Add/Drop vendors
# * nov_data       <- csv file of Sales Analysis Table in EVP (Previous time period you would like to analyze)
# * dec_data       <- csv file of Sales Analysis Table in EVP (Current time period you would like to analyze)
# * pd.ExcelWriter <- Path of excel output file
# * info_df.to_csv <- Path of csv output file  
#
################################################################################################################################################################################################################################################################################################################################################

import pandas as pd
import numpy as np

# List of vendors, add here when cooperating new vendors
vendor_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']

# Creating Excel Writer Object from Pandas  
writer = pd.ExcelWriter('output/analysis_result/Sales_by_Vendors/Top_Bottom_Sku_by_Vendors.xlsx',engine='xlsxwriter', mode='w')   
workbook = writer.book

info_list = []
for v in vendor_list:
    # import two files of previous and current month
    nov_data = pd.read_csv('data/sale_data_nov_trim.CSV')  
    dec_data = pd.read_csv('data/sale_data_dec_trim.CSV')

    # Filter to EVP data only
    #nov_data = nov_data[nov_data['Sales Sku'].str.contains('EVP')] #9042
    #dec_data = dec_data[dec_data['Sales Sku'].str.contains('EVP')] #10303

    # Filter vendor's data
    nov_syn_data = nov_data[nov_data['Fulfillment Channel Name'] == v] #610
    dec_syn_data = dec_data[dec_data['Fulfillment Channel Name'] == v] #441

    # Keep columns: Sku, Total sales
    nov_syn_sale_data = nov_syn_data[['Sku', 'Total Sales']]
    dec_syn_sale_data = dec_syn_data[['Sku', 'Total Sales']]

    # Group by sku
    nov_syn_agg_data = nov_syn_sale_data.groupby('Sku')['Total Sales'].agg(['sum','count'])
    dec_syn_agg_data = dec_syn_sale_data.groupby('Sku')['Total Sales'].agg(['sum','count'])

    # Sort by total sales
    nov_syn_agg_data = nov_syn_agg_data.sort_values(by=['sum'], ascending=False)
    dec_syn_agg_data = dec_syn_agg_data.sort_values(by=['sum'], ascending=False)

    # top 10 negative overlapped items
    overlap_data = pd.merge(dec_syn_agg_data, nov_syn_agg_data, how='inner', on='Sku')
    overlap_data = overlap_data.reset_index()

    # Rename column
    overlap_data = overlap_data.rename(columns={"sum_x": "sales_dec", "count_x": "count_dec","sum_y": "sales_nov", "count_y": "count_nov"})

    # add a column for sales difference between two months
    overlap_data['diff'] = overlap_data['sales_dec'] - overlap_data['sales_nov']

    # sort by sales difference
    overlap_data = overlap_data.sort_values(by=['diff'], ascending=True)

    # top 10 negative impact
    top10_negative = overlap_data[["Sku", "diff"]][:10]
    top10_negative.index = np.arange(1, len(top10_negative) + 1)

    # top 10 positive overlapped items
    top10_positive = overlap_data[["Sku", "diff"]][-10:]
    top10_positive = top10_positive.sort_values(by=['diff'], ascending=False)
    top10_positive.index = np.arange(1, len(top10_positive) + 1)

    # top 10 sales in Synnex (Nov, Dec)
    top10_sales_nov = nov_syn_agg_data[:10].rename(columns={'sum':'Sales'})
    top10_sales_nov.reset_index(inplace=True)
    top10_sales_nov.index += 1

    top10_sales_dec = dec_syn_agg_data[:10].rename(columns={'sum':'Sales'})
    top10_sales_dec.reset_index(inplace=True)
    top10_sales_dec.index += 1

    # Basic stats
    # Total sales in Dec & separated by two types of items
    nov_sales = nov_syn_agg_data['sum'].sum()
    dec_sales = dec_syn_agg_data['sum'].sum()
    overlap_item_sales = overlap_data['sales_dec'].sum()
    new_item_sales = dec_sales - overlap_item_sales
    #print(dec_sales, overlap_item_sales, new_item_sales)

    # Total difference of overlapped items
    total_breakeven_overlap_item = overlap_data['diff'].sum()
    #print(total_breakeven_overlap_item)

    # Quantity of overlapped items & new items
    total_qty_items_current_month = dec_syn_agg_data['sum'].count()
    total_qty_items_previous_month = nov_syn_agg_data['sum'].count()
    total_qty_overlap_items = overlap_data['Sku'].count()
    total_qty_new_items = total_qty_items_current_month - total_qty_overlap_items
    #print(total_qty_items, total_qty_overlap_items, total_qty_new_items)
    
    # Churn rate, growth rate, overall rate
    if total_qty_items_previous_month != 0:
        churn_rate = (total_qty_items_previous_month - total_qty_overlap_items) / total_qty_items_previous_month
        growth_rate = total_qty_new_items / total_qty_items_previous_month
        overall_rate = (total_qty_items_current_month-total_qty_items_previous_month) / total_qty_items_previous_month
    else:
        churn_rate = 0
        growth_rate = 0
        overall_rate = 0

    # File in Excel
    worksheet=workbook.add_worksheet(v)
    writer.sheets[v] = worksheet
    worksheet.write_string(0, 0, 'Top 10 overlapped items with negative effect: ')
    top10_negative.to_excel(writer,sheet_name=v,startrow=1 , startcol=0)   
    worksheet.write_string(13, 0, 'Top 10 overlapped items with positive effect: ')
    top10_positive.to_excel(writer,sheet_name=v,startrow=14, startcol=0)
    worksheet.write_string(26, 0, 'Top 10 sales in previous month: ')
    top10_sales_nov.to_excel(writer,sheet_name=v,startrow=27, startcol=0)
    worksheet.write_string(39, 0, 'Top 10 sales in current month: ')
    top10_sales_dec.to_excel(writer,sheet_name=v,startrow=40, startcol=0)
        
    # Integrate business information with each vendors
    business_data = [v, round(nov_sales,2), round(dec_sales,2), total_qty_items_previous_month, total_qty_items_current_month,                      total_qty_overlap_items, total_qty_new_items, round(churn_rate,2), round(growth_rate,2),                      round(overall_rate,2)]
    info_list.append(business_data)
    
# Create a file for business information
info_df = pd.DataFrame(info_list, columns=['Vendor', 'Sales in Previous month', 'Sales in Current month',                                            'Numbers of SKUs in Previous month', 'Numbers of SKUs in Current month',                                            'Overlapped SKUs', 'New SKUs', 'Churn rate', 'Growth rate', 'Overall rate'])
info_df = info_df.set_index('Vendor')
info_df.to_csv('output/analysis_result/Sales_by_Vendors/Sales_Info_by_AllVendors.csv')

# Save file for xlsx
writer.save()