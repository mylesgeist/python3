#!/usr/bin/python

import os
import csv
from datetime import datetime
import xlsxwriter

print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                        Starting the Amazon Error Report Script                                     #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("---------------------------------------------------------------------------------------------------------------")
print("")

prev_x4400 = 593
prev_x5461 = 2565
prev_x5002_8123 = 90
prev_x8541 = 6228
prev_x8542 = 1453
prev_x8560 = 293
prev_x8566 = 353
prev_x8567 = 498
prev_x8572 = 353
prev_x99001 = 993
prev_x17001_300404 = 592

# print("Creating Globals")
do_summary = True
do_amazon_fields = True
do_xall_errors = True
do_x4400 = True
do_x5461 = True
do_x5002_8123 = True
do_x8541 = True
do_x8542 = True
do_x8560 = True
do_x8566 = True
do_x8567 = True
do_x8572 = True
do_x99001 = True
do_x17001_300404 = True

do_asin = True
do_amazon_fields = True
do_amazon_fields_file = True
do_x8541_amazon_brand = True
do_x8541_amazon_manufacturer = True
do_x8541_amazon_product_type = True
do_x8541_amazon_part_number = True
do_x8541_amazon_item_name = True
do_x8541_amazon_model_number = True

do_99001_fields = True
do_age_range_description = True
do_apparel_size_class = True
do_apparel_size_system = True
do_bottoms_size_class = True
do_bottoms_size_system = True
do_headwear_size_class = True
do_headwear_size_system = True
do_shirt_size_class = True
do_shirt_size_system = True
do_target_gender = True

names_list = ['marcus']
csv_file = 'in-data.csv'
now = datetime.now()
temp_asin_list = []
current_date = now.strftime('%Y-%m-%d')

os.sep = """/"""
fix_script_dir = 'D:/Code-Projects/Python3/new_scripts/FixAmazonErrors/'
in_files_dir = str(fix_script_dir) + 'InFiles/'
in_process_dir = str(fix_script_dir) + 'InProcessFiles/'
sl_data_dir = str(fix_script_dir) + 'SLDataFiles/'
sl_check = int(0)
sl_check_list = []
sl_check_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
print("Checking to see if this feed was an SL feed")
for sl_check_row in sl_check_reader:
    if len(sl_check_row) >= 4:
        sl_check_list.append(sl_check_row)

for all_range in range(100):
    # for all_lists in sl_check_list:
    if sl_check_list[0]:
        if '-SL' in sl_check_list[0][1]:
            sl_check = sl_check + 1
            del sl_check_list[0]
        # print('')
        # print(sl_check_row)
sl_check_list = []
if sl_check >= 50:
    report_file_name = str(current_date) + '-AmazonErrors-SL.xlsx'
    x8541_report_file_name = str(current_date) + '-Amazon8541Fields-SL.xlsx'
    x99001_report_file_name = str(current_date) + '-Amazon99001Fields-SL.xlsx'
    print(' -- Creating the Amazon SL Error Report since this has SL data')
else:
    report_file_name = str(current_date) + '-AmazonErrors.xlsx'
    x8541_report_file_name = str(current_date) + '-Amazon8541Fields.xlsx'
    x99001_report_file_name = str(current_date) + '-Amazon99001Fields.xlsx'
    print(' -- Creating the Regular Amazon Error Report since this does not have SL data')

# exit(0)
amazon_report_path = 'D:/Code-Projects/Python3/amazon-report/Excel/'
# amazon_report_path = 'D:/Documents/Amazon/Weekly Error Report/'
absolute_file_name = amazon_report_path + report_file_name
x8541_absolute_file_name = amazon_report_path + x8541_report_file_name
x99001_absolute_file_name = amazon_report_path + x99001_report_file_name
errors_workbook = xlsxwriter.Workbook(absolute_file_name)
x8541_errors_workbook = xlsxwriter.Workbook(x8541_absolute_file_name)
x99001_errors_workbook = xlsxwriter.Workbook(x99001_absolute_file_name)
original_heading_list = ['original-record-number', 'sku', 'error-code', 'error-type', 'error-message']
error_code_list = ["4400", "5461", "5665", "6024", "6027", "6030", "8026", "8047", "8058", "8541", "8542", "8560", "8566",
                   "8567", "8572", "17002", "20014", "90041", "90057", "90202", "99001", "99010", "99022", "99038", "300060"]
if not do_xall_errors:
    do_xall_errors = True

# print("Creating the sheets for the excel file")
print("-------------------------------------------------------------")
print("")


def create_all_errors():
    global errors_workbook
    print("Creating the all errors sheet")
    all_errors_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    all_errors_list = []
    print("-- Populating the All Errors List")
    for all_row in all_errors_reader:
        if all_row not in all_errors_list:
            if len(all_row) >= 4:
                all_errors_list.append(all_row)
    print("-- Creating the All Errors Worksheet")
    all_errors_worksheet = errors_workbook.add_worksheet('All Errors')
    print("-- Creating the table in the All Errors Worksheet")
    all_errors_worksheet.add_table('A1:E' + str(len(all_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the All Errors Worksheet")
    all_errors_num = int(2)
    for i in all_errors_list:
        all_errors_worksheet.write_row('A' + str(all_errors_num) + ':E' + str(all_errors_num), all_errors_list[all_errors_num - 2])
        all_errors_num = all_errors_num + 1
    return all_errors_list


def create_4400_errors():
    global errors_workbook
    print("Creating the 4400 sheet")
    x4400_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x4400_errors_list = []
    print("-- Populating 4400 dictionary")
    for x4400_row in x4400_reader:
        if x4400_row not in x4400_errors_list:
            if len(x4400_row) >= 4:
                if x4400_row[2] == '4400':
                    x4400_errors_list.append(x4400_row)
    print("-- Creating the All 4400 Worksheet")
    x4400_errors_worksheet = errors_workbook.add_worksheet('4400')
    print("-- Creating the table in the 4400 Errors Worksheet")
    x4400_errors_worksheet.add_table('A1:E' + str(len(x4400_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 4400 Errors Worksheet")
    x4400_errors_num = int(2)
    for i in x4400_errors_list:
        x4400_errors_worksheet.write_row('A' + str(x4400_errors_num) + ':E' + str(x4400_errors_num), x4400_errors_list[x4400_errors_num - 2])
        x4400_errors_num = x4400_errors_num + 1
    return x4400_errors_list


def create_5461_errors():
    global errors_workbook
    print("Creating the 5461 sheet")
    x5461_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x5461_errors_list = []
    print("-- Populating 5461 dictionary")
    for x5461_row in x5461_reader:
        if x5461_row not in x5461_errors_list:
            if len(x5461_row) >= 4:
                if x5461_row[2] == '5461':
                    x5461_errors_list.append(x5461_row)
    print("-- Creating the All 5461 Worksheet")
    x5461_errors_worksheet = errors_workbook.add_worksheet('5461')
    print("-- Creating the table in the 5461 Errors Worksheet")
    x5461_errors_worksheet.add_table('A1:E' + str(len(x5461_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 5461 Errors Worksheet")
    x5461_errors_num = int(2)
    for i in x5461_errors_list:
        x5461_errors_worksheet.write_row('A' + str(x5461_errors_num) + ':E' + str(x5461_errors_num), x5461_errors_list[x5461_errors_num - 2])
        x5461_errors_num = x5461_errors_num + 1
    return x5461_errors_list


def create_5002_8123_errors():
    global errors_workbook
    print("Creating the 5002_8123 sheet")
    x5002_8123_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x5002_8123_errors_list = []
    print("-- Populating All 5002_8123 dictionary")
    for x5002_8123_row in x5002_8123_reader:
        if x5002_8123_row not in x5002_8123_errors_list:
            if len(x5002_8123_row) >= 4:
                if int(x5002_8123_row[2]) >= int('5002'):
                    if int(x5002_8123_row[2]) <= int('8123'):
                        if int(x5002_8123_row[2]) != int('5461'):
                            x5002_8123_errors_list.append(x5002_8123_row)
    print("-- Creating the All 5002_8123 Worksheet")
    x5002_8123_errors_worksheet = errors_workbook.add_worksheet('5002 - 8123')
    print("-- Creating the table in the 5002_8123 Errors Worksheet")
    x5002_8123_errors_worksheet.add_table('A1:E' + str(len(x5002_8123_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 5002_8123 Errors Worksheet")
    x5002_8123_errors_num = int(2)
    for i in x5002_8123_errors_list:
        x5002_8123_errors_worksheet.write_row('A' + str(x5002_8123_errors_num) + ':E' + str(x5002_8123_errors_num), x5002_8123_errors_list[x5002_8123_errors_num - 2])
        x5002_8123_errors_num = x5002_8123_errors_num + 1
    return x5002_8123_errors_list


def create_8541_errors():
    global do_asin
    global do_amazon_fields
    global do_x8541_amazon_brand
    global do_x8541_amazon_manufacturer
    global do_x8541_amazon_product_type
    global do_x8541_amazon_part_number
    global do_x8541_amazon_item_name
    global do_x8541_amazon_model_number
    global temp_asin_list
    global x8541_error_num
    global errors_workbook
    global x8541_errors_workbook
    global temp_asin_list
    global original_heading_list
    x8541_errors_list = []
    column_count = int(0)
    original_heading_list = ['original-record-number', 'sku', 'error-code', 'error-type', 'error-message']
    print("Creating the 8541 sheet")
    if do_amazon_fields:
        x8541_error_num = int(1)
        x8541_dict = {}
        x8541_temp_dict = {}
        fields_list = []
        x8541_reader = csv.reader(open(csv_file, mode='r', encoding='utf8'), delimiter='\t')
        print("-- Populating All errors dictionary")
        for x8541_row in x8541_reader:
            if len(x8541_row) >= 4:
                if '8541' in x8541_row[2]:
                    x8541_key = x8541_row[0]
                    x8541_temp_dict[x8541_key] = x8541_row[0:]
        x8541_error_num = int(1)
        print('-- Populating the dictionary with the proper fields')
        for i in x8541_temp_dict.values():
            x8541_errors_row = dict(zip(original_heading_list, i))
            x8541_errors_row['fields'] = ''
            if do_asin:
                x8541_errors_row['asin'] = ''
            if do_x8541_amazon_brand:
                x8541_errors_row['company_brand'] = ''
                x8541_errors_row['amazon_brand'] = ''
            if do_x8541_amazon_manufacturer:
                x8541_errors_row['company_manufacturer'] = ''
                x8541_errors_row['amazon_manufacturer'] = ''
            if do_x8541_amazon_product_type:
                x8541_errors_row['company_product_type'] = ''
                x8541_errors_row['amazon_product_type'] = ''
            if do_x8541_amazon_part_number:
                x8541_errors_row['company_part_number'] = ''
                x8541_errors_row['amazon_part_number'] = ''
            if do_x8541_amazon_item_name:
                x8541_errors_row['company_item_name'] = ''
                x8541_errors_row['amazon_item_name'] = ''
            if do_x8541_amazon_model_number:
                x8541_errors_row['company_model_number'] = ''
                x8541_errors_row['amazon_model_number'] = ''
            x8541_dict[x8541_error_num] = {x8541_error_num: x8541_errors_row}
            x8541_error_num = x8541_error_num + 1
        x8541_error_num = int(1)
        print('-- -- Seperating the fields from the error message')
        for keys, values in x8541_dict.items():
            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['error-message']
            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'].split("""already in the Amazon catalog: """, 10)
            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'][-1].strip()
            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'].split(""". If this is the right ASIN for your product""", 1)
            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'][0].strip()
            # print(x8541_dict[x8541_error_num][x8541_error_num]['fields'])
            x8541_error_num = x8541_error_num + 1
        x8541_error_num = int(1)
        if do_x8541_amazon_brand:
            fields_list.append('brand')
        if do_x8541_amazon_manufacturer:
            fields_list.append('manufacturer')
        if do_x8541_amazon_product_type:
            fields_list.append('product_type')
        if do_x8541_amazon_part_number:
            fields_list.append('part_number')
        if do_x8541_amazon_item_name:
            fields_list.append('item_name')
        if do_x8541_amazon_model_number:
            fields_list.append('model_number')
        print('-- -- Iterating through the fields field and seperating the Azmazon and company fields data')
        print('      This may take a while')
        x8541_error_num = int(1)
        for keys, values in x8541_dict.items():
            x8541_dict_runs = int(1)
            x8541_temp_list = []
            if do_amazon_fields:
                while x8541_dict_runs <= 10:
                    for r in fields_list:
                        current_field = r
                        if str(x8541_dict[x8541_error_num][x8541_error_num]['fields']).startswith(str(current_field)):
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['fields']
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field].split(""" / Amazon: '""", 1)[-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field].split("""')""", 1)
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field][0].strip()
                            if str(x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field]).startswith('"') and  str(x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field]).endswith('"'):
                                x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field][1:-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field].strip()
                            if str(x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field]).startswith("'") and str(x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field]).endswith("'"):
                                x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field][1:-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_' + current_field].strip()
                            x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['fields']
                            x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field].split('Merchant: ', 1)[-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field].split(' / Amazon:', 1)[0]
                            if str(x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field]).startswith('"') and str(x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field]).endswith('"'):
                                x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field][1:-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field].strip()
                            if str(x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field]).startswith("'") and str(x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field]).endswith("'"):
                                x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field][1:-1]
                            x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field] = x8541_dict[x8541_error_num][x8541_error_num]['company_' + current_field].strip()
                            x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'].split("""'), """, 1)
                            if isinstance(x8541_dict[x8541_error_num][x8541_error_num]['fields'], list):
                                if len(x8541_dict[x8541_error_num][x8541_error_num]['fields']) > 1:
                                    x8541_dict[x8541_error_num][x8541_error_num]['fields'] = x8541_dict[x8541_error_num][x8541_error_num]['fields'][-1]
                                else:
                                    x8541_dict[x8541_error_num][x8541_error_num]['fields'] = ''
                    x8541_dict_runs = x8541_dict_runs + 1
            if do_asin:
                x8541_dict[x8541_error_num][x8541_error_num]['asin'] = x8541_dict[x8541_error_num][x8541_error_num]['error-message']
                x8541_dict[x8541_error_num][x8541_error_num]['asin'] = x8541_dict[x8541_error_num][x8541_error_num]['asin'].split("""data provided matches ASIN """, 1)
                x8541_dict[x8541_error_num][x8541_error_num]['asin'] = x8541_dict[x8541_error_num][x8541_error_num]['asin'][-1].strip()
                x8541_dict[x8541_error_num][x8541_error_num]['asin'] = x8541_dict[x8541_error_num][x8541_error_num]['asin'].split(""", but the following data""", 1)
                x8541_dict[x8541_error_num][x8541_error_num]['asin'] = x8541_dict[x8541_error_num][x8541_error_num]['asin'][0].strip()
            x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['original-record-number'])
            x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['sku'])
            x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['error-code'])
            x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['error-type'])
            if do_asin:
                x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['asin'])
            if do_amazon_fields:
                if do_x8541_amazon_brand:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['company_brand'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['company_brand'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['company_brand'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['acompany_brand'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['company_brand'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand'] = x8541_dict[x8541_error_num][x8541_error_num]['company_brand'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_brand'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand'])
                if do_x8541_amazon_manufacturer:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'])
                if do_x8541_amazon_product_type:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_product_type'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'])
                if do_x8541_amazon_part_number:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_part_number'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'])
                if do_x8541_amazon_item_name:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_item_name'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'])
                if do_x8541_amazon_model_number:
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'].replace("&quot;", "\u0022")
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'].replace("&amp;", "\u0026")  # STRING TO " and "
                    # x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'].replace("&quot;", "\u0022")
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'].replace("&amp;", "\u0026")  # STRING TO " and "
                    x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'] = x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'].replace("&apos;", "\u0027")  # TO APOSTROPHE
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['company_model_number'])
                    x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'])
            x8541_temp_list.append(x8541_dict[x8541_error_num][x8541_error_num]['error-message'])
            x8541_error_num = x8541_error_num + int(1)
            x8541_errors_list.append(x8541_temp_list)
        x8541_error_num = int(1)
    table_columns = {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}]}
    if do_asin:
        column_count = column_count + 1
        table_columns['columns'].append({'header': 'ASIN'})
    if do_x8541_amazon_brand:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Brand'})
        table_columns['columns'].append({'header': 'Amazon Brand'})
    if do_x8541_amazon_manufacturer:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Manufacturer'})
        table_columns['columns'].append({'header': 'Amazon Manufacturer'})
    if do_x8541_amazon_product_type:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Product Type'})
        table_columns['columns'].append({'header': 'Amazon Product Type'})
    if do_x8541_amazon_part_number:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Part Number'})
        table_columns['columns'].append({'header': 'Amazon Part Number'})
    if do_x8541_amazon_item_name:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Item Name'})
        table_columns['columns'].append({'header': 'Amazon Item Name'})
    if do_x8541_amazon_model_number:
        column_count = column_count + 2
        table_columns['columns'].append({'header': 'company Model Number'})
        table_columns['columns'].append({'header': 'Amazon Model Number'})
    table_columns['columns'].append({'header': 'Error Message'})
    col_letter = str('E')
    if column_count == 1:
        col_letter = 'F'
    if column_count == 2:
        col_letter = 'G'
    if column_count == 3:
        col_letter = 'H'
    if column_count == 4:
        col_letter = 'I'
    if column_count == 5:
        col_letter = 'J'
    if column_count == 6:
        col_letter = 'K'
    if column_count == 7:
        col_letter = 'L'
    if column_count == 8:
        col_letter = 'M'
    if column_count == 9:
        col_letter = 'N'
    if column_count == 10:
        col_letter = 'O'
    if column_count == 11:
        col_letter = 'P'
    if column_count == 12:
        col_letter = 'Q'
    if column_count == 13:
        col_letter = 'R'
    if not do_amazon_fields:
        table_columns = {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]}
        x8541_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
        print("-- Populating 8541 list")
        for x8541_row in x8541_reader:
            if x8541_row[2] == '8541':
                x8541_errors_list.append(x8541_row)
        print("-- Creating the All 8541 Worksheet")
    x8541_errors_worksheet = errors_workbook.add_worksheet('8541')
    print("-- Creating the table in the 8541 Errors Worksheet")
    x8541_errors_worksheet.add_table('A1:' + col_letter + str(len(x8541_errors_list) + 1), table_columns)
    x8541_error_num = int(1)
    for lists in x8541_errors_list:
        x8541_errors_worksheet.write_row('A' + str(x8541_error_num + 1) + col_letter + str(x8541_error_num), x8541_errors_list[x8541_error_num - 1])
        x8541_error_num = x8541_error_num + 1
    if do_amazon_fields_file:
        if do_asin:
            x8541_asin_list = []
            x8541_asin_columns = {'columns': [{'header': 'SKU'}, {'header': 'ASIN'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['asin']) >= 4:
                    x8541_asin_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['asin'], 'Yes', 'Yes']
                    x8541_asin_list.append(x8541_asin_row)
                x8541_error_num = x8541_error_num + 1
            x8541_asin_worksheet = x8541_errors_workbook.add_worksheet('ASIN')
            x8541_asin_worksheet.add_table('A1:D' + str(len(x8541_asin_list) + 1), x8541_asin_columns)
            x8541_asin_error_num = int(1)
            for lists in x8541_asin_list:
                x8541_asin_worksheet.write_row('A' + str(x8541_asin_error_num + 1) + ':D' + str(x8541_asin_error_num), x8541_asin_list[x8541_asin_error_num - 1])
                x8541_asin_error_num = x8541_asin_error_num + 1

        if do_x8541_amazon_brand:
            x8541_brand_list = []
            broken_brand_list = []
            x8541_brand_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Brand'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_brand']) >= 1:
                    x8541_brand_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_brand'], 'Yes', 'Yes']
                    x8541_brand_list.append(x8541_brand_row)
                x8541_error_num = x8541_error_num + 1
            x8541_brand_worksheet = x8541_errors_workbook.add_worksheet('Amazon Brand')
            x8541_brand_worksheet.add_table('A1:D' + str(len(x8541_brand_list) + 1), x8541_brand_columns)
            x8541_brand_error_num = int(1)
            for brands_lists in x8541_brand_list:
                x8541_brand_worksheet.write_row('A' + str(x8541_brand_error_num + 1) + ':D' + str(x8541_brand_error_num), x8541_brand_list[x8541_brand_error_num - 1])
                x8541_brand_error_num = x8541_brand_error_num + 1

        if do_x8541_amazon_manufacturer:
            x8541_manufacturer_list = []
            broken_manufacturer_list = []
            x8541_manufacturer_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Manufacturer'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_manufacturer']) >= 1:
                    x8541_manufacturer_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_manufacturer'], 'Yes', 'Yes']
                    x8541_manufacturer_list.append(x8541_manufacturer_row)
                x8541_error_num = x8541_error_num + 1
            x8541_manufacturer_worksheet = x8541_errors_workbook.add_worksheet('Amazon Manufacturer')
            x8541_manufacturer_worksheet.add_table('A1:D' + str(len(x8541_manufacturer_list) + 1), x8541_manufacturer_columns)
            x8541_manufacturer_error_num = int(1)
            for manufacturers_lists in x8541_manufacturer_list:
                x8541_manufacturer_worksheet.write_row('A' + str(x8541_manufacturer_error_num + 1) + ':D' + str(x8541_manufacturer_error_num), x8541_manufacturer_list[x8541_manufacturer_error_num - 1])
                x8541_manufacturer_error_num = x8541_manufacturer_error_num + 1

        if do_x8541_amazon_product_type:
            x8541_type_list = []
            broken_type_list = []
            x8541_type_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Product Type'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_product_type']) >= 1:
                    x8541_type_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_product_type'], 'Yes', 'Yes']
                    # print(x8541_type_row)
                    x8541_type_list.append(x8541_type_row)
                x8541_error_num = x8541_error_num + 1
            # print(x8541_type_list)
            x8541_type_worksheet = x8541_errors_workbook.add_worksheet('Amazon Product Type')
            x8541_type_worksheet.add_table('A1:D' + str(len(x8541_type_list) + 1), x8541_type_columns)
            x8541_type_error_num = int(1)
            for types_lists in x8541_type_list:
                x8541_type_worksheet.write_row('A' + str(x8541_type_error_num + 1) + ':D' + str(x8541_type_error_num), x8541_type_list[x8541_type_error_num - 1])
                x8541_type_error_num = x8541_type_error_num + 1

        if do_x8541_amazon_part_number:
            x8541_part_list = []
            broken_part_list = []
            x8541_part_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Part Number'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_part_number']) >= 1:
                    x8541_part_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_part_number'], 'Yes', 'Yes']
                    x8541_part_list.append(x8541_part_row)
                x8541_error_num = x8541_error_num + 1
            x8541_part_worksheet = x8541_errors_workbook.add_worksheet('Amazon Part Number')
            x8541_part_worksheet.add_table('A1:D' + str(len(x8541_part_list) + 1), x8541_part_columns)
            x8541_part_error_num = int(1)
            for parts_lists in x8541_part_list:
                x8541_part_worksheet.write_row('A' + str(x8541_part_error_num + 1) + ':D' + str(x8541_part_error_num), x8541_part_list[x8541_part_error_num - 1])
                x8541_part_error_num = x8541_part_error_num + 1

        if do_x8541_amazon_item_name:
            x8541_item_list = []
            broken_item_list = []
            x8541_item_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Item Name'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_item_name']) >= 1:
                    x8541_item_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_item_name'], 'Yes', 'Yes']
                    x8541_item_list.append(x8541_item_row)
                x8541_error_num = x8541_error_num + 1
            x8541_item_worksheet = x8541_errors_workbook.add_worksheet('Amazon Item Name')
            x8541_item_worksheet.add_table('A1:D' + str(len(x8541_item_list) + 1), x8541_item_columns)
            x8541_item_error_num = int(1)
            for items_lists in x8541_item_list:
                x8541_item_worksheet.write_row('A' + str(x8541_item_error_num + 1) + ':D' + str(x8541_item_error_num), x8541_item_list[x8541_item_error_num - 1])
                x8541_item_error_num = x8541_item_error_num + 1

        if do_x8541_amazon_model_number:
            x8541_model_list = []
            broken_model_list = []
            x8541_model_columns = {'columns': [{'header': 'SKU'}, {'header': 'Amazon Model Number'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x8541_error_num = 1
            for keys, values in x8541_dict.items():
                if len(x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number']) >= 1 or len(x8541_dict[x8541_error_num][x8541_error_num]['company_model_number']) >= 1:
                    x8541_model_row = [x8541_dict[x8541_error_num][x8541_error_num]['sku'], x8541_dict[x8541_error_num][x8541_error_num]['amazon_model_number'], 'Yes', 'Yes']
                    x8541_model_list.append(x8541_model_row)
                x8541_error_num = x8541_error_num + 1
            x8541_model_worksheet = x8541_errors_workbook.add_worksheet('Amazon Model Number')
            x8541_model_worksheet.add_table('A1:D' + str(len(x8541_model_list) + 1), x8541_model_columns)
            x8541_model_error_num = int(1)
            for models_lists in x8541_model_list:
                x8541_model_worksheet.write_row('A' + str(x8541_model_error_num + 1) + ':D' + str(x8541_model_error_num), x8541_model_list[x8541_model_error_num - 1])
                x8541_model_error_num = x8541_model_error_num + 1

        x8541_errors_workbook.close()
    return x8541_errors_list


def create_8542_errors():
    global errors_workbook
    global temp_asin_list
    print("Creating the 8542 sheet")
    x8542_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x8542_errors_temp_list = []
    x8542_errors_list = []
    print("-- Populating All 8542 list")
    for x8542_row in x8542_reader:
        if x8542_row not in x8542_errors_temp_list:
            if len(x8542_row) >= 4:
                if x8542_row[2] == '8542':
                    x8542_errors_temp_list.append(x8542_row)
    print("-- Creating the All 8542 Worksheet")
    x8542_errors_worksheet = errors_workbook.add_worksheet('8542')
    print("-- Creating the table in the 8542 Errors Worksheet")
    if do_asin:
        table_columns = {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'ASIN'}, {'header': 'Error Message'}]}
    if not do_asin:
        table_columns = {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]}
    print("-- Populating the table in the 8542 Errors Worksheet")
    # print(x8542_errors_list)
    if do_asin:
        print("-- Adding the ASIN Fields to the 8542 sheet")
        x8542_error_num = int(2)
        while x8542_error_num <= len(x8542_errors_temp_list):
            # try:
            for lists in x8542_errors_temp_list:
                if lists[0] and lists[1] and lists[2] and lists[3] and lists[4]:
                    temp_asin_list = [lists[0], lists[1], lists[2], lists[3], lists[4], lists[4]]
                    temp_asin_list[4] = temp_asin_list[4].split("""provided correspond to the ASIN  """, 1)
                    temp_asin_list[4] = temp_asin_list[4][-1].strip()
                    temp_asin_list[4] = temp_asin_list[4].split(""", but some information contradicts""", 1)
                    temp_asin_list[4] = temp_asin_list[4][0].strip()
                    x8542_errors_list.append(temp_asin_list)
                    x8542_error_num = x8542_error_num + 1
    x8542_error_num = int(2)
    if do_asin:
        x8542_errors_worksheet.add_table('A1:F' + str(len(x8542_errors_list) + 1), table_columns)
    if not do_asin:
        x8542_errors_worksheet.add_table('A1:E' + str(len(x8542_errors_list) + 1), table_columns)
    for lists in x8542_errors_list:
        x8542_errors_worksheet.write_row('A' + str(x8542_error_num) + ':E' + str(x8542_error_num), x8542_errors_list[x8542_error_num - 2])
        x8542_error_num = x8542_error_num + 1
    return x8542_errors_list


def create_8560_errors():
    global errors_workbook
    print("Creating the 8560 sheet")
    x8560_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x8560_errors_list = []
    print("-- Populating 8560 list")
    for x8560_row in x8560_reader:
        if x8560_row not in x8560_errors_list:
            if len(x8560_row) >= 4:
                if x8560_row[2] == '8560':
                    x8560_errors_list.append(x8560_row)
    print("-- Creating the All 8560 Worksheet")
    x8560_errors_worksheet = errors_workbook.add_worksheet('8560')
    print("-- Creating the table in the 8560 Errors Worksheet")
    x8560_errors_worksheet.add_table('A1:E' + str(len(x8560_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 8560 Errors Worksheet")
    x8560_errors_num = int(2)
    # print(all_errors_list)
    for i in x8560_errors_list:
        x8560_errors_worksheet.write_row('A' + str(x8560_errors_num) + ':E' + str(x8560_errors_num), x8560_errors_list[x8560_errors_num - 2])
        x8560_errors_num = x8560_errors_num + 1
    return x8560_errors_list


def create_8566_errors():
    global errors_workbook
    print("Creating the 8566 sheet")
    x8566_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x8566_errors_list = []
    print("-- Populating 8566 list")
    for x8566_row in x8566_reader:
        if x8566_row not in x8566_errors_list:
            if len(x8566_row) >= 4:
                if x8566_row[2] == '8566':
                    x8566_errors_list.append(x8566_row)
    print("-- Creating the All 8566 Worksheet")
    x8566_errors_worksheet = errors_workbook.add_worksheet('8566')
    print("-- Creating the table in the 8566 Errors Worksheet")
    x8566_errors_worksheet.add_table('A1:E' + str(len(x8566_errors_list) + 1), {'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'}, {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 8566 Errors Worksheet")
    x8566_errors_num = int(2)
    # print(all_errors_list)
    for i in x8566_errors_list:
        x8566_errors_worksheet.write_row('A' + str(x8566_errors_num) + ':E' + str(x8566_errors_num),  x8566_errors_list[x8566_errors_num - 2])
        x8566_errors_num = x8566_errors_num + 1
    return x8566_errors_list


def create_8567_errors():
    global errors_workbook
    print("Creating the 8567 sheet")
    x8567_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x8567_errors_list = []
    print("-- Populating 8567 list")
    for x8567_row in x8567_reader:
        if x8567_row not in x8567_errors_list:
            if len(x8567_row) >= 4:
                if x8567_row[2] == '8567':
                    x8567_errors_list.append(x8567_row)
    print("-- Creating the All 8566 Worksheet")
    x8567_errors_worksheet = errors_workbook.add_worksheet('8567')
    print("-- Creating the table in the 8566 Errors Worksheet")
    x8567_errors_worksheet.add_table('A1:E' + str(len(x8567_errors_list) + 1), {
        'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'},
                    {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 8567 Errors Worksheet")
    x8567_errors_num = int(2)
    # print(all_errors_list)
    for i in x8567_errors_list:
        x8567_errors_worksheet.write_row('A' + str(x8567_errors_num) + ':E' + str(x8567_errors_num), x8567_errors_list[x8567_errors_num - 2])
        x8567_errors_num = x8567_errors_num + 1
    return x8567_errors_list


def create_8572_errors():
    global errors_workbook
    print("Creating the 8572 sheet")
    x8572_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x8572_errors_list = []
    print("-- Populating 8567 list")
    for x8572_row in x8572_reader:
        if x8572_row not in x8572_errors_list:
            if len(x8572_row) >= 4:
                if x8572_row[2] == '8572':
                    x8572_errors_list.append(x8572_row)
    print("-- Creating the All 8566 Worksheet")
    x8572_errors_worksheet = errors_workbook.add_worksheet('8572')
    print("-- Creating the table in the 8572 Errors Worksheet")
    x8572_errors_worksheet.add_table('A1:E' + str(len(x8572_errors_list) + 1), {
        'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'},
                    {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 8567 Errors Worksheet")
    x8572_errors_num = int(2)
    # print(all_errors_list)
    for i in x8572_errors_list:
        x8572_errors_worksheet.write_row('A' + str(x8572_errors_num) + ':E' + str(x8572_errors_num), x8572_errors_list[x8572_errors_num - 2])
        x8572_errors_num = x8572_errors_num + 1
    return x8572_errors_list


def create_99001_errors():
    global errors_workbook
    global x99001_errors_workbook
    print("Creating the 99001 sheet")
    x99001_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x99001_errors_list = []
    print("-- Populating 99001 list")
    for x99001_row in x99001_reader:
        if x99001_row not in x99001_errors_list:
            if len(x99001_row) >= 4:
                if x99001_row[2] == '99001':
                    x99001_errors_list.append(x99001_row)
    print("-- Creating the All 8566 Worksheet")
    x99001_errors_worksheet = errors_workbook.add_worksheet('99001')
    print("-- Creating the table in the 99001 Errors Worksheet")
    x99001_errors_worksheet.add_table('A1:E' + str(len(x99001_errors_list) + 1), {
        'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'},
                    {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 99001 Errors Worksheet")
    x99001_errors_num = int(2)
    # print(all_errors_list)
    for i in x99001_errors_list:
        x99001_errors_worksheet.write_row('A' + str(x99001_errors_num) + ':E' + str(x99001_errors_num), x99001_errors_list[x99001_errors_num - 2])
        x99001_errors_num = x99001_errors_num + 1

    if do_99001_fields == True:
        x99001_dict = {}
        x99001_errors_num = int(1)
        for all_99001s in x99001_errors_list:
            x99001_dict[x99001_errors_num] = {'sku': all_99001s[1], 'error': all_99001s[4]}
            for keys, values in x99001_dict.items():
                if "Age Range Description" in values['error']:
                    values['field'] = "Age Range Description"
                    values['value'] = 'Adult'
                elif "Apparel Size Class" in values['error']:
                    values['field'] = "Apparel Size Class"
                    values['value'] = 'Alpha'
                elif "Apparel Size System" in values['error']:
                    values['field'] = "Apparel Size System"
                    values['value'] = 'US'
                elif "bottoms_size_class" in values['error']:
                    values['field'] = "bottoms_size_class"
                    values['value'] = 'Alpha'
                elif "bottoms_size_system" in values['error']:
                    values['field'] = "bottoms_size_system"
                    values['value'] = 'US'
                elif "headwear_size_class" in values['error']:
                    values['field'] = "headwear_size_class"
                    values['value'] = 'Alpha'
                elif "headwear_size_system" in values['error']:
                    values['field'] = "headwear_size_system"
                    values['value'] = 'US'
                elif "shirt_size_class" in values['error']:
                    values['field'] = "shirt_size_class"
                    values['value'] = 'Alpha'
                elif "shirt_size_system" in values['error']:
                    values['field'] = "shirt_size_system"
                    values['value'] = 'US'
                elif "Target Gender" in values['error']:
                    values['field'] = "Target Gender"
                    values['value'] = 'Unisex'
                else:
                    values['field'] = ''
                    values['value'] = ''

            x99001_errors_num = x99001_errors_num + 1

        if do_age_range_description:
            x99001_age_list = []
            broken_age_list = []
            x99001_age_columns = {'columns': [{'header': 'SKU'}, {'header': 'Age Range Description'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "Age Range Description":
                    x99001_age_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_age_list.append(x99001_age_row)
                x99001_error_num = x99001_error_num + 1
            x99001_age_worksheet = x99001_errors_workbook.add_worksheet('Age Range Description')
            x99001_age_worksheet.add_table('A1:D' + str(len(x99001_age_list) + 1), x99001_age_columns)
            x99001_age_error_num = int(1)
            for ages_lists in x99001_age_list:
                x99001_age_worksheet.write_row('A' + str(x99001_age_error_num + 1) + ':D' + str(x99001_age_error_num), x99001_age_list[x99001_age_error_num - 1])
                x99001_age_error_num = x99001_age_error_num + 1

        if do_apparel_size_class:
            x99001_apparel_class_list = []
            broken_apparel_class_list = []
            x99001_apparel_class_columns = {'columns': [{'header': 'SKU'}, {'header': 'Apparel Size Class'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "Apparel Size Class":
                    x99001_apparel_class_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_apparel_class_list.append(x99001_apparel_class_row)
                x99001_error_num = x99001_error_num + 1
            x99001_apparel_class_worksheet = x99001_errors_workbook.add_worksheet('Apparel Size Class')
            x99001_apparel_class_worksheet.add_table('A1:D' + str(len(x99001_apparel_class_list) + 1), x99001_apparel_class_columns)
            x99001_apparel_class_error_num = int(1)
            for apparel_classes_lists in x99001_apparel_class_list:
                x99001_apparel_class_worksheet.write_row(
                    'A' + str(x99001_apparel_class_error_num + 1) + ':D' + str(x99001_apparel_class_error_num),
                    x99001_apparel_class_list[x99001_apparel_class_error_num - 1])
                x99001_apparel_class_error_num = x99001_apparel_class_error_num + 1

        if do_apparel_size_system:
            x99001_apparel_system_list = []
            broken_apparel_system_list = []
            x99001_apparel_system_columns = {'columns': [{'header': 'SKU'}, {'header': 'Apparel Size System'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "Apparel Size System":
                    x99001_apparel_system_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_apparel_system_list.append(x99001_apparel_system_row)
                x99001_error_num = x99001_error_num + 1
            x99001_apparel_system_worksheet = x99001_errors_workbook.add_worksheet('Apparel Size System')
            x99001_apparel_system_worksheet.add_table('A1:D' + str(len(x99001_apparel_system_list) + 1), x99001_apparel_system_columns)
            x99001_apparel_system_error_num = int(1)
            for apparel_systems_lists in x99001_apparel_system_list:
                x99001_apparel_system_worksheet.write_row('A' + str(x99001_apparel_system_error_num + 1) + ':D' + str(x99001_apparel_system_error_num), x99001_apparel_system_list[x99001_apparel_system_error_num - 1])
                x99001_apparel_system_error_num = x99001_apparel_system_error_num + 1

        if do_bottoms_size_class:
            x99001_bottoms_class_list = []
            broken_bottoms_class_list = []
            x99001_bottoms_class_columns = {'columns': [{'header': 'SKU'}, {'header': 'Bottoms Size Class'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "bottoms_size_class":
                    x99001_bottoms_class_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_bottoms_class_list.append(x99001_bottoms_class_row)
                x99001_error_num = x99001_error_num + 1
            x99001_bottoms_class_worksheet = x99001_errors_workbook.add_worksheet('Bottoms Size Class')
            x99001_bottoms_class_worksheet.add_table('A1:D' + str(len(x99001_bottoms_class_list) + 1), x99001_bottoms_class_columns)
            x99001_bottoms_class_error_num = int(1)
            for bottoms_classes_lists in x99001_bottoms_class_list:
                x99001_bottoms_class_worksheet.write_row('A' + str(x99001_bottoms_class_error_num + 1) + ':D' + str(x99001_bottoms_class_error_num), x99001_bottoms_class_list[x99001_bottoms_class_error_num - 1])
                x99001_bottoms_class_error_num = x99001_bottoms_class_error_num + 1

        if do_bottoms_size_system:
            x99001_bottoms_system_list = []
            broken_bottoms_system_list = []
            x99001_bottoms_system_columns = {'columns': [{'header': 'SKU'}, {'header': 'Bottoms Size System'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "bottoms_size_system":
                    x99001_bottoms_system_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_bottoms_system_list.append(x99001_bottoms_system_row)
                x99001_error_num = x99001_error_num + 1
            x99001_bottoms_system_worksheet = x99001_errors_workbook.add_worksheet('Bottoms Size System')
            x99001_bottoms_system_worksheet.add_table('A1:D' + str(len(x99001_bottoms_system_list) + 1), x99001_bottoms_system_columns)
            x99001_bottoms_system_error_num = int(1)
            for bottoms_systems_lists in x99001_bottoms_system_list:
                x99001_bottoms_system_worksheet.write_row('A' + str(x99001_bottoms_system_error_num + 1) + ':D' + str(x99001_bottoms_system_error_num), x99001_bottoms_system_list[x99001_bottoms_system_error_num - 1])
                x99001_bottoms_system_error_num = x99001_bottoms_system_error_num + 1

        if do_headwear_size_class:
            x99001_headwear_class_list = []
            broken_headwear_class_list = []
            x99001_headwear_class_columns = {'columns': [{'header': 'SKU'}, {'header': 'Headwear Size Class'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "headwear_size_class":
                    x99001_headwear_class_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_headwear_class_list.append(x99001_headwear_class_row)
                x99001_error_num = x99001_error_num + 1
            x99001_headwear_class_worksheet = x99001_errors_workbook.add_worksheet('Headwear Size Class')
            x99001_headwear_class_worksheet.add_table('A1:D' + str(len(x99001_headwear_class_list) + 1), x99001_headwear_class_columns)
            x99001_headwear_class_error_num = int(1)
            for headwear_classes_lists in x99001_headwear_class_list:
                x99001_headwear_class_worksheet.write_row('A' + str(x99001_headwear_class_error_num + 1) + ':D' + str(x99001_headwear_class_error_num), x99001_headwear_class_list[x99001_headwear_class_error_num - 1])
                x99001_headwear_class_error_num = x99001_headwear_class_error_num + 1

        if do_headwear_size_system:
            x99001_headwear_system_list = []
            broken_headwear_system_list = []
            x99001_headwear_system_columns = {'columns': [{'header': 'SKU'}, {'header': 'Headwear Size System'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "headwear_size_system":
                    x99001_headwear_system_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_headwear_system_list.append(x99001_headwear_system_row)
                x99001_error_num = x99001_error_num + 1
            x99001_headwear_system_worksheet = x99001_errors_workbook.add_worksheet('Headwear Size System')
            x99001_headwear_system_worksheet.add_table('A1:D' + str(len(x99001_headwear_system_list) + 1), x99001_headwear_system_columns)
            x99001_headwear_system_error_num = int(1)
            for headwear_systems_lists in x99001_headwear_system_list:
                x99001_headwear_system_worksheet.write_row('A' + str(x99001_headwear_system_error_num + 1) + ':D' + str(x99001_headwear_system_error_num), x99001_headwear_system_list[x99001_headwear_system_error_num - 1])
                x99001_headwear_system_error_num = x99001_headwear_system_error_num + 1

        if do_shirt_size_class:
            x99001_shirt_class_list = []
            broken_shirt_class_list = []
            x99001_shirt_class_columns = {'columns': [{'header': 'SKU'}, {'header': 'Shirt Size Class'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "shirt_size_class":
                    x99001_shirt_class_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_shirt_class_list.append(x99001_shirt_class_row)
                x99001_error_num = x99001_error_num + 1
            x99001_shirt_class_worksheet = x99001_errors_workbook.add_worksheet('Shirt Size Class')
            x99001_shirt_class_worksheet.add_table('A1:D' + str(len(x99001_shirt_class_list) + 1), x99001_shirt_class_columns)
            x99001_shirt_class_error_num = int(1)
            for shirt_classes_lists in x99001_shirt_class_list:
                x99001_shirt_class_worksheet.write_row('A' + str(x99001_shirt_class_error_num + 1) + ':D' + str(x99001_shirt_class_error_num), x99001_shirt_class_list[x99001_shirt_class_error_num - 1])
                x99001_shirt_class_error_num = x99001_shirt_class_error_num + 1

        if do_shirt_size_system:
            x99001_shirt_system_list = []
            broken_shirt_system_list = []
            x99001_shirt_system_columns = {'columns': [{'header': 'SKU'}, {'header': 'Shirt Size System'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "shirt_size_system":
                    x99001_shirt_system_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_shirt_system_list.append(x99001_shirt_system_row)
                x99001_error_num = x99001_error_num + 1
            x99001_shirt_system_worksheet = x99001_errors_workbook.add_worksheet('Shirt Size System')
            x99001_shirt_system_worksheet.add_table('A1:D' + str(len(x99001_shirt_system_list) + 1), x99001_shirt_system_columns)
            x99001_shirt_system_error_num = int(1)
            for shirt_systems_lists in x99001_shirt_system_list:
                x99001_shirt_system_worksheet.write_row('A' + str(x99001_shirt_system_error_num + 1) + ':D' + str(x99001_shirt_system_error_num), x99001_shirt_system_list[x99001_shirt_system_error_num - 1])
                x99001_shirt_system_error_num = x99001_shirt_system_error_num + 1

        if do_target_gender:
            x99001_gender_list = []
            broken_gender_list = []
            x99001_gender_columns = {'columns': [{'header': 'SKU'}, {'header': 'Target Gender'}, {'header': 'Approved?'}, {'header': 'Import to WCE?'}]}
            x99001_error_num = 1
            for keys, values in x99001_dict.items():
                if values['field'] == "Target Gender":
                    x99001_gender_row = [values['sku'], values['value'], 'Yes', 'Yes']
                    x99001_gender_list.append(x99001_gender_row)
                x99001_error_num = x99001_error_num + 1
            x99001_gender_worksheet = x99001_errors_workbook.add_worksheet('Target Gender')
            x99001_gender_worksheet.add_table('A1:D' + str(len(x99001_gender_list) + 1), x99001_gender_columns)
            x99001_gender_error_num = int(1)
            for genders_lists in x99001_gender_list:
                x99001_gender_worksheet.write_row(
                    'A' + str(x99001_gender_error_num + 1) + ':B' + str(x99001_gender_error_num),
                    x99001_gender_list[x99001_gender_error_num - 1])
                x99001_gender_error_num = x99001_gender_error_num + 1
    x99001_errors_workbook.close()
    return x99001_errors_list


def create_17001_300404_errors():
    global errors_workbook
    print("Creating the 17001_300404 sheet")
    x17001_300404_reader = csv.reader(open(csv_file, mode='r', encoding='utf8', errors="surrogateescape"), delimiter='\t')
    x17001_300404_errors_list = []
    print("-- Populating 17001_300404 list")
    for x17001_300404_row in x17001_300404_reader:
        if x17001_300404_row not in x17001_300404_errors_list:
            if len(x17001_300404_row) >= 4:
                if int(x17001_300404_row[2]) >= int('17001'):
                    if int(x17001_300404_row[2]) != int('99001'):
                        x17001_300404_errors_list.append(x17001_300404_row)
    print("-- Creating the All 17001_300404 Worksheet")
    x17001_300404_errors_worksheet = errors_workbook.add_worksheet('17001_300404')
    print("-- Creating the table in the 17001_300404 Errors Worksheet")
    x17001_300404_errors_worksheet.add_table('A1:E' + str(len(x17001_300404_errors_list) + 1), {
        'columns': [{'header': 'Original Record Number'}, {'header': 'SKU'}, {'header': 'Error Code'},
                    {'header': 'Error Type'}, {'header': 'Error Message'}]})
    print("-- Populating the table in the 17001_300404 Errors Worksheet")
    x17001_300404_errors_num = int(2)
    for i in x17001_300404_errors_list:
        x17001_300404_errors_worksheet.write_row('A' + str(x17001_300404_errors_num) + ':E' + str(x17001_300404_errors_num), x17001_300404_errors_list[x17001_300404_errors_num - 2])
        x17001_300404_errors_num = x17001_300404_errors_num + 1
    return x17001_300404_errors_list


def create_summary():
    global summary_list
    summary_heading_list = ['Error Code', 'Percentage of Errors', "This Week's Errors", "Previous Errors", "Changes"]
    summary_columns = {
        'columns': [{'header': 'Error Code'}, {'header': 'Percentage of Errors'}, {'header': "This Week's Errors"}, {'header': 'Previous Errors'}, {'header': 'Changes'}]}
    summary_label = ['Summary:', '', '', '', '']
    totals_summary_list = ['Totals', '', tw_totals, prev_totals, int(tw_totals) - int(prev_totals)]
    print("Creating the Summary Worksheet")
    summary_worksheet = errors_workbook.add_worksheet('Summary')
    print("-- Creating the table in the Summary Worksheet")
    summary_worksheet.add_table('B6:F' + str(len(summary_list) + 6), summary_columns)
    print("-- Populating the table in the Summary Worksheet")
    summary_num = int(2)
    summary_worksheet.write_row('B5:F5', summary_label)
    for i in summary_list:
        summary_worksheet.write_row('B' + str(summary_num + 5) + ':F' + str(summary_num + 5), summary_list[summary_num - 2])
        summary_num = summary_num + 1
    summary_worksheet.write_row("B19:F19", totals_summary_list)
    return summary_list


def percentage(part, whole):
    percentage = round(100 * float(part) / float(whole), 2)
    return str(percentage) + "%"


summary_list = []
tw_totals = int(0)
prev_totals = int(0)
if do_xall_errors:
    zall_errors = len(create_all_errors())
if do_x4400:
    z4400 = len(create_4400_errors())
    p4400 = percentage(z4400, zall_errors)
    x4400_summary_list = ['4400', p4400, z4400, prev_x4400, int(z4400) - int(prev_x4400)]
    summary_list.append(x4400_summary_list)
    tw_totals = tw_totals + z4400
    prev_totals = prev_totals + prev_x4400
if do_x5461:
    z5461 = len(create_5461_errors())
    p5461 = percentage(z5461, zall_errors)
    x5461_summary_list = ['5461', p5461, z5461, prev_x5461, int(z5461) - int(prev_x5461)]
    summary_list.append(x5461_summary_list)
    tw_totals = tw_totals + z5461
    prev_totals = prev_totals + prev_x5461
if do_x5002_8123:
    z5002_8123 = len(create_5002_8123_errors())
    p5002_8123 = percentage(z5002_8123, zall_errors)
    x5002_8123_summary_list = ['5002 - 8123', p5002_8123, z5002_8123, prev_x5002_8123,
                               int(z5002_8123) - int(prev_x5002_8123)]
    summary_list.append(x5002_8123_summary_list)
    tw_totals = tw_totals + z5002_8123
    prev_totals = prev_totals + prev_x5002_8123
if do_x8541:
    z8541 = len(create_8541_errors())
    p8541 = percentage(z8541, zall_errors)
    x8541_summary_list = ['8541', p8541, z8541, prev_x8541, int(z8541) - int(prev_x8541)]
    summary_list.append(x8541_summary_list)
    tw_totals = tw_totals + z8541
    prev_totals = prev_totals + prev_x8541
if do_x8542:
    z8542 = len(create_8542_errors())
    p8542 = percentage(z8542, zall_errors)
    x8542_summary_list = ['8542', p8542, z8542, prev_x8542, int(z8542) - int(prev_x8542)]
    summary_list.append(x8542_summary_list)
    tw_totals = tw_totals + z8542
    prev_totals = prev_totals + prev_x8542
if do_x8560:
    z8560 = len(create_8560_errors())
    p8560 = percentage(z8560, zall_errors)
    x8560_summary_list = ['8560', p8560, z8560, prev_x8560, int(z8560) - int(prev_x8560)]
    summary_list.append(x8560_summary_list)
    tw_totals = tw_totals + z8560
    prev_totals = prev_totals + prev_x8560
if do_x8566:
    z8566 = len(create_8566_errors())
    p8566 = percentage(z8566, zall_errors)
    x8566_summary_list = ['8566', p8566, z8566, prev_x8566, int(z8566) - int(prev_x8566)]
    summary_list.append(x8566_summary_list)
    tw_totals = tw_totals + z8566
    prev_totals = prev_totals + prev_x8566
if do_x8567:
    z8567 = len(create_8567_errors())
    p8567 = percentage(z8567, zall_errors)
    x8567_summary_list = ['8567', p8567, z8567, prev_x8567, int(z8567) - int(prev_x8567)]
    summary_list.append(x8567_summary_list)
    tw_totals = tw_totals + z8567
    prev_totals = prev_totals + prev_x8567
if do_x8572:
    z8572 = len(create_8572_errors())
    p8572 = percentage(z8572, zall_errors)
    x8572_summary_list = ['8572', p8572, z8572, prev_x8572, int(z8572) - int(prev_x8572)]
    summary_list.append(x8572_summary_list)
    tw_totals = tw_totals + z8572
    prev_totals = prev_totals + prev_x8572
if do_x99001:
    z99001 = len(create_99001_errors())
    p99001 = percentage(z99001, zall_errors)
    x99001_summary_list = ['99001', p99001, z99001, prev_x99001, int(z99001) - int(prev_x99001)]
    summary_list.append(x99001_summary_list)
    tw_totals = tw_totals + z99001
    prev_totals = prev_totals + prev_x99001
if do_x17001_300404:
    z17001_300404 = len(create_17001_300404_errors())
    p17001_300404 = percentage(z17001_300404, zall_errors)
    x17001_300404_summary_list = ['17001 - 300404', p17001_300404, z17001_300404, prev_x17001_300404, int(z17001_300404) - int(prev_x17001_300404)]
    summary_list.append(x17001_300404_summary_list)
    tw_totals = tw_totals + z17001_300404
    prev_totals = prev_totals + prev_x17001_300404
if do_summary:
    create_summary()

errors_workbook.close()
print(" ")
print("---------------------------------------------------------------------------------------------------------------")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                       The Amazon Error Report Script has finished                                  #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
