#!/usr/bin/python

import csv
import datetime as datetime
import http.client
import json
import os
import shutil
import xml.dom.minidom
import ibm_db
import pandas as pd
from mws import Products

print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                        Starting the Amazon Error Report Script                                     #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("---------------------------------------------------------------------------------------------------------------")
print("")
# where org_id = s-ID, salsify:id=ID, api_token = YOUR-AUTH-TOKEN
# where org_id = s-ID, salsify:id=ID, api_token = YOUR-AUTH-TOKEN


now = datetime.datetime.now()
year = now.strftime("%Y")
date_now = now.strftime("%Y-%m-%d")
current_date = datetime.datetime.today()
print('Current Date = ' + str(current_date))
os.sep = """/"""
fix_script_dir = 'dir-for-dir'
work_files_dir = fix_script_dir + 'WorkFiles/'
archive_dir = str(work_files_dir) + 'Archive/'
in_files_dir = str(work_files_dir) + 'InFiles/'
in_process_dir = str(work_files_dir) + 'InProcessFiles/'
sl_data_dir = str(work_files_dir) + 'SLDataFiles/'
json_directory = str(work_files_dir) + 'JSON/'
xerrors_json_file = json_directory + 'xerrors_json.json'
xml_in_directory = str(work_files_dir) + 'XMLInFiles/'
xml_Process_directory = str(work_files_dir) + 'XMLProcessFiles/'
conn = ibm_db.connect('DATABASE=database-name;' 'HOSTNAME=place-database-ip-here;' 'PORT=50000;' 'PROTOCOL=TCPIP;' 'UID=db2inst1;' 'PWD=password-here;', '', '')

xerrors_dict = {}
xerrors_temp_dict = {}
xerrors_sku_list = []
fields_list = []
xerrors_list = []
column_count = int(0)

original_heading_list = ['ORIGINAL-RECORD-NUMBER', 'SKU', 'ERROR-CODE', 'ERROR-TYPE', 'ERROR-MESSAGE']
xerrors_df = pd.DataFrame(columns=['ORIGINAL-RECORD-NUMBER', 'SKU', 'ERROR-CODE', 'ERROR-TYPE', 'ERROR-MESSAGE', 'PRODUCTIDTYPE', 'STANDARDPRODUCTID'])
print('')
print(' -- Creating the initial Dictionary')
print(' -- Populating the dictionary with the proper fields')
print(' -- Creating the list of SKUs from the error report files.')


def getbranderrors():
	global xerrors_sku_list
	global xerrors_df
	global xerrors_temp_dict
	for process_files in os.listdir(in_process_dir):
		open_process_file = open(in_process_dir + process_files, mode='r', encoding='ISO-8859-1')
		xerrors_reader = csv.reader(open_process_file, delimiter='\t')
		for xerrors_row in xerrors_reader:
			if len(xerrors_row) >= 4:
				if '5461' in xerrors_row[2] or \
						'Amazon catalog: brand (Merchant:' in xerrors_row[4] or \
						'that are conflicting: brand (Merchant:' in xerrors_row[4] or \
						'Brands should be registered through Brand Registry' in xerrors_row[4]:
					if xerrors_row[1] not in xerrors_sku_list:
						xerrors_sku_list.append(xerrors_row[1])
					xerrors_key = xerrors_row[1]
					xerrors_temp_dict[xerrors_key] = xerrors_row[0:]
		brand_errors_num = int(1)
		print('-- Populating the dictionary with the proper fields')
		for i in xerrors_temp_dict.values():
			xerrors_row = dict(zip(original_heading_list, i))
			xerrors_row['STANDARDPRODUCTID'] = ''
			xerrors_row['PRODUCTIDTYPE'] = ''
			# xerrors_row['AMAZONBRAND'] = ''
			xerrors_dict[brand_errors_num] = {brand_errors_num: xerrors_row}
			xerrors_df.append(xerrors_row, ignore_index=True)
			brand_errors_num = brand_errors_num + 1
		open_process_file.close()
	return xerrors_df, xerrors_dict, xerrors_sku_list


getbranderrors()
print(xerrors_df)
print(xerrors_dict)
results_list = []
partnumbers_list = []
results_dict = {}
errors_num = int(1)
print('Getting the PRODUCTIDTYPE and STANDARDPRODUCTID from the stage database')
for skus in xerrors_sku_list:
	if conn:
		product_sql = """SELECT TRIM(PARTNUMBER) AS SKU, TRIM(STANDARDPRODUCTID) AS STANDARDPRODUCTID, TRIM(PRODUCTIDTYPE) AS PRODUCTIDTYPE FROM XAMAZONFEED
						WHERE partnumber IN ('%s');""" % str(skus)
		product_stmt = ibm_db.exec_immediate(conn, product_sql)
		product_all_info = ibm_db.fetch_assoc(product_stmt)
		while product_all_info:
			results_dict[str(skus)] = product_all_info
			product_all_info = ibm_db.fetch_assoc(product_stmt)
		errors_num = errors_num + 1

# # The Amazon API section. This should get the data from Amazon
# products_api = Products(
# access_key=os.environ["MWS_ACCESS_KEY"],
# secret_key=os.environ["MWS_SECRET_KEY"],
# account_id=os.environ["MWS_ACCOUNT_ID"],
# auth_token=os.environ["MWS_AUTH_TOKEN"],
# )
# my_market = Marketplaces.US


print('Processing the xml files in the XMLInFiles directory')
for all_in_xml_files in os.listdir(xml_in_directory):
	openfile = open(str(xml_in_directory + all_in_xml_files), 'r').read()
	openfile = openfile.replace('<ns2:', '<')
	openfile = openfile.replace('</ns2:', '</')
	openfile = open(str(xml_Process_directory + all_in_xml_files), 'w', encoding='utf-8').write(str(openfile))

to_salsify_dict = {}
xerrors_xml_dict = {}
xerrors_xml_num = int(1)
print('Parsing the XML to get the values we need to send to Salsify.')
for all_xml_files in os.listdir(xml_Process_directory):
	xerrors_xml_dom = xml.dom.minidom.parse(xml_Process_directory + all_xml_files)
	xerrors_Id = False
	xerrors_IdType = False
	xerrors_brand = False
	for product in xerrors_xml_dom.getElementsByTagName('GetMatchingProductForIdResult'):
		while not xerrors_Id and not xerrors_IdType and not xerrors_brand:
			xerrors_main_tag = xerrors_xml_dom.getElementsByTagName('GetMatchingProductForIdResult')
			print('Line 179 = ' + str(xerrors_main_tag))
			xerrors_Id = product.getAttribute('Id')
			print('Line 182 = ' + str(xerrors_Id))
			xerrors_IdType = product.getAttribute('IdType')
			print('Line 184 = ' + str(xerrors_IdType))
			xerrors_brand = xerrors_xml_dom.getElementsByTagName('Brand')
			xerrors_brand = xerrors_brand[0].firstChild.nodeValue
			print('Line 186 = ' + str(xerrors_brand))
			xerrors_PartNumber = xerrors_xml_dom.getElementsByTagName('PartNumber')
			xerrors_PartNumber = xerrors_PartNumber[0].firstChild.nodeValue
			print('Line 189 = ' + str(xerrors_PartNumber))
			print('')
		to_salsify_dict[str(xerrors_PartNumber)] = {'SKU': str(xerrors_PartNumber),
													'STANDARDPRODUCTID': str(xerrors_IdType),
													'PRODUCTIDTYPE': str(xerrors_IdType),
													'Company Tools Part Number': str(xerrors_PartNumber),
													'Brand Name (Amazon)': str(xerrors_brand),
													'Import to WCE?': 'Yes',
													'Import to SFCC?': 'Yes'}
for all_stuff in to_salsify_dict.values():
	del all_stuff['SKU']
	del all_stuff['STANDARDPRODUCTID']
	del all_stuff['PRODUCTIDTYPE']
to_salsify_json = json.dumps(to_salsify_dict, indent=4)
print(to_salsify_json)
print('')

for xerrors_skus in xerrors_sku_list:
	conn = http.client.HTTPSConnection("app	.salsify.com")
	headers = {'authorization': "Bearer place-ID-HERE",}
	conn.request("GET", "/api/v1/orgs/s-place-salsify-account-id here/products/" + str(xerrors_skus), headers=headers)
	res = conn.getresponse()
	data = res.read()
	xerrors_json_file = str(json_directory) + str(xerrors_skus) + '-json.json'

	xerrors_json = data.decode("ANSI")
	write_json = open(xerrors_json_file, 'w')
	write_json.write(xerrors_json)
	print(str(xerrors_json_file) + ' was created.')
	write_json.close()



print(" ")
print("---------------------------------------------------------------------------------------------------------------")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                       The Amazon Error Report Script has finished                                  #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")


