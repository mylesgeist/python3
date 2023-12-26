#!/usr/bin/env python
# Written in Python 2

import sys
import time
import xml.dom.minidom

today = time.strftime('%Y%m%d')
column_list = "Name Part Number, Approved?"
dom = xml.dom.minidom.parse(open('Filename'+today+'.xml'))
orig_stdout = sys.stdout
CSVFile = open('C:\\TEMP\\Salsify_Activation.csv', 'w')
entries = []
sys.stdout = CSVFile
print column_list
sys.stdout = orig_stdout
items = 0
passed_items = 0
failed_items = 0
not_printed = True

print "------------------------"
print ""
print " List of Failed Items:"
print ""

for product in dom.getElementsByTagName('Product'):
	for identifier in product.getElementsByTagName('Identifier'):
		not_printed = True
		if identifier.getAttribute('IdentifierName') == 'Part Number':
			try:
				sys.stdout = CSVFile
				entry = identifier.getAttribute('Value')
				print str(entry) + ',' + 'Yes'
				items = items + 1
				passed_items = passed_items + 1
			except:
				sys.stdout = orig_stdout
				for identifier in product.getElementsByTagName('Identifier'):
					if identifier.getAttribute('IdentifierName') == "UPC":
						if not_printed is True:
							print "UPC = " + str(identifier.getAttribute('Value'))
							items = items + 1
							failed_items = failed_items + 1
							not_printed = False
					elif identifier.getAttribute('IdentifierName') == "Internet Product Code (IPC)":
						if not_printed is True:
							print "IPC = " + str(identifier.getAttribute('Value'))
							items = items + 1
							failed_items = failed_items + 1
							not_printed = False

sys.stdout = orig_stdout
if failed_items == 0:
	print "There were no failed items."

print ""
print "------------------------"

print "Number of Items Passed = " + str(passed_items)
print 'Number of Errors       = ' + str(failed_items)
print "Number of Total Items  = " + str(items)

CSVFile.close()

if failed_items >= 1:
	print ""
	print "Program Exiting With Some Items With An Error".title()
	sys.exit(1)
else:
	print ""
	print "Program Completed Without Any Errors".title()
	sys.exit(1)
