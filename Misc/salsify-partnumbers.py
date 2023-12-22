#!/usr/bin/python
# coding: utf-8
#print "|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
#print "###############################################################################################################"
#print "######                                                                                                   ######"
#print "######                       Begining Empty text box                                                     ######"
#print "######                                                                                                   ######"
#print "###############################################################################################################"
import os
import fnmatch
import xml.dom.minidom
import datetime
import sys
import pycurl
import json
import certifi
from StringIO import StringIO

script_start_time = datetime.datetime.utcnow().isoformat('T')
print "---------------------------------------------------------------------------------------------------------------"
print ""

##################
# Global Variables
##################
# variables for working locally.
rootDir = 'C://Users//admin//TEMP//salsify-products'
destDir = 'C://Users//admin//TEMP'
file1 = open(str(destDir) + '//datafile.csv', 'w')
partNumberDict = {}
partNumberList = []
notInSalsifyList = []

# Variables for working with the Salsify API.
import_id = sys.argv[1:]
import_id = str(import_id)
status = 'queued'
progress = '0'
start_time = str(datetime.datetime.utcnow().isoformat('T'))
end_time = str(datetime.datetime.utcnow().isoformat('T'))
print "Checking to see which items in the daily XML files are not in Salsify already.".title()
print "".title()


# open each XML file in a directory and get part number from it and put it in a list
def getpartnumberlist():
	global partNumberList
	print "Going Through The XML Files and Picking Out The Part Numbers And Putting them in a list.".title()
	print "The time is now ".title() + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time'
	for subdir, dir, files in os.walk(rootDir):

		# We go through each file in the directory here.
		for allFiles in files:

			# Check and make sure it's a PRODUCT XML file
			if fnmatch.fnmatch(allFiles, 'Company_ContentExport_Delta_PRODUCT_*.xml'):

				# We open the XML file here and parse the XML using minidom.
				with open(str(rootDir) + '//' + str(allFiles), 'r') as currentFile:
					dom = xml.dom.minidom.parse(currentFile)

					# This loops through the file and gets the part number from the tag we get the info from
					for product in dom.getElementsByTagName('Product'):
						for identifier in product.getElementsByTagName('Identifier'):
							if identifier.getAttribute('IdentifierName') in ('Company Tools Part Number'):

								# Evaluate each part number to see if it's already in the list
								# If it's not already in the list then we add it.
								# This prevents duplicates.
								partnumber = identifier.getAttribute('Value').encode('ascii', errors='ignore').decode()
								partnumber = partnumber.strip()
								partnumber = partnumber.upper()
								if partnumber not in partNumberList:
									partNumberList.append(partnumber)
				# close the file so it doesn't stay open.
				currentFile.close()

	partNumberList = partNumberList.sort()
	return partNumberList
	print "We now have the List that we need starting at ".title() + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time'.title()
	print "".title()

#print partNumberList

#testparts = ['48-32-4484','DWA4769','17932','DWST24190K','100-LG','DWST25294K','2718-22HD','XT336T','18824','DWST25292K','SS560VSC-31','2101R','T50079','18816W','GHO12V-08N','DWST24191K','RT-BT','DWST22760K','252292','460060','40126N','783642','18820W','2718-21HD','SX-105 XC','18728','RT-FS','605DAT','18724','US40-01','XT337T','T10R-01-R8','S-65 XC','50UTZ2.75S','429392','253692','CL150','PB-2620','18720W1','RT-PHFS','S-60 XC','HSE2.4S','282510','HSD2.55S','2718-20','PE-2620S','427632','18828W','18716W','17232W',]

partNumbersProcessed = 0
numberNotInSalsify = 0

def processpartnumberlist():
	global partNumbersProcessed
	global numberNotInSalsify
	for part in partNumberList:
		try:
			print("Checking Part Number: %s" % part)
			partNumbersProcessed = partNumbersProcessed + 1
			retries = 0
			json_buffer = StringIO()
			c = pycurl.Curl()
			c.setopt(c.URL, 'https://app.salsify.com/api/orgs/place-account-id-here/products/' + str(part) + '?auth_token=place-auth-token-here')
			c.setopt(pycurl.HTTPGET, 1)
			c.setopt(c.WRITEFUNCTION, json_buffer.write)
			c.setopt(c.CAINFO, certifi.where())
			c.setopt(pycurl.HTTPHEADER, ['Accept:application/json'])
			#c.setopt(pycurl.POSTFIELDS, 'auth_token=place-auth-token-here')
			c.setopt(c.VERBOSE, False)
			c.perform()
			c.close()

			jsonobject = json_buffer.getvalue()
			#print jsonobject

			#print json.dumps(jsonobject, indent=4, sort_keys=True)
			partresults = json.loads(jsonobject)
			jsonpartnum = partresults["data"]['external_id']

			if jsonpartnum.startswith("[") and jsonpartnum.endswith("]"):
				jsonpartnum = jsonpartnum[1:-1]
			if jsonpartnum.startswith("'") and jsonpartnum.endswith("'"):
				jsonpartnum = jsonpartnum[1:-1]
			if jsonpartnum.startswith('"') and jsonpartnum.endswith('"'):
				jsonpartnum = jsonpartnum[1:-1]

			if part.startswith("[") and part.endswith("]"):
				part = part[1:-1]f
			if part.startswith("'") and part.endswith("'"):
				part = part[1:-1]
			if part.startswith('"') and part.endswith('"'):
				part = import_id[1:-1]

			if jsonpartnum != part:
				if part not in notInSalsifyList:
					notInSalsifyList.append(part)
					file1.write(str(part) + "\n")
					numberNotInSalsify = numberNotInSalsify + 1
					print "        Part Number %s: Failed " % part
			else:
				print "        Part Number %s: Passed " % part

		except:
			if part not in notInSalsifyList:
				notInSalsifyList.append(part)
				file1.write(str(part) + "\n")
				numberNotInSalsify = numberNotInSalsify + 1
				print "        Part Number %s: Failed " % part


getpartnumberlist()

numOfParts = 0
for i in partNumberList:
	numOfParts = numOfParts + 1
print numOfParts

processpartnumberlist()

file1.close()

print "--------------------------------"
print "Part Numbers Processed: %s" % partNumbersProcessed
print "Part Numbers not in Salsify: %s" % numberNotInSalsify


print ""
print "----------------------------------------------------------------------------------------------------------------"
#print "###############################################################################################################"
#print "######                                                                                                   ######"
#print "######                       Ending empty text box                                                       ######"
#print "######                                                                                                   ######"
#print "###############################################################################################################"
#print "|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"


