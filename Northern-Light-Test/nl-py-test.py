#!/usr/bin/python

import os
import datetime as datetime
import csv
import xlsxwriter

print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                        Starting the Northern Light Python Test Script                                     #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("---------------------------------------------------------------------------------------------------------------")
print("")

# Set system attributes
now = datetime.datetime.now()
year = now.strftime("%Y")
date_now = now.strftime("%Y-%m-%d")
current_date = datetime.datetime.today()
print('Current Date = ' + str(current_date))
os.sep = """/"""

# Set Directories
nl_py_test_dir = os.getcwd()
in_files_dir = str(nl_py_test_dir) + 'InFiles/'
out_files_dir = str(nl_py_test_dir) + 'OutFiles/'
in_file1 = str(in_files_dir) + './in-data-1.txt'
in_file2 = str(out_files_dir) + './in-data-2.txt'

in_file1_dict = {}
in_file2_dict = {}


# def get_file_data:
#     for I files in nl_py_test_dir



print(" ")
print("---------------------------------------------------------------------------------------------------------------")
print("###############################################################################################################")
print("######                                                                                                    #####")
print("######                       The Northern Light Python Test Script has finished                                  #####")
print("######                                                                                                    #####")
print("###############################################################################################################")
print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")


