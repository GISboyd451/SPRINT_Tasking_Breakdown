
#############################################################################################
######## This area is a development area for future connection sharepoint if needed #########
# Different authentication scheme? 
# You can check this by inspecting the network traffic in Firebug or the Chrome Developer Tools.
# Luckily, the requests library supports many authentication options: http://docs.python-requests.org/en/latest/user/authentication/
# Possible sharepoint integration code?:
'''
import requests
from requests_ntlm import HttpNtlmAuth

requests.get("http://sharepoint-site.com", auth=HttpNtlmAuth('DOMAIN\\USERNAME','PASSWORD'))
'''
# https://pypi.org/project/Office365-REST-Python-Client/

########## End Development area ######################################################################
######################################################################################################
######################################################################################################
import time
import os
import sys
import datetime
import pandas as pd
import xlsxwriter
from glob import glob
#import onedrivesdk ### Not needed as of 02/03/2020

# Global Variables
root  = r'\\blm\\dfs\\loc\\EGIS\\ProjectsNational\\NationalDataQuality\\Sprint\\analysis_tools\\Tasking_Breakdown\\'
# not_ready variable is taken from the Onedrive_call_download.py script ; comment out if not needed
# Do we want to subset input question:

if (sys.version_info > (3, 0)):
    # Python 3 code in this block
    a = 1
    q = input("Do you want to subset file_list?(y/n): ")
    if q == 'y':
        print ('Copy the list of elements from the Onedrive_call_download script into input --------> Tasking Breakdowns Not Ready for Master:')
        print ('The names pasted in input will not be included in the test_master.xlsx output file. Only tasking breakdown files that are ready will be appended.')
        print ("""Example: ['Dawn_tasking_sheet.xlsm', 'Jamie_tasking_sheet.xlsm', 'John_tasking_sheet.xlsm']""")
        not_ready = eval(input("input: ")) # Convert input to list
    else:
        print ('Performing full run, no subset of file_list.')
        pass
else:
    # Python 2 code in this block
    a = 2
    q = raw_input("Do you want to subset file_list?(y/n): ")
    if q == 'y':
        print ('Copy the list of elements from the Onedrive_call_download script into input --------> Tasking Breakdowns Not Ready for Master:')
        print ('The names pasted in input will not be included in the test_master.xlsx output file. Only tasking breakdown files that are ready will be appended.')
        print ("""Example: ['Dawn_tasking_sheet.xlsm', 'Jamie_tasking_sheet.xlsm', 'John_tasking_sheet.xlsm']""")
        not_ready = input("input: ") # python 2, eval is built in
    else:
        print ('Performing full run, no subset of file_list.')
        pass
    
#####not_ready = [] # Can hardcode if needed
# gui_fram output path
gui_output = r'\\blm\\dfs\\loc\\EGIS\\ProjectsNational\\NationalDataQuality\\Sprint\\analysis_tools\\Sprint_gui\\outputs'

# Current_download from sharepoint
current_download = r'\\blm\\dfs\\loc\\EGIS\\ProjectsNational\\NationalDataQuality\\Sprint\\analysis_tools\\Tasking_Breakdown\\team_tasking_breakdown\\Current_download'

# Master file where we will append team submissions to.
try:
    master_file = pd.ExcelFile(root + os.sep +'test_master.xlsx')
    master_file = pd.read_excel(master_file, sheet_name = "Master")
    master_file = pd.DataFrame(master_file)
    print ('Connected to Master File/Sheet.')
except:
    print ("Master File/Sheet not found.")

# Directory where all team tasking excel files are downloaded to from sharepoint.
# Change directory
try:
    os.chdir(current_download)
    print ("Directory found.")
except:
    print ("Directory not found.")

# Get list of all excel files by pattern
file_list_1 = glob(os.path.join('*.xlsm'))

# Remove the "Not ready" Tasking breakdown from file lists
if q == 'y':
    file_list = [x for x in file_list_1 if x not in not_ready]
    print ('Created subset.')
else:
    file_list = file_list_1
    pass
# Remove Master Excel from file_list
#while 'test_master.xlsx' in file_list: file_list.remove('test_master.xlsx')

# Computer system time.
present_time = datetime.datetime.now()
print (present_time)

# Define calculation to compare timeStamp column and get the latest submission from team member.
def nearest(items, pivot):
    return min(items, key=lambda x: abs(x - pivot))

# Temp object to append xlsx to.
tmp = []

for f in file_list:
    if a == 1:
        try:
            xlsx = pd.read_excel(f, sheet_name='Archive')
            print ('Connected to %s' % f)

            # Drop the Total and Manual Backup rows.
            xlsx = xlsx[xlsx.Task != "Total"]
            xlsx = xlsx[xlsx.Task != "Manual Backup"]

            # Fill in rows to complete time sheet format.
            xlsx['Month#'] = xlsx['Month#'].fillna(method='ffill')
            xlsx['Year#'] = xlsx['Year#'].fillna(method='ffill')
            xlsx['Month'] = xlsx['Month'].fillna(method='ffill')
            xlsx['Month_Year#'] = xlsx['Month_Year#'].fillna(method='ffill')
            xlsx['TimeStamp_Submission'] = xlsx['TimeStamp_Submission'].fillna(method='ffill')
            
            # Eliminate duplicated title rows.
            xlsx = xlsx[xlsx.Month.str.contains('Month') == False]

            # Convert TimeStamp_Submission string format  to datetime. 
            xlsx['TimeStamp_Submission'] = pd.to_datetime(xlsx['TimeStamp_Submission'])
            # Nearest Time calc.
            nearest_submit = nearest(xlsx['TimeStamp_Submission'], present_time)

            # Select rows based off of the variable nearest_submit.
            xlsx = xlsx.loc[xlsx['TimeStamp_Submission'] == nearest_submit]

            # Append xlsx data to temp object.
            tmp.append(xlsx)
        except:
            print ('Could not connect to %s' % f)
    else:
        try:
            # Python 2 format (sheetname)
            xlsx = pd.read_excel(f, sheetname='Archive')
            print ('Connected to %s' % f)

            # Drop the Total and Manual Backup rows.
            xlsx = xlsx[xlsx.Task != "Total"]
            xlsx = xlsx[xlsx.Task != "Manual Backup"]

            # Fill in rows to complete time sheet format.
            xlsx['Month#'] = xlsx['Month#'].fillna(method='ffill')
            xlsx['Year#'] = xlsx['Year#'].fillna(method='ffill')
            xlsx['Month'] = xlsx['Month'].fillna(method='ffill')
            xlsx['Month_Year#'] = xlsx['Month_Year#'].fillna(method='ffill')
            xlsx['TimeStamp_Submission'] = xlsx['TimeStamp_Submission'].fillna(method='ffill')
            
            # Eliminate duplicated title rows.
            xlsx = xlsx[xlsx.Month.str.contains('Month') == False]

            # Convert TimeStamp_Submission string format  to datetime. 
            xlsx['TimeStamp_Submission'] = pd.to_datetime(xlsx['TimeStamp_Submission'])
            # Nearest Time calc.
            nearest_submit = nearest(xlsx['TimeStamp_Submission'], present_time)

            # Select rows based off of the variable nearest_submit.
            xlsx = xlsx.loc[xlsx['TimeStamp_Submission'] == nearest_submit]

            # Append xlsx data to temp object.
            tmp.append(xlsx)
        except:
            print ('Could not connect to %s' % f)
# 
# Use pd.concat to merge a list of DataFrame into a single big DataFrame.
master_file = pd.concat(tmp)

# Remove NA, NAN, from dataframe.
master_file = master_file.fillna("")

# Write to Master File -----> Append to 'Master' sheet.
out_path = root + os.sep +'test_master.xlsx'
writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
master_file.to_excel(writer, sheet_name = "Master", index = False)

gui_path = gui_output + os.sep + 'Task_master_sheet_update.xlsx'
www = pd.ExcelWriter(gui_path, engine='xlsxwriter')
master_file.to_excel(www, sheet_name = "Master", index = False)

try:
    writer.save()
    www.save()
    print ('Master File Updated.')
except:
    print ('Failed to save Master File.')
