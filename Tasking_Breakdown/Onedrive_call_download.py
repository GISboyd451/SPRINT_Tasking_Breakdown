


import time
import os
import sys
import datetime
import pandas as pd
from shutil import copyfile
from glob import glob
#import O365
#import onedrivesdk # This package is deprecated
#import microsoftgraph # Other Microsoft API

##### NOTES #####
# onedrive_dir path will need to be updated depending on the user running the script. The user will need to sync the 'Team Tasking Breakdown' from the sharepoint page in order to process the sheets.
# Current path to (Team Tasking Breakdown) used by the sprint team: 
# BLM Sprint Team > Shared Files > Team Tasking Breakdown
#
#
##### Notes #####
#
###onedrive_dir = r'C:\\Users\\akboyd\\DOI\\BLM Sprint Team - Team Tasking Tasking Breakdown (1)' ##### Backup, do not delete. Match formatting below.
#
doi_dir = 'C:\\Users\\akboyd\\DOI' ##### Change to user's path
onedrive = 'BLM Sprint Team - Team Tasking Breakdown (1)' ##### Change to user's onedrive
#
# Current_download
current_download = r'\\blm\\dfs\\loc\\EGIS\\ProjectsNational\\NationalDataQuality\\Sprint\\analysis_tools\\Tasking_Breakdown\\team_tasking_breakdown\\Current_download'
# Backup Current
backup = r'\\blm\\dfs\\loc\\EGIS\\ProjectsNational\\NationalDataQuality\\Sprint\\analysis_tools\\Tasking_Breakdown\\team_tasking_breakdown\\Current_download\\backup'

if (sys.version_info > (3, 0)):
    # Python 3 code in this block
    a = 1
    print ('Run type:')
    print ('Type <1> (just the number) for full run.')
    print ('Type <2> (just the number) for submission check')
    run_type = int(input("Run Type?: "))
else:
    # Python 2 code in this block
    a = 2
    print ('Run type:')
    print ('Type <1> (just the number) for full run.')
    print ('Type <2> (just the number) for submission check')
    run_type = int(raw_input("Run Type?: "))
    
if run_type == 1:
    # Change directory
    try:
        os.chdir(doi_dir)
        os.chdir(onedrive)
        print ("Connected to Onedrive.")
    except:
        print ("Onedrive not found.")

    # Get list of all excel files by pattern
    drive_list = glob(os.path.join('*.xlsm'))

    # Copy files over to current download directory and to backup
    for i in drive_list:
        copyfile(i,current_download+os.sep+i)
        copyfile(i, backup+os.sep+i)

    print ('Copies and backup created.')
    print ('Ready to update Master.xlsx')
else:
    pass


# Check Submission Timestamps on each tasking breakdown
# Typical submission is within first week of new month, but lets check 'nearest submit' <= 9 days of current date.
present_date = datetime.date.today()
present_time = datetime.datetime.now()

# Define calculation to compare timeStamp column and get the latest submission from team member.
def nearest(items, pivot):
    return min(items, key=lambda x: abs(x - pivot))

# Change directory
try:
    os.chdir(current_download)
    print ("Current Downloads Folder Found.")
except:
    print ("Directory not found.")

# Get list of all excel files by pattern
file_list = glob(os.path.join('*.xlsm'))

not_ready = []

for f in file_list:
    if a == 1:
        try:
            xlsx = pd.read_excel(f, sheet_name='Archive')

            # Drop the Total and Manual Backup rows.
            xlsx = xlsx[xlsx.Task != "Total"]
            xlsx = xlsx[xlsx.Task != "Manual Backup"]

            xlsx['TimeStamp_Submission'] = xlsx['TimeStamp_Submission'].fillna(method='ffill')

            # Eliminate duplicated title rows.
            xlsx = xlsx[xlsx.Month.str.contains('Month') == False]

            # Convert TimeStamp_Submission string format  to datetime. 
            xlsx['TimeStamp_Submission'] = pd.to_datetime(xlsx['TimeStamp_Submission'])
            # Nearest Time calc.
            nearest_submit = nearest(xlsx['TimeStamp_Submission'], present_time)
            delta = abs((present_date - nearest_submit.date()).days)
            if delta <= 9:
                pass
            else:
                not_ready.append(f)
        except:
            print ('Could not connect to %s' % f)
            not_ready.append(f)
    else:
        try:
            xlsx = pd.read_excel(f, sheetname='Archive')

            # Drop the Total and Manual Backup rows.
            xlsx = xlsx[xlsx.Task != "Total"]
            xlsx = xlsx[xlsx.Task != "Manual Backup"]

            xlsx['TimeStamp_Submission'] = xlsx['TimeStamp_Submission'].fillna(method='ffill')

            # Eliminate duplicated title rows.
            xlsx = xlsx[xlsx.Month.str.contains('Month') == False]

            # Convert TimeStamp_Submission string format  to datetime. 
            xlsx['TimeStamp_Submission'] = pd.to_datetime(xlsx['TimeStamp_Submission'])
            # Nearest Time calc.
            nearest_submit = nearest(xlsx['TimeStamp_Submission'], present_time)
            delta = abs((present_date - nearest_submit.date()).days)
            if delta <= 9:
                pass
            else:
                not_ready.append(f)
        except:
            print ('Could not connect to %s' % f)
            not_ready.append(f)

# The error being here: invalid timestamp, they havn't submitted their sheet yet, excel sheet name or file was modified and could not conenct, or something happend in the submission process.
print ('Current File_list:')
print (file_list)
print (' ') # Space to print pretty
print ('Tasking Breakdowns Not Ready for Master or Require Check:')
print (not_ready)

# Wait for 10 seconds
time.sleep(10)
