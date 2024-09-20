# -*- coding: utf-8 -*-
"""
Created on Tue Sep 17 20:09:49 2024

This script converts your Outlook agenda to an .xlsx file, correctly formatted
so that it can be directly pasted into SAP.

In order for it to work, follow these steps:
    1.) Create categories in your agenda, each labelled in the following format:
    
        "Project Name | Project Code" (spaces included)
        
    where the Project Name is your description for the project e.g."RWMP WP4" 
    and the Project Code is the code you would enter into SAP, e.g "K.9372.23"
    2.) Create events/appointments in your agenda in line with the work you do.
    3.) Categorise each with the categories created in 1.)
    4.) Run this script!

Note: the script currently only sees your agenda for the current week we are in.
You can manually adjust the timing below.

@author: chapman & ChatGPT 
"""

import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# Initialize Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)  # 9 is the Calendar folder
appointments = calendar.Items

# Define the date range for the export (e.g., past year)
#user_input = input("Please enter something: ")

current_date = datetime.now()
start_date = current_date - timedelta(days=current_date.weekday())
end_date = start_date + timedelta(days=4)

# Restrict calendar items to the date range
items = calendar.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

# Filter by date range
restriction = "[Start] >= '{0}' AND [End] <= '{1}'".format(
    start_date.strftime("%m/%d/%Y"),
    end_date.strftime("%m/%d/%Y")
)
restricted_items = items.Restrict(restriction)

#Set up to find prject code (Outlook category must be in form Project Name | Project Code)
delimiter = "| "

# Extract appointment details
appointments = []
for item in restricted_items:
    appointment = {
        "Project Name": item.Categories.partition(delimiter)[0] if item.Categories.partition(delimiter)[2] != '' else 'None',
        "Project Code": item.Categories.partition(delimiter)[2] if item.Categories.partition(delimiter)[2] != '' else 'None',
        "Day"         : item.Start.date().strftime("%A"),
        "Duration"    : item.Duration/60,
    }
    appointments.append(appointment)
    print(appointment)

#initial custom array for SAP 
columns =  ['Project Name', 'Project Code','Col3','Col4','Col5','Col6', 'Monday', 'Col8','Col9','Tuesday','Col11','Col12','Wednesday','Col14','Col15','Thursday','Col17','Col18','Friday']
projectDict = {a['Project Name']:a['Project Code'] for a in appointments}
timeSheet = pd.DataFrame(0, columns = columns, index = projectDict)
timeSheet['Project Name'] = projectDict.keys()
timeSheet['Project Code'] = projectDict.values()

for appointment in appointments:
   # print(type(appointment['Duration']))
    timeSheet.loc[appointment['Project Name'], appointment['Day']] += appointment['Duration'] #.loc[row, column]]
timeSheet.round(0)                                                                                                        
timeSheet.loc[:, 'Col3':'Friday'] = timeSheet.loc[:, 'Col3':'Friday'].astype(str).applymap(lambda x: x.replace('.', ',')) #converting all . to , because we live in Europe
timeSheet.replace ('0', np.nan, inplace=True)    #replacing all zeroes with an empty cell

# Export to Excel
output_file = "timeSheet.xlsx"
timeSheet.to_excel(output_file, index=False)

print(f"Time sheet exported successfully to {output_file}")
print(output_file)