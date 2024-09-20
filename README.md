This script converts your Outlook agenda to an .xlsx file, correctly formatted
so that it can be directly pasted into SAP.

In order for it to work, follow these steps:
    1.) Create categories in your agenda, each labelled in the following format:
    
        "Project Name | Project Code" (spaces included)
        
    where the Project Name is your description for the project e.g."RWMP WP4" 
    and the Project Code is the code you would enter into SAP, e.g "K.9372.23"
    2.) Create events/appointments in your agenda in line with the work you do.
    3.) Categorise each with the categories created in 1.) (see Note 1 below)
    4.) Run this script!

Note 1: You can leave appointments uncategorised. These will be counted as "None"
in the spreadsheet.
Note 2: the script currently only sees your agenda for the current week we are in.
You can manually adjust the timing below.
