# Combine BOM
This repository is home to the python script that will combine multiple BOMs into one large concise BOM, which alleviates the headache associated with purchasing materials. 

# Description 
When running this script, the user will be able to create one large flat BOM for purchasing purposes.  This script will only operate on .xls files, and not .xlsx files. This script will automatically sift through all files in the current working directory, and with each file, it will iterate over all sheets.  If the user is wanting to skip a file, he/she could simply change the extension of the file temporarily to something other than .xls.   

# Revisions
v0.3 -- Committed just for history's sake. 

v0.4 -- Committed just for history's sake.

v0.5 -- Supports reading multiple .xls files, and those files may have multiple sheets.  Association column created so assembly association can be understood when viewing the combined BOM.  File names of the BOM do not matter, if the extension is *.xls, the file will be opened and parsed. Columns that must exist in the BOM are QPN, MFG, MFGPN, Description, QTY, CR1, CR1PN, and Notes.  These columns can be blank, however, the script will not continue if they're missing.  

v0.6 -- No longer break out of the loop at the detection of the first blank row.  This allows for better organization of the BOM, such that header comments can be located on one of the BOM rows.  

V0.7 -- Improved debug messages displayed to user when script is running.  Fixed critical bug in which first row of extracted BOM data wasn't printing due to the fact that the for loop iterated over len(asso)-1, essentially.  

V1.0 -- Incoming data from cells is now parsed with unicode method in attempt to prevent crashing when unique ascii symbols are encountered.  Opposed to just closing, program will report which file/sheet is invalid so user knows where to make a correction.  The program will not close if there are remaining sheets, but the currently active sheet contains, for example, change descriptions.  

v1.1 -- No longer are internal white spaces removed from descriptions, notes, etc., but rather only those that are leading or trailing.  This prevents descriptions, notes, etc. from being run together.  
