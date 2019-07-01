"""
FILE: Combined_BOM.py

PURPOSE: 
When running this script, the user will be able to create one large flat BOM
for purchasing purposes.  This script will only operate on .xls files, and not
.xlsx files.  This script will automatically sift through all files in the 
current working directory, and with each file, it will iterate over all sheets. 

AUTHOR: 
Clinton G. 

TODO: 	Nothing

"""
import sys
import random
import time
import csv
import re
import os
import xlrd		
import xlwt

## DEFINE VRIABLES ##
#####################
MFGPN_col 	= 0				# Column number containing the MFGPN
QPN_col 	= 0				# Column number containing QPN
MFG_col 	= 0				# Column location for manufacturer part number
DES_col 	= 0 			# Column location for descritpion part number
QTY_col 	= 0 			# Column location for quantity field
CR1_col		= 0				# Column location for supplier name
CR1PN_col	= 0				# Column location for supplier's PN
NOTE_col 	= 0 			# Column location for "notes" field
BOM_HEADER 	= ["QPN","QTY","DES","MFG","MFGPN","CR1","CR1PN","NOTES"]
EXTS		= ('.xls')		# Support file extensions -- not currently used

MAX_HITS = 5			#This many hits will trigger us to leave the searching loop
data_start = 0			#This is the row where the data starts

search_header 	= []		# Set equal to BOM_HEADER and pop elements until we find all the colums we're looking for
header 			= []		# This array will define the column locations for the header
qpn 			= []        # Pull in all QPNs into a list. This will make them easier to work with later
asso 			= []       	# Pull in all associations into a list. This will make them easier to work with later
qty 			= []        # Pull in all QTYs into a list. This will make them easier to work with later
des 			= []		# Pull in all Descriptions into a list. This will make them easier to work with later
mfg 			= []		# Pull in all Manufactures into a list. This will make them easier to work with later
mfgpn 			= []		# Pull in all Manufacturing Part Numbers into a list. This will make them easier to work with later
cr1 			= []		# Pull in all suppler names into a list. This will make them easier to work with later
cr1pn 			= []		# Pull in all supplier pn's into a list. This will make them easier to work with later
notes 			= []		# Pull all note values into a list. This will make them easier to work with later


## FUNCTIONS ##
######################	   
def debugbreak():
	while(1):
		pass
		
def clean_value(textin):
	temptext = textin
	temptext = temptext.lstrip('text:u\'')     #Remove the initial part of the string that we don't need 'text:u'   
	temptext = temptext.replace(" ","")			#This will remove any and all white spaces
	temptext = temptext.replace("'","")			#This will remove any and all white spaces
	if(temptext.find("number:") != -1):
		temptext = temptext.replace("number:","")			#This will remove any and all white spaces
	if(temptext.find("mpty:") != -1):
		temptext = temptext.replace("mpty:","")			#This will remove any and all white spaces
	return temptext
	
def clean_des(textin):
	temptext = textin
	temptext = temptext.lstrip('text:u\'')     #Remove the initial part of the string that we don't need 'text:u'  
	temptext = temptext.replace("'","")			#This will remove any and all white spaces
	if(temptext.find("mpty:") != -1):
		temptext = temptext.replace("mpty:","")			#This will remove any and all white spaces
	return temptext
	   
#****************************************************************************** 
#******************************  ---MAIN---  **********************************
#******************************************************************************   
if __name__ == '__main__':

	path = os.getcwd()
	
	for (path, dirs, files) in os.walk(path):		# Find path/dirs/files
		path
		dirs
		files

	print "Files found in directory: " + str(len(files))

	# Iterate through all files in directory
	for i in range(len(files)-1):
		
		# Search through and open files that are appended with *.xls
		if(files[i].endswith(".xls")):
			workbook = xlrd.open_workbook(files[i])     #Open the workbook that we are going to parse though 
			worksheets = workbook.sheet_names()             #Grab the names of the worksheets -- I believe this line is critical.
			num_sheets = len(worksheets)				#This is the number of sheet
			
			print "\n\n==============================================="
			print "==============================================="
			print "File opened: " + str(files[i])
			print "The number of worksheets is: " + str(num_sheets)
			print "+++++++++++++++++++++++++++++++++++++++++++++++"
			
			# Iterate through all sheets in the opened file
			for sh in range (num_sheets):
				sheet_valid = True											# Later set to false, for example, in cases where sheet only contains change descriptions
				current_sheet = workbook.sheet_by_name(worksheets[sh])		# Grab the first worksheet
				print 'Now operating on worksheet: ', str(worksheets[sh])
				association = raw_input("Enter a unique association / high-level QPN for this worksheet (i.e. Prog Cbl): ") 
				
				num_rows = current_sheet.nrows - 1      #This is how many rows are in the worksheet
				num_cols = current_sheet.ncols - 1 		#This is how many columns are in the sheet

				for curr_row in range (num_rows + 1):			# Find the header locations
					row = current_sheet.row(curr_row)					#Grab the current row
					search_header = ["QPN","QTY","DES","MFG","MFGPN","CR1","CR1PN","NOTES"]					# Load up headers we need to search for
					print 'Search header before starting: ', search_header

					# Iterate over columns of current row
					for curr_col in range (num_cols + 1):
						temptext = unicode(row[curr_col])           # This is the fifth cell of the current row
						temptext = temptext.lstrip('text:u\'')     	# Remove the initial part of the string that we don't need 'text:u'   
						temptext = temptext.rstrip('\'')     		# Remove the initial part of the string that we don't need 'text:u'   
						temptext = temptext.replace(" ","")			# This will remove any and all white spaces
						
						if((temptext.find("QPN")!=-1) or (temptext.find("qpn")!=-1)):
							QPN_col = curr_col
							search_header.remove("QPN")
						
						elif((temptext.find("MFGPN")!=-1) or (temptext.find("MFG PN")!=-1)):	#Look for MFGPN header
							MFGPN_col = curr_col
							search_header.remove("MFGPN")
						
						elif((temptext.find("MFG")!=-1) and (temptext.find("PN") == -1)):		#Look for MFG -- make sure PN is not present
							MFG_col = curr_col
							search_header.remove("MFG")
						
						elif((temptext.find("Des")!=-1) or (temptext.find("DES")!=-1) or (temptext.find("Description")!=-1) or (temptext.find("DESCRIPTION")!=-1)):		#Look for Description
							DES_col = curr_col
							search_header.remove("DES")
						
						elif((temptext.find("Qty")!=-1) or (temptext.find("QTY")!=-1)):		#Look for Quantity field.  
							search_header.remove("QTY")
							QTY_col = curr_col
						
						elif( ((temptext.find("cr1")!=-1) or (temptext.find("CR1")!=-1)) and (len(temptext) <= 4) ):		#Look for CR1, and cannot have PN as in CR1PN
							search_header.remove("CR1")
							CR1_col = curr_col
						
						elif( ((temptext.find("cr1pn")!=-1) or (temptext.find("CR1PN")!=-1)) and (len(temptext) > 4) ):		#Look for CR1PN
							search_header.remove("CR1PN")
							CR1PN_col = curr_col
						
						elif((temptext.find("notes")!=-1) or (temptext.find("NOTES")!=-1) or (temptext.find("Notes")!=-1) or (temptext.find("Note")!=-1) or (temptext.find("note")!=-1) ):		#Look for Notes 
							search_header.remove("NOTES")
							NOTE_col = curr_col
					
					if( (len(search_header) == 0) ):		# Found all header fields
						data_start = curr_row + 1			# Plenty of confidence at this point that we've found data start
						print 'Data appears to start on row: ', data_start
						row = current_sheet.row(data_start)					#Grab the current row
						print 'Sample data in start row: ', clean_value(unicode(row[QPN_col])), ' ', clean_value(unicode(row[DES_col])), ' ', clean_value(unicode(row[MFG_col])), ' ', clean_value(unicode(row[MFGPN_col]))
						break
					
					elif( (curr_row == 10) and (len(search_header) > 0) and sh < num_sheets ):		# Some header fields are missing, so shutdown
						print "* File: ", str(files[i]), "Sheet: " + str(worksheets[sh]) + " -- did not find headers: ", search_header
						user_input = raw_input("Press any key to exit...")
						sys.exit(0)

					elif((curr_row == 10) and (len(search_header) > 0) and sh >= num_sheets ):
						sheet_valid = False
						print "* File: ", str(files[i]), "Invalid Sheet: " + str(worksheets[sh]) + " -- did not find headers: ", search_header
						break
					
				if(sheet_valid):
					print "QPN column found to be: " + str(QPN_col)		
					print "QTY column found to be: " + str(QTY_col)
					print "Description column found to be: " + str(DES_col)		
					print "MFG column found to be: " + str(MFG_col)
					print "MFGPN column found to be: " + str(MFGPN_col)
					print "CR1 column found to be: " + str(CR1_col)
					print "CR1PN column found to be: " + str(CR1PN_col)
					print "NOTES column found to be: " + str(NOTE_col)
					
					header = [0,QPN_col,DES_col,MFG_col,MFGPN_col,CR1_col,CR1PN_col,QTY_col,NOTE_col]
					header_values = ["Association","QPN","DES","MFG","MFGPN","CR1","CR1PN","QTY","NOTES"]
					
					# Now iterate through all rows of the current sheet and populate the data lists
					blank_row_count = 0		# Reset number of blank rows detected.  When three in a row are detected, break out of the loop. 
					for curr_row in range (data_start,num_rows + 1):
						row = current_sheet.row(curr_row)					#Grab the current row
						
						
						# If multiple columns are blank, break out of this loop for these are empty cells
						if( (len(clean_value(unicode(row[QPN_col]))) <= 1) and ( len(clean_des(unicode(row[DES_col]))) <= 1) and
							( len(clean_value(unicode(row[MFG_col]))) <= 1) ):
							blank_row_count += 1				# Increase value of blank row count
							print 'Blank row detected at row (', curr_row, ')'
						else:
							blank_row_count = 0					
							asso.append(association)				#For each row in the BOM, we need to append the association
							print 'Sample data current row (', curr_row,'): ', clean_value(unicode(row[QPN_col])), ' ', clean_value(unicode(row[DES_col])), ' ', clean_value(unicode(row[MFG_col])), ' ', clean_value(unicode(row[MFGPN_col]))
							
							current_value = clean_value(unicode(row[QPN_col]))
							qpn.append(current_value)			
							
							current_value = clean_des(unicode(row[DES_col]))
							des.append(current_value)
							
							current_value = clean_value(unicode(row[MFG_col]))
							mfg.append(current_value)
							
							current_value = clean_value(unicode(row[MFGPN_col]))
							mfgpn.append(current_value)
							
							current_value = clean_value(unicode(row[CR1_col]))
							cr1.append(current_value)
							
							current_value = clean_value(unicode(row[CR1PN_col]))
							cr1pn.append(current_value)
							
							current_value = clean_value(unicode(row[QTY_col]))
							qty.append(current_value)
							
							current_value = clean_des(unicode(row[NOTE_col]))
							notes.append(current_value)

						if(blank_row_count >= 3):
							break								# Too many blank rows detected, so break out of the loop.  
					
	ob = xlwt.Workbook()						# Create a document for our combined BOM
	Sheet1 = ob.add_sheet("Sheet1")				# Add a sheet to our new workbook

	print "\n+++++++++++++++++++++++++++++++++++++++++++++++"
	print "+++++++++++++++++++++++++++++++++++++++++++++++"
	print "Creating combined BOM"
	
	# Write the header values
	for i in range (len(header)):
		Sheet1.write(0,i,header_values[i])
	
	# Write rows of the combined BOM
	for i in range (len(asso)):				
		Sheet1.write(i+1,0,asso[i])  #Offset by one to account for header that's been written
		Sheet1.write(i+1,1,qpn[i])
		Sheet1.write(i+1,2,des[i])
		Sheet1.write(i+1,3,mfg[i])
		Sheet1.write(i+1,4,mfgpn[i])
		Sheet1.write(i+1,5,cr1[i])
		Sheet1.write(i+1,6,cr1pn[i])
		Sheet1.write(i+1,7,qty[i])
		Sheet1.write(i+1,8,notes[i])
		print ".",

	ob.save("CombinedBOM.xls")
	null=raw_input('Press enter to close window.')
