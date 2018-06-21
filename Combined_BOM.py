"""
FILE: Combined_BOM.py

PURPOSE: 
When running this script, the user will be able to create one large flat BOM
for purchasing purposes.  This script will only operate on .xls files, and not
.xlsx files.  This script will automatically sift through all files in the 
current working directory, and with each file, it will iterate over all sheets. 

AUTHOR: 
Clinton G. 

TODO: 	For future work, it would be nice to support .xlsx files.
		As it currently sites, boiling down a large complex BOM means
		there could be multiple lines of the same items -- for example, 
		a #6-32 screw is called out on multiple assemblies.  Future work
		should incorporate a feature added early on in which duplicates are 
		removed and the quantity adjusted to account for this. 

"""
import sys
import random
import time
import csv
import re
import os
import xlrd		#TODO should be able to remove these lines
import xlwt

## DEFINE VRIABLES ##
#####################
MFGPN_col 	= 0				#Column number containing the MFGPN
QPN_col 	= 0				#Column number containing QPN
MFG_col 	= 0				#Column location for manufacturer part number
DES_col 	= 0 			#Column location for descritpion part number
QTY_col 	= 0 			#Column location for quantity field
CR1_col		= 0				# Column location for supplier name
CR1PN_col	= 0				# Column location for supplier's PN
NOTE_col 	= 0 			#Column location for "notes" field
BOM_HEADER 	= ["QPN","QTY","DES","MFG","MFGPN","CR1","CR1PN","NOTES"]

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
	# Find path/dirs/files
	for (path, dirs, files) in os.walk(path):
		path
		dirs
		files

	print "Files found in directory: " + str(len(files))

	# Iterate through all files in directory
	for i in range(len(files)-1):
		
		# Search through and open files that are appended with *.xlsx
		if((files[i].find("xls") != -1) and not (files[i].find("xlsx") != -1)):
			
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
				current_sheet = workbook.sheet_by_name(worksheets[sh])		#Grab the first worksheet
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
						temptext = str(row[curr_col])                  #This is the fifth cell of the current row
						temptext = temptext.lstrip('text:u\'')     #Remove the initial part of the string that we don't need 'text:u'   
						temptext = temptext.rstrip('\'')     		#Remove the initial part of the string that we don't need 'text:u'   
						temptext = temptext.replace(" ","")			#This will remove any and all white spaces
						
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
						
						elif((temptext.find("notes")!=-1) or (temptext.find("NOTES")!=-1) or (temptext.find("Notes")!=-1) or (temptext.find("Note")!=-1) or (temptext.find("note")!=-1) ):		#Look for Notes TODO make this work
							search_header.remove("NOTES")
							NOTE_col = curr_col
					
					if( (len(search_header) == 0) ):		# Found all header fields
						data_start = curr_row + 1			# Plenty of confidence at this point that we've found data start
						break
					
					elif( (curr_row == 10) and (len(search_header) > 0) ):		# Some header fields are missing, so shutdown
						print "On sheet: " + str(worksheets[sh]) + " -- did not find headers: ", search_header
						sys.exit(0)
				
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
				for curr_row in range (data_start,num_rows + 1):
					row = current_sheet.row(curr_row)					#Grab the current row

					# If multiple columns are blank, break out of this loop for these are empty cells
					if( (len(clean_value(str(row[QPN_col]))) <= 1) and ( len(clean_des(str(row[DES_col]))) <= 1) and
						( len(clean_value(str(row[MFG_col]))) <= 1) ):
						break
					
					asso.append(association)				#For each row in the BOM, we need to append the association
					print ".",
					
					current_value = clean_value(str(row[QPN_col]))
					qpn.append(current_value)			
					
					current_value = clean_des(str(row[DES_col]))
					des.append(current_value)
					
					current_value = clean_value(str(row[MFG_col]))
					mfg.append(current_value)
					
					current_value = clean_value(str(row[MFGPN_col]))
					mfgpn.append(current_value)
					
					current_value = clean_value(str(row[CR1_col]))
					cr1.append(current_value)
					
					current_value = clean_value(str(row[CR1PN_col]))
					cr1pn.append(current_value)
					
					current_value = clean_value(str(row[QTY_col]))
					qty.append(current_value)
					
					current_value = clean_des(str(row[NOTE_col]))
					notes.append(current_value)
				print ''   # Space after we print periods
	
	ob = xlwt.Workbook()						# Create a document for our combined BOM
	Sheet1 = ob.add_sheet("Sheet1")				# Add a sheet to our new workbook

	print "\n+++++++++++++++++++++++++++++++++++++++++++++++"
	print "+++++++++++++++++++++++++++++++++++++++++++++++"
	print "Creating combined BOM"
	
	# Write the header values
	for i in range (len(header)):
		Sheet1.write(0,i,header_values[i])
	
	# Write rows of the combined BOM
	for i in range (1,len(asso)):				
		Sheet1.write(i,0,asso[i])
		Sheet1.write(i,1,qpn[i])
		Sheet1.write(i,2,des[i])
		Sheet1.write(i,3,mfg[i])
		Sheet1.write(i,4,mfgpn[i])
		Sheet1.write(i,5,cr1[i])
		Sheet1.write(i,6,cr1pn[i])
		Sheet1.write(i,7,qty[i])
		Sheet1.write(i,8,notes[i])
		print ".",

	ob.save("CombinedBOM.xls")
	time.sleep(3)

		
	# TODO Future work might include eliminating duplicate items on the flat BOM.  
	# print '------------------------'
	# print 'Did you read the Readme.txt file first???'


	# qpn = []        #Suck all QPNs into a list. This will make them easier to work with later
	# qty = []        #Suck all QTYs into a list. This will make them easier to work with later
	# des = []		#Suck all Descriptions into a list. This will make them easier to work with later
	# mfg = []		#Suck all Manufactures into a list. This will make them easier to work with later
	# mfgpn = []		#Suck all Manufacturing Part Numbers into a list. This will make them easier to work with later
	# cr1 = []		#Suck all Distributors into a list. This will make them easier to work with later
	# cr1pn = []		#Suck all Distributor Part Numbers into a list. This will make them easier to work with later

	# memqpn = []   	#The user will only get a particular QPN and MFGPN copied to 'memory' once. Only QTY is then updated
	# memqty = []     #QTY is updated when a part getting sucked into memory is already present
	# memdes = []		#'Memory' field that keeps track of descriptions
	# memmfg = []		#'Memory' field that keeps track of Manufactures
	# memmfgpn = []	#The user will only get a particular QPN and MFGPN copied to 'memory' once. Only QTY is then updated
	# memcr1 = []		#'Memory' field that keeps track of descriptions
	# memcr1pn = []	#'Memory' field that keeps track of distributor Part Numbers

	# # Get all parameter stored in an array
	# for row in bomreader:
		
	# 	if row[0] != "QPN" and row[1] != "QTY":                  #Do not copy the header to memory!        
	# 		qpn.append(row[0])              #QPN parameter
	# 		qty.append(float(row[1]))       #Quantity parameter
	# 		des.append(row[2])              #Description parameter
	# 		mfg.append(row[3])              #MFG parameter
	# 		mfgpn.append(row[4])            #MFGPN parameter
	# 		cr1.append(row[5])              #Supplier parameter
	# 		cr1pn.append(row[6])            #Supplier parameter part number



	# qpn_re = re.compile('[0-9xX]{3}-[0-9xX]{6}-[0-9xX]{4}') #we need to make sure we have valid QPNs in the BOM
	# for myqpn in qpn:
	# 	if len(qpn_re.findall(myqpn)) == 0:
	# 		BadFormat = True
	# 		print 'Found QPN that is of incorrect format!!!'
			
	# for i in range(0,len(qpn)):     #This is the line we are currently looking at NOT IN 'Memory'
	# 	if len(memqpn) == 0:       	#this should only run once when there are no entries in the memory qpn
	# 		memqpn.append(qpn[i])   
	# 		memqty.append(qty[i])
	# 		memdes.append(des[i])
	# 		memmfg.append(mfg[i])
	# 		memmfgpn.append(mfgpn[i])
	# 		memcr1.append(cr1[i])
	# 		memcr1pn.append(cr1pn[i])
					
	# 	else:
	# 		for k in range (0,len(memqpn)):		#Go through all the entries in 'memory' and verify that the MFGPN and QPN does not already exist. 
	# 			if qpn[i] == memqpn[k]:     # Check against all entries in memory
	# 				if mfgpn[i] == memmfgpn[k]: #Double check to make sure the MFGPN matches
	# 					memqty[k] = memqty[k] + qty[i]      #Increase the QTY that is listed in mem
	# 					inmem = True;
	# 		if inmem == False:							#If the particular line item does not exist then put item in 'memory'
	# 			memqpn.append(qpn[i])
	# 			memqty.append(qty[i])
	# 			memdes.append(des[i])
	# 			memmfg.append(mfg[i])
	# 			memmfgpn.append(mfgpn[i])
	# 			memcr1.append(cr1[i])
	# 			memcr1pn.append(cr1pn[i])
	# 		inmem = False                   #Reset this after being set true from above.

	##print memqpn
	##print '------------------------'
	##print memqty
	##print '------------------------'
	##print memdes
	##print '------------------------'
	##print memmfg
	##print '------------------------'
	##print memmfgpn
	##print '------------------------'
	##print memcr1
	##print '------------------------'
	##print memcr1pn

	# print '------------------------'
	# tmpline1 = "The old BOM was " + str(len(qpn)) + " line items long."
	# tmpline2 = "The new BOM is only " + str(len(memqpn)) + " line items long."
	# print tmpline1
	# print tmpline2
	# print '------------------------'

	# file_out = 'combined_BOM.csv'
	# fid_out = open(file_out,'w') # 'w' this creates a new file if it does not already exist

	# for i in range (0,len(memqpn)):
	# 	tmpline = ('\"' + memqpn[i] + '\",\"' + str(memqty[i]) + '\","' + memdes[i] + '\",\"' + memmfg[i] + '\",\"' + memmfgpn[i] + '\",\"'
	# 	+ memcr1[i] + '\",\"' + memcr1pn[i] + '\"' + '\n')
	# 	fid_out.write(tmpline)

	# fid_out.close()

	# print 'FINISHED'    
	# time.sleep(2)   #Pause so the reader will read what is above !!!

