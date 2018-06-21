############################################################
# FILE: Combined_BOM_v0D3.py
#
# PURPOSE: Pull in a BOM that contains many duplicate entries, 
#          consolidate entries and update quantities, then 
#          spit out a new BOM with consolidated entries.
#
# AUTHOR: Clinton J Guenther
#
# REV HISTORY: 	See Text File
#
# TODO: This is far from being complete!  Goal is to have this
#		script autmatically combine multible BOM (excel files).
#		This code does not run, at the moment.  
############################################################
import xlrd
import xlwt
import random
import time
import csv
import re
import os
import sys

## DEFINE VRIABLES ##
#####################
MFGPN_col = 0			#Column number containing the MFGPN
QPN_col = 0				#Column number containing QPN
MFG_col = 0				#Column location for manufacturer part number
DES_col = 0 			#Column location for descritpion part number
QTY_col = 0 			#Column location for quantity field

Search_Hits = 0 		#This number is incremented when we find values we are searching for in a particular row
Max_Hits = 5			#This many hits will trigger us to leave the searching loop
data_start = 0			#This is the row where the data strats

header = []		#This array will define the colum locations for the header
qpn = []        #Pull in all QPNs into a list. This will make them easier to work with later
asso = []       #Pull in all QPNs into a list. This will make them easier to work with later
qty = []        #Pull in all QTYs into a list. This will make them easier to work with later
des = []		#Pull in all Descriptions into a list. This will make them easier to work with later
mfg = []		#Pull in all Manufactures into a list. This will make them easier to work with later
mfgpn = []		#Pull in all Manufacturing Part Numbers into a list. This will make them easier to work with later

path = os.getcwd()
## DEFINE THE FILES THAT THE USER WISHES TO OPEN ##
###################################################
for (path, dirs, files) in os.walk(path):
       path
       dirs
       files

## DEF FUNCTION ##
######################	   
def debugbreak():
	while(1):
		a=1
		
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
	   
## MAIN LOOP ##
######################

print "Files found in directory: " + str(len(files))

for i in range(len(files)-1):
	
	if(files[i].find("xls") != -1):
		workbook = xlrd.open_workbook(files[i])     #Open the workbook that we are going to parse though 
		worksheets = workbook.sheet_names()             #Grab the names of the worksheets -- I believe this line is critical.
		num_sheets = len(worksheets)				#This is the number of sheet
		
		print "\n\n==============================================="
		print "==============================================="
		print "File opened: " + str(files[i])
		print "The number of worksheets is: " + str(num_sheets)
		association = raw_input("Enter the product name for this BOM (i.e. TWC): ") 
		print "+++++++++++++++++++++++++++++++++++++++++++++++"
		for sh in range (num_sheets):
			current_sheet = workbook.sheet_by_name(worksheets[sh])		#Grab the first worksheet
			
			num_rows = current_sheet.nrows - 1      #This is how many rows are in the worksheet
			num_cols = current_sheet.ncols - 1 		#This is how many columns are in the sheet

			for curr_row in range (num_rows + 1):			#Find the header locations
				row = current_sheet.row(curr_row)					#Grab the current row
				Search_Hits = 0
				for curr_col in range (num_cols + 1):
					temptext = str(row[curr_col])                  #This is the fifth cell of the current row
					temptext = temptext.lstrip('text:u\'')     #Remove the initial part of the string that we don't need 'text:u'   
					temptext = temptext.replace(" ","")			#This will remove any and all white spaces
					if(temptext.find("QPN")!=-1):
						#print "QPN column found to be: " + str(curr_col)
						QPN_col = curr_col
						Search_Hits += 1
					elif(temptext.find("MFGPN")!=-1):							#Look for MFGPN header
						#print "MFGPN column found to be: " + str(curr_col)
						MFGPN_col = curr_col
						Search_Hits += 1
						#print "Search Hits: " + str(Search_Hits)
					elif(temptext.find("MFG")!=-1 and temptext.find("PN") == -1):		#Look for MFG -- make sure PN is not present
						#print "MFG column found to be: " + str(curr_col)
						MFG_col = curr_col
						Search_Hits += 1
						#print "Search Hits: " + str(Search_Hits)
					elif(temptext.find("Des")!=-1 or temptext.find("DES")!=-1):		#Look for MFG -- make sure PN is not present
						#print "Description column found to be: " + str(curr_col)
						DES_col = curr_col
						Search_Hits += 1
						#print "Search Hits: " + str(Search_Hits)
					elif(temptext.find("Qty")!=-1 or temptext.find("QTY")!=-1):		#Look for MFG -- make sure PN is not present
						#print "Description column found to be: " + str(curr_col)
						QTY_col = curr_col
						Search_Hits += 1
						#print "Search Hits: " + str(Search_Hits)
				if(Search_Hits >= Max_Hits):
					data_start = curr_row + 1
					break
			
			print "QPN column found to be: " + str(QPN_col)		
			print "Description column found to be: " + str(DES_col)		
			print "MFG column found to be: " + str(MFG_col)
			print "MFGPN column found to be: " + str(MFGPN_col)
			print "QTY column found to be: " + str(QTY_col)
			
			header = [0,QPN_col,DES_col,MFG_col,MFGPN_col,QTY_col]
			header_values = ["Association","QPN","DES","MFG","MFGPN","QTY"]
			
			for curr_row in range (data_start,num_rows + 1):
				row = current_sheet.row(curr_row)					#Grab the current row
				asso.append(association)				#For each row in the BOM, we need to append the association
				print ".",
				#print row[QPN_col]
				#debugbreak()
				current_value = clean_value(str(row[QPN_col]))
				qpn.append(current_value)			
				
				current_value = clean_des(str(row[DES_col]))
				des.append(current_value)
				
				current_value = clean_value(str(row[MFG_col]))
				mfg.append(current_value)
				
				current_value = clean_value(str(row[MFGPN_col]))
				mfgpn.append(current_value)
				
				current_value = clean_value(str(row[QTY_col]))
				qty.append(current_value)
#print qpn
#print asso	
#print "Length of association: " + str(len(asso))
#debugbreak()
ob = xlwt.Workbook()						#Now save out the combined BOM
Sheet1 = ob.add_sheet("Sheet1")	#Add a sheet to our new workbook
#Sheet1.write(<row>,<column>,<information>)
#Sheet1.write(0,0,"Test")

print "+++++++++++++++++++++++++++++++++++++++++++++++"
print "+++++++++++++++++++++++++++++++++++++++++++++++"
print "Creating combined BOM"
for i in range (len(header)):
	Sheet1.write(0,i,header_values[i])
for i in range (1,len(asso)):				#Vertical movement 
	Sheet1.write(i,0,asso[i])
	#print qpn[i]
	Sheet1.write(i,1,qpn[i])
	Sheet1.write(i,2,des[i])
	Sheet1.write(i,3,mfg[i])
	Sheet1.write(i,4,mfgpn[i])
	Sheet1.write(i,5,qty[i])
	print ".",

ob.save("CombinedBOM.xls")
time.sleep(3)
sys.exit()
debugbreak()

	
### This is where I stopped	 
print '------------------------'
print 'Did you read the Readme.txt file first???'


qpn = []        #Suck all QPNs into a list. This will make them easier to work with later
qty = []        #Suck all QTYs into a list. This will make them easier to work with later
des = []		#Suck all Descriptions into a list. This will make them easier to work with later
mfg = []		#Suck all Manufactures into a list. This will make them easier to work with later
mfgpn = []		#Suck all Manufacturing Part Numbers into a list. This will make them easier to work with later
cr1 = []		#Suck all Distributors into a list. This will make them easier to work with later
cr1pn = []		#Suck all Distributor Part Numbers into a list. This will make them easier to work with later

memqpn = []   	#The user will only get a particular QPN and MFGPN copied to 'memory' once. Only QTY is then updated
memqty = []     #QTY is updated when a part getting sucked into memory is already present
memdes = []		#'Memory' field that keeps track of descriptions
memmfg = []		#'Memory' field that keeps track of Manufactures
memmfgpn = []	#The user will only get a particular QPN and MFGPN copied to 'memory' once. Only QTY is then updated
memcr1 = []		#'Memory' field that keeps track of descriptions
memcr1pn = []	#'Memory' field that keeps track of distributor Part Numbers

# Get all parameter stored in an array
for row in bomreader:
    
    if row[0] != "QPN" and row[1] != "QTY":                  #Do not copy the header to memory!        
        qpn.append(row[0])              #QPN parameter
        qty.append(float(row[1]))       #Quantity parameter
        des.append(row[2])              #Description parameter
        mfg.append(row[3])              #MFG parameter
        mfgpn.append(row[4])            #MFGPN parameter
        cr1.append(row[5])              #Supplier parameter
        cr1pn.append(row[6])            #Supplier parameter part number



qpn_re = re.compile('[0-9xX]{3}-[0-9xX]{6}-[0-9xX]{4}') #we need to make sure we have valid QPNs in the BOM
for myqpn in qpn:
    if len(qpn_re.findall(myqpn)) == 0:
        BadFormat = True
        print 'Found QPN that is of incorrect format!!!'
        
for i in range(0,len(qpn)):     #This is the line we are currently looking at NOT IN 'Memory'
    if len(memqpn) == 0:       	#this should only run once when there are no entries in the memory qpn
        memqpn.append(qpn[i])   
        memqty.append(qty[i])
        memdes.append(des[i])
        memmfg.append(mfg[i])
        memmfgpn.append(mfgpn[i])
        memcr1.append(cr1[i])
        memcr1pn.append(cr1pn[i])
                
    else:
        for k in range (0,len(memqpn)):		#Go through all the entries in 'memory' and verify that the MFGPN and QPN does not already exist. 
            if qpn[i] == memqpn[k]:     # Check against all entries in memory
                if mfgpn[i] == memmfgpn[k]: #Double check to make sure the MFGPN matches
                    memqty[k] = memqty[k] + qty[i]      #Increase the QTY that is listed in mem
                    inmem = True;
        if inmem == False:							#If the particular line item does not exist then put item in 'memory'
            memqpn.append(qpn[i])
            memqty.append(qty[i])
            memdes.append(des[i])
            memmfg.append(mfg[i])
            memmfgpn.append(mfgpn[i])
            memcr1.append(cr1[i])
            memcr1pn.append(cr1pn[i])
        inmem = False                   #Reset this after being set true from above.

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

print '------------------------'
tmpline1 = "The old BOM was " + str(len(qpn)) + " line items long."
tmpline2 = "The new BOM is only " + str(len(memqpn)) + " line items long."
print tmpline1
print tmpline2
print '------------------------'

file_out = 'combined_BOM.csv'
fid_out = open(file_out,'w') # 'w' this creates a new file if it does not already exist

for i in range (0,len(memqpn)):
    tmpline = ('\"' + memqpn[i] + '\",\"' + str(memqty[i]) + '\","' + memdes[i] + '\",\"' + memmfg[i] + '\",\"' + memmfgpn[i] + '\",\"'
    + memcr1[i] + '\",\"' + memcr1pn[i] + '\"' + '\n')
    fid_out.write(tmpline)

fid_out.close()

print 'FINISHED'    
time.sleep(2)   #Pause so the reader will read what is above !!!

