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
# TODO: Nothing 
############################################################

import random
import time
import csv
import re

done = 0
bomreader = csv.reader(open('inputbom.csv','r'),delimiter=',',quotechar='"')
inmem = False		#This will keep track as to whether or not a particular entry is in 'memory' or not
BadFormat = False		#Use the regular expressions to determine whether or not we have good or bad format


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

