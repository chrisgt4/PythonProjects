#Script to load Qualys raw data from XML
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 6/4/2015

import os
import re
import csv
import pyodbc

#define the directory to search for the file
strFlLoc = r"C:\Users\CurtisC\Desktop\qualys-raw-2015052000\April\Other"
flList = os.listdir(strFlLoc)

#find the xml file using regex to parse
for f in flList:
    flNmPat = re.compile(".{1,100}xml")
    if flNmPat.match(f) != None:
        flNm = strFlLoc + "\\" + f

lines = []

ofile = open('Qualys_Asset_Groups.csv',"wb")
ofile2 = open('Qualys_IP_Ranges.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer2 = csv.writer(ofile2, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['Asset_Group_Title']
header2 = ['IP_Range_Start','IP_Range_End']
writer.writerow(header)
writer2.writerow(header2)

#parse the xml file into a list object
with open(flNm, 'r') as f:
    for line in f:
        cleanline = line.strip()
        lines.append(cleanline)

print "finished loading xml to list"
isAstGrp = False
isIPRange = False
isIPDetail = False

for l in lines:
    L = []
    if l == "<USER_ASSET_GROUPS>":
        isAstGrp = True
    elif l == "</USER_ASSET_GROUPS>":
        isAstGrp = False
    if isAstGrp == True and l != "<USER_ASSET_GROUPS>":
        cleanl = l.replace("<ASSET_GROUP_TITLE><![CDATA[","")
        cleanl2 = cleanl.replace("]]></ASSET_GROUP_TITLE>","")
        L.append(cleanl2)
        writer.writerow(L)
    if l == "<COMBINED_IP_LIST>":
        isIPRange = True
    elif l == "</COMBINED_IP_LIST>":
        isIPRange = False
    if isIPRange == True and l != "</COMBINED_IP_LIST>":
        if l == "<RANGE>":
            isIPDetail = True
        elif l == "</RANGE>":
            isIPDetail = False
    if isIPRange == True and isIPDetail == True and l != "</COMBINED_IP_LIST>":
        if l.find("START", 0, len(l)) > 0:
            T = []
            cleanl = l.replace("<START>","")
            cleanl2 = cleanl.replace("</START>","")
            T.append(cleanl2)
        elif l.find("END", 0, len(l)) > 0:
            cleanl = l.replace("<END>","")
            cleanl2 = cleanl.replace("</END>","")
            T.append(cleanl2)
            writer2.writerow(T)

#close the file
f.close()
ofile.close()
ofile2.close()

print "*****Script successfully complete*****"
