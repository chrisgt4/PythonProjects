# - written by Chris Curtis - Internal Audit
# - last revision on 6/1/2016
# - version 1.0 of the script
# - initial development of script

import xlrd
import csv
import os
import re

hdr = ['SvrName','Instance','Version','Patch','Domain','WindowsOS','BusOrg','RCSExposure','RegSCI','ExternallyFacing','Tier2','Tier3','RCSScore','RCSDataSensitivity','Apps','SOX']

ofile = open('DB_Scope_SQLSvr.csv','wb')
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer.writerow(hdr)

#constant defining the pickup location for Gary Mikula's AWS Logging report
strFlPickupLoc = r"xxxxxxx"
#constant for defining the name of the file excluding those characters after the dash
strFlNmConv = "Database Population"

#get all files and directories in the specified location
flList = os.listdir(strFlPickupLoc)

#iterate through all files and directories
for f in flList:
    #create a regular expression to check the name of the file
    flNmPat = re.compile("^" + strFlNmConv + ".{1,9}" + "xlsx$")
    #check if file matches the regex pattern
    if flNmPat.match(f) != None:
        #if pattern matches then set flName to file name
        flName = strFlPickupLoc + "\\" + f

counter1 = 0

#open and set readFl to the identified file above
with xlrd.open_workbook(flName) as wb:
    for ws in wb.sheets():
        if ws.name == "SQLSvr":
            for row in range(ws.nrows):
                if counter1 != 0:
                    values = []
                    apps = {}
                    for col in range(ws.ncols):
                        if col == 7:
                            svrNmLookup = ws.cell(row,col).value
                            values.append(ws.cell(row,6).value)
                            values.append(ws.cell(row,8).value)
                            values.append(ws.cell(row,9).value)
                            values.append(ws.cell(row,10).value)
                            values.append(ws.cell(row,11).value)
                            values.append(ws.cell(row,12).value)
                            appCtr = 0
                            lWs = wb.sheet_by_name('Lookup for SQLSvr')
                            for lrow in range(lWs.nrows):
                                if lrow != 0:
                                    for lcol in range(lWs.ncols):
                                        if lcol == 0:
                                            if lWs.cell(lrow,lcol).value == svrNmLookup:
                                                if appCtr == 0:
                                                    apps['appNm'] = []
                                                    apps['busOrg'] = []
                                                    apps['t2'] = []
                                                    apps['t3'] = []
                                                    apps['rcsDS'] = []
                                                    apps['rcsE'] = []
                                                    apps['rcsS'] = []
                                                    apps['sox'] = []
                                                    apps['ef'] = []
                                                    apps['regSci'] = []
                                                    apps['appNm'].append(lWs.cell(lrow,1).value)
                                                    apps['busOrg'].append(lWs.cell(lrow,2).value)
                                                    apps['t2'].append(lWs.cell(lrow,3).value)
                                                    apps['t3'].append(lWs.cell(lrow,4).value)
                                                    apps['rcsDS'].append(lWs.cell(lrow,5).value)
                                                    apps['rcsE'].append(lWs.cell(lrow,6).value)
                                                    apps['rcsS'].append(lWs.cell(lrow,7).value)
                                                    apps['sox'].append(lWs.cell(lrow,8).value)
                                                    apps['ef'].append(lWs.cell(lrow,9).value)
                                                    apps['regSci'].append(lWs.cell(lrow,10).value)
                                                else:
                                                    if lWs.cell(lrow,1).value not in apps['appNm']:
                                                        apps['appNm'].append(lWs.cell(lrow,1).value)
                                                        apps['busOrg'].append(lWs.cell(lrow,2).value)
                                                        apps['t2'].append(lWs.cell(lrow,3).value)
                                                        apps['t3'].append(lWs.cell(lrow,4).value)
                                                        apps['rcsDS'].append(lWs.cell(lrow,5).value)
                                                        apps['rcsE'].append(lWs.cell(lrow,6).value)
                                                        apps['rcsS'].append(lWs.cell(lrow,7).value)
                                                        apps['sox'].append(lWs.cell(lrow,8).value)
                                                        apps['ef'].append(lWs.cell(lrow,9).value)
                                                        apps['regSci'].append(lWs.cell(lrow,10).value)
                                                appCtr += 1
                            if not len(apps) == 0:
                                for key in apps:
                                    print key
                                    varStr = ''
                                    aCtr = 0
                                    l = apps[key]
                                    for val in l:
                                        if aCtr == 0:
                                            varStr = val
                                        else:
                                            varStr = varStr + ', ' + val
                                        aCtr += 1
                                    values.append(varStr)
                            writer.writerow(values)
                counter1 += 1
ofile.close()
print "--------------------------------------------------Script successfully run-------------------------------------------------------------"



