# - written by Chris Curtis - Internal Audit
# - last revision on 6/1/2016
# - version 1.0 of the script
# - initial development of script

import xlrd
import csv
import os
import re

hdr = ['DB_Name','Environment','DB_Status','Version_#','DB_Type','BusOrg','RCSExposure','RegSCI','ExternallyFacing','Tier2','Tier3','RCSScore','RCSDataSensitivity','Apps','SOX']

ofile = open('DB_Scope2.csv','wb')
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer.writerow(hdr)

#constant defining the pickup location for Gary Mikula's AWS Logging report
strFlPickupLoc = r"xxxx"
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
        if ws.name == "Oracle DB Inv 060116":
            for row in range(ws.nrows):
                if counter1 != 0:
                    values = []
                    apps = {}
                    for col in range(ws.ncols):
                        if col == 1:
                            dbNm = ws.cell(row,col).value
                            values.append(dbNm)
                            values.append(ws.cell(row,0).value)
                            values.append(ws.cell(row,3).value)
                            values.append(ws.cell(row,4).value)
                            values.append(ws.cell(row,5).value)
                            appCtr = 0
                            lWs = wb.sheet_by_name('App2Db')
                            for lrow in range(lWs.nrows):
                                if lrow != 0:
                                    for lcol in range(lWs.ncols):
                                        if lcol == 0:
                                            if lWs.cell(lrow,lcol).value == dbNm:
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



