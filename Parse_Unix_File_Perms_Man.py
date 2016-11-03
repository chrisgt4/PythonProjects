#Script to load Unix file permission data into a SQL Server database
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 12/8/2015
#commented out lines are for writing to a csv instead of to a database, can be recommented back in

import sys
sys.path.append('./Lib/site-packages/')
sys.path.append('F:/Python Projects/')
import os
import re
import csv
import pypyodbc

#define the directory to search for the file
strFlLoc = r"xxxx"
flList = os.listdir(strFlLoc)
#strFlNmConv = "xxxxx"
flNms = []

#find the text files using regex to parse
for f in flList:
    flNmPat = re.compile(".{1,100}txt")
    if flNmPat.match(f) != None:
        flNms.append(strFlLoc + "\\" + f)

#setup database connection to local SQL Server
cnn = pypyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxxx;")
cursor = cnn.cursor()

#clean the database table
sql = "DELETE FROM dbo.Man_FilePerms_112515;"
cursor.execute(sql)

#uncomment if writing to csv
#ofile = open('Qualys_Data.csv',"wb")
#writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
#header = ['IP_Addr','DNS_Name','Operating_System','Netbios']
#writer.writerow(header)

#parse that Unix text data like a bad habit
for f in flNms:
    varList1 = f.split(".")
    varTxt1 = varList1[0]
    varList1 = varTxt1.split("\\")
    lSize = len(varList1)
    machNm = varList1[lSize - 1]
    lines = []
    with open(f, 'r') as fl:
        for line in fl:
            cleanline = line.strip()
            lines.append(cleanline)
    for l in lines:
        if l[:1] == "/":
            flNmPath = l[:-1]
        if l[:1] == "-":
            vTxt = l.split(" ")
            cleanT = []
            for v in vTxt:
                if v != "":
                    cleanT.append(v)
            flPermTxt = cleanT[0]
            flOwner = cleanT[2]
            flGroup = cleanT[3]
            lastMod = cleanT[5] + " " + cleanT[6] + " " + cleanT[7]
            flName = cleanT[8]
            if flPermTxt[9] == "x":
                worldExec = "yes"
            else:
                worldExec = "no"
            if flPermTxt[7] == "r":
                worldRead = "yes"
            else:
                worldRead = "no"
            if flPermTxt[8] == "w":
                worldWrite = "yes"
            else:
                worldWrite = "no"
            cursor.execute("INSERT INTO dbo.Man_FilePerms_112515 (MachineName,FileNamePath,FilePermissionsText,FileOwner,FileGroup,LastModified,FileName,WorldExecutable,WorldReadable,WorldWritable) VALUES (?,?,?,?,?,?,?,?,?,?);",
            (machNm,flNmPath,flPermTxt,flOwner,flGroup,lastMod,flName,worldExec,worldRead,worldWrite))
            cursor.execute("Commit;")
            #writer.writerow(L)
    
#close the file
#f.close()
#ofile.close()

#close the database connection
cursor.close()
cnn.close()

print "*****Script successfully complete*****"
