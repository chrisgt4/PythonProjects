#Script to parse SVN logs
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 10/12/16

import sys
import os
import re
import pypyodbc

#define the directory to search for the file
strFlLoc = r"xxxx"
flList = os.listdir(strFlLoc)

flNames = []

#find the xml file using regex to parse
for f in flList:
    flNmPat = re.compile(".{1,100}log")
    if flNmPat.match(f) != None:
        flNm = strFlLoc + "\\" + f
        flNames.append(flNm)

#setup database connection to local SQL Server
cnn = pypyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxxx;DATABASE=xxxx;")
cursor = cnn.cursor()

#clean the database table
sql = "DELETE FROM dbo.Log_Text;"
cursor.execute(sql)
cursor.execute("Commit;")

vals = []
counter = 0

for flNm in flNames:
    with open(flNm, 'r') as f:
        for line in f:
            counter += 1
            varList = []
            cleanline = line.strip()
            varList = line.split(" ")
            tstamp = varList[0][1:]
            usernm = varList[2]
            sql = "INSERT INTO dbo.Log_Text (ID,Time_Stamp,User_Name,Log_Text,File_Name) VALUES (?,?,?,?,?);"
            cursor.execute(sql,(counter,tstamp,usernm,cleanline,flNm))
            cursor.execute("Commit;")
    f.close()


#close the database connection
cursor.close()
cnn.close()

print "*****Script successfully complete*****"
