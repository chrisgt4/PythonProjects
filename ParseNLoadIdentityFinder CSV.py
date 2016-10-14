#Identity Finder CSV Parser Script
# - written by Chris Curtis - Internal Audit
# - last revision on 2/19/16
# - version 1.0 of the script
# - version 1.0 release notes
#     - initial development of script
#     - performs ETL on the CSV file generated by Identity Finder
#     - loads the transformed data into a SQL Server database

#imports module for working with file systems and files
import os
#imports module for using regular expressions
import re
#imports module for dealing with csv files
import csv
#imports module for connecting to databases
import pypyodbc
#imports module to deal with dates
from datetime import datetime
#import module for dealing with the local system
import sys

#constant defining the pickup location for Gary Mikula's AWS Logging report
strFlPickupLoc = r"xxxx"
#constant for defining the name of the file excluding those characters after the dash
strFlNmConv = "Full Scan Results"

#get all files and directories in the specified location
flList = os.listdir(strFlPickupLoc)

#iterate through all files and directories
for f in flList:
    #create a regular expression to check the name of the file
    flNmPat = re.compile(strFlNmConv + ".{1,9}" + "csv")
    #check if file matches the regex pattern
    if flNmPat.match(f) != None:
        #if pattern matches then set flName to file name
        flName = strFlPickupLoc + "\\" + f

#define the connection to the SQL Server database to place data
cnxn = pypyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxx;")
#create a database cursor object to execute sql statements on this database
cursor = cnxn.cursor()

#cleanup tables
cursor.execute("DELETE FROM dbo.IF_Files_Email;")
cursor.execute("Commit;")
cursor.execute("DELETE FROM dbo.IF_Detail_Email;")
cursor.execute("Commit;")

#open and set readFl to the identified file above and begin parsing the file
with open(flName,"r") as readFl:
    reader = csv.reader(readFl)
    #get rid of the column headers that don't need to be loaded
    next(reader)
    next(reader)
    next(reader)
    next(reader)
    #iterate through all the data in the csv file
    counter = 0
    for data in reader:
        if len(data[0]) > 0:
            counter = counter + 1
            flID = counter
            flType = data[0]
            loc = data[1]
            #locs = loc.split("\\")
            #newlocs = locs[2:len(locs) - 1]
            #l = "\\"
            #newloc = l.join(newlocs)
            if data[2] == "Unknown":
                dtMod = None
            else:
                dtMod = datetime.strptime(data[2],"%m/%d/%Y")
            size = data[3]
            mtCat = data[4]
            mtDet = data[5]
            mtCnt = int(data[6])
            #dynamically build a sql statement to insert data to IF_Files table
            cursor.execute("INSERT INTO dbo.IF_Files_Email (File_ID,File_Type,Location,Date_Mod,Size,Match_Cat,Match_Det,Match_Count) VALUES (?,?,?,?,?,?,?,?);",(flID,flType,loc,dtMod,size,mtCat,mtDet,mtCnt))
            cursor.execute("Commit;")
        #dynamically build a sql statement to insert data to IF_Detail table
        mt2Cat = data[4]
        mt2Det = data[5]
        mt2Cnt = int(data[6])
        cursor.execute("INSERT INTO dbo.IF_Detail_Email (File_ID,Match_Type,Match_Detail,Match_Count) VALUES (?,?,?,?);",(flID,mt2Cat,mt2Det,mt2Cnt))
        cursor.execute("Commit;")

#close the database cursor
cursor.close()
#close the database connection
cnxn.close()
                       
print "----------------------------------Script Successfully Run---------------------------------------------------------"




