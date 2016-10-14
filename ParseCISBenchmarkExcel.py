#ParseCISBenchmarkExcel Python script
# - written by Chris Curtis - Internal Audit
# - last revision on 7/5/2016
# - version 1.0 of the script
# - initial development of script
# - this version is written to read from an Excel file that was created from the CIS PDF version converted to Excel

#import module for sys objects
import sys
#add to PATH environment variable the location where the additional functions are located
#not being used right now so commented out
#sys.path.append('./Projects/')
#imports module for working with file systems and files
import os
#imports module for using regular expressions
import re
#imports module for dealing with csv files
import csv
#imports module for connecting to databases
import pyodbc
#imports module for reading excel files
import xlrd

#constant defining the pickup location
strFlPickupLoc = "F:\Python Projects\CIS"
#constant for defining the name of the file excluding those characters after the dash
strFlNmConv = "CIS_Amazon"

#get all files and directories in the specified location
flList = os.listdir(strFlPickupLoc)

#iterate through all files and directories
for f in flList:
    #create a regular expression to check the name of the file
    flNmPat = re.compile("^" + strFlNmConv + ".{1,100}" + "xlsx$")
    #check if file matches the regex pattern
    if flNmPat.match(f) != None:
        #if pattern matches then set flName to file name
        flName = strFlPickupLoc + "\\" + f

#define the connection to the SQL Server database to place data
cnxn = pyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server Native Client 10.0};SERVER=xxx;DATABASE=xxx;")
cnxn2 = pyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server Native Client 10.0};SERVER=xxx;DATABASE=xxx;")
#create a database cursor object to execute sql statements on this database
cursor = cnxn.cursor()
cursor2 = cnxn2.cursor()

#delete all current records in the AWS_Logging table
cursor.execute("DELETE FROM dbo.CIS_Benchmark_Text;")
#commit change to the database
cursor.execute("Commit;")

counter = 0

#open and set readFl to the identified file above
with xlrd.open_workbook(flName) as wb:
    for ws in wb.sheets():
        for row in range(ws.nrows):
            if ws.cell(row,0).ctype == 1:
                counter = counter + 1
                value = ws.cell(row,0).value
                cursor.execute("INSERT INTO dbo.CIS_Benchmark_Text (VarText,LineNum) VALUES (?,?);", (value, counter))
                cursor.execute("Commit;")

sql = "ALTER TABLE dbo.CIS_Benchmark_Text ALTER COLUMN LineNum INT;"
cursor.execute(sql)
cursor.execute("Commit;")

sql = "SELECT * FROM dbo.CIS_Benchmark_Text ORDER BY LineNum;"

ofile = open('CIS_Benchmark_Clean.csv',"wb")
writer = csv.writer(ofile, delimiter='|', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['Title','Profile','Description','Rationale','Remediation']
writer.writerow(header)

for row in cursor.execute(sql):
    curText = row[0].strip()
    titlePat = re.compile("^\d{1,2}\.\d{1,2}\.\d{1,2}.{1,200}|\d{1,2}\.\d{1,2}.{1,200}$")
    if titlePat.match(curText) != None:
        values = []
        titleRowStart = row[1]
        title = ''
        for row in cursor2.execute("SELECT * FROM dbo.CIS_Benchmark_Text WHERE LineNum >=? ORDER BY LineNum;", titleRowStart):
            otherText = row[0].strip()
            otherText = otherText.replace('\r','')
            otherText = otherText.replace('\n','')
            if otherText == "Profile Applicability:":
                break
            else:
                title = title + " " + otherText
        values.append(title)
    if curText == "Profile Applicability:":
        profRow = row[1] + 1
        cursor2.execute("SELECT * FROM dbo.CIS_Benchmark_Text WHERE LineNum = ?;", profRow)
        row2 = cursor2.fetchone()
        if row2:
            profile = row2[0]
            values.append(profile)
    if curText == "Description:":
        descRowStart = row[1] + 1
    if curText == "Rationale:":
        descRowEnd = row[1] - 1
        ratRowStart = row[1] + 1
        descText = ""
        for row3 in cursor2.execute("SELECT * FROM dbo.CIS_Benchmark_Text WHERE LineNum >=? AND LineNum <=?;", (descRowStart, descRowEnd)):
            descText = descText + " " + row3[0].strip()
        values.append(descText)
    if curText == "Audit:":
        ratRowEnd = row[1] - 1
        ratText = ""
        for row3 in cursor2.execute("SELECT * FROM dbo.CIS_Benchmark_Text WHERE LineNum >=? AND LineNum <=?;", (ratRowStart, ratRowEnd)):
            ratText = ratText + " " + row3[0].strip()
        values.append(ratText)
    if curText == "Remediation:":
        remRowStart = row[1] + 1
        remText = ""
        for row3 in cursor2.execute("SELECT * FROM dbo.CIS_Benchmark_Text WHERE LineNum >=? ORDER BY LineNum;", remRowStart):
            otherText = row3[0].strip()
            otherText = otherText.replace('\r','')
            otherText = otherText.replace('\n','')
            if titlePat.match(otherText) != None or otherText == "References:":
                break
            else:
                remText = remText + " " + otherText
        values.append(remText)
        writer.writerow(values)
                    
#close the database cursor
cursor.close()
cursor2.close()
#close the database connection
cnxn.close()
cnxn2.close()
#close the file object
ofile.close()

print "************************************Script Successful*********************************************"



