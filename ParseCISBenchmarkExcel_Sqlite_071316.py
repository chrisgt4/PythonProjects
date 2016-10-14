#ParseCISBenchmarkExcel Python script
# - written by Chris Curtis - Internal Audit
# - last revision on 7/13/2016
# - version 1.0 of the script
# - initial development of script
# - this version is written to read from an Excel file that was created from the CIS PDF version converted to Excel with Adobe Acrobat
# - this version also is written to use a sqlite database backend


#imports module for working with file systems and files
import os
#imports module for using regular expressions
import re
#imports module for dealing with csv files
import csv
#imports module for reading excel files
import xlrd
#imports module for sqlite database
import sqlite3

#constant defining the pickup location - adjust to location where CIS file is located and where to create the sqlite database
#also define the name of the Access database file
strFlPickupLoc = "F:\Python Projects\CIS"
strDbName = 'CIS_Db.sqlite'
#constant for defining the name of the file excluding those characters after the dash
strFlNmConv = "abcd"

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

#create new sqlite database
cnn = sqlite3.connect(strDbName)
cursor = cnn.cursor()
cursor2 = cnn.cursor()

#create table to drop data and delete all records if it exists
cursor.execute('CREATE TABLE if not exists CIS_Benchmark_Text (VarText VARCHAR, LineNum INT)')
cursor.execute('DELETE FROM CIS_Benchmark_Text;')

counter = 0

#open and set readFl to the identified file above
with xlrd.open_workbook(flName) as wb:
    for ws in wb.sheets():
        for row in range(ws.nrows):
            if ws.cell(row,0).ctype == 1:
                counter = counter + 1
                value = ws.cell(row,0).value
                cursor.execute("INSERT INTO CIS_Benchmark_Text (VarText,LineNum) VALUES (?,?);", (value, counter))

#sql = "ALTER TABLE CIS_Benchmark_Text ALTER COLUMN LineNum INT;"
#cursor.execute(sql)

sql = "SELECT * FROM CIS_Benchmark_Text ORDER BY LineNum;"

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
        for row in cursor2.execute("SELECT * FROM CIS_Benchmark_Text WHERE LineNum >=? ORDER BY LineNum;", (titleRowStart,)):
            otherText = row[0].strip()
            otherText = otherText.replace('\r','')
            otherText = otherText.replace('\n','')
            if otherText == "Profile Applicability:":
                break
            else:
                title = title + " " + otherText
        values.append(title.encode('utf-8'))
    if curText == "Profile Applicability:":
        profRow = row[1] + 1
        cursor2.execute("SELECT * FROM CIS_Benchmark_Text WHERE LineNum = ?;", (profRow,))
        row2 = cursor2.fetchone()
        if row2:
            profile = row2[0].encode('utf-8')
            values.append(profile)
    if curText == "Description:":
        descRowStart = row[1] + 1
    if curText == "Rationale:":
        descRowEnd = row[1] - 1
        ratRowStart = row[1] + 1
        descText = ""
        for row3 in cursor2.execute("SELECT * FROM CIS_Benchmark_Text WHERE LineNum >=? AND LineNum <=?;", (descRowStart, descRowEnd)):
            descText = descText + " " + row3[0].strip()
        values.append(descText.encode('utf-8'))
    if curText == "Audit:":
        ratRowEnd = row[1] - 1
        ratText = ""
        for row3 in cursor2.execute("SELECT * FROM CIS_Benchmark_Text WHERE LineNum >=? AND LineNum <=?;", (ratRowStart, ratRowEnd)):
            ratText = ratText + " " + row3[0].strip()
        values.append(ratText.encode('utf-8'))
    if curText == "Remediation:":
        remRowStart = row[1] + 1
        remText = ""
        for row3 in cursor2.execute("SELECT * FROM CIS_Benchmark_Text WHERE LineNum >=? ORDER BY LineNum;", (remRowStart,)):
            otherText = row3[0].strip()
            otherText = otherText.replace('\r','')
            otherText = otherText.replace('\n','')
            if titlePat.match(otherText) != None or otherText == "References:":
                break
            else:
                remText = remText + " " + otherText
        values.append(remText.encode('utf-8'))
        print values
        writer.writerow(values)
                    
#close the database cursor
cursor.close()
cursor2.close()
#close the database connection
cnn.commit()
cnn.close()
#close the file object
ofile.close()

print "************************************Script Successful*********************************************"



