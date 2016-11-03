import os
import csv
import re
import pyodbc

strFlLoc = r"xxxx"
flList = os.listdir(strFlLoc)
flNms = []

for f in flList:
    flNmPat = re.compile(".{1,100}csv")
    if flNmPat.match(f) != None:
        flNms.append(strFlLoc + "\\" + f)

cnn = pyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxxx;DATABASE=xxx;")
cursor = cnn.cursor()

cursor.execute("DELETE FROM dbo.Tripwire_Build_Checks;")
cursor.execute("Commit;")

for f in flNms: 
    with open(f,"r") as readFl:
        reader = csv.reader(readFl)
        columns = next(reader)
        counter = 0
        for data in reader:
            nNm = data[0]
            nType = data[1]
            pol = data[2]
            parTstGrp = data[3]
            tstNm = data[4]
            descr = data[5]
            elemnt = data[6]
            resTime = data[7]
            resState = data[8]
            actVal = data[9]

            cursor.execute("INSERT INTO dbo.Tripwire_Build_Checks (Node_Name,Node_Type,Policy,Parent_Test_Group,Test_Name,Description,Element,Result_Time,Result_State,Actual_Value) VALUES (?,?,?,?,?,?,?,?,?,?);",
            (nNm,nType,pol,parTstGrp,tstNm,descr,elemnt,resTime,resState,actVal))
            cursor.execute("Commit;")
    readFl.close()

cursor.close()
cnn.close()

print "**************************Script Successfully Completed********************************"
