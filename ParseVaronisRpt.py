import os
import csv
import re
import pyodbc

strFlLoc = r"xxxxxx"
flList = os.listdir(strFlLoc)

for f in flList:
    flNmPat = re.compile(".{1,100}csv")
    if flNmPat.match(f) != None:
        flNm = strFlLoc + "\\" + f

cnn = pyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxx;")
cursor = cnn.cursor()

with open(flNm,"r") as readFl:
    reader = csv.reader(readFl)
    columns = next(reader)
    counter = 0
    for data in reader:
        fSrv = data[0]
        aPath = data[1]
        usrGrp = data[2]
        acctNm = data[3]
        perms = data[4]
        inher = data[5]

        rootFlPat = re.compile(r"^[A-Z]:\\(?:[A-Z0-9\-]{1,20}\\Users|\w{1,20}\\[A-Z0-9\-]{1,20}\\Users)$")
        if rootFlPat.match(aPath) != None:
            isRootFol = "Y"
        else:
            isRootFol = "N"
        fDrivePat = re.compile(r"^[A-Z]:\\(?:[A-Z0-9\-]{1,20}\\Users\\" + acctNm + r"|\w{1,20}\\[A-Z0-9\-]{1,20}\\Users\\" + acctNm + ")$")
        if fDrivePat.match(aPath) != None:
            isFDrv = "Y"
        else:
            isFDrv = "N"
        fConPat = re.compile("^.{0,7}F.{0,7}$")
        if fConPat.match(perms) != None:
            fControl = "Y"
        else:
            fControl = "N"
        modPat = re.compile("^.{0,7}M.{0,7}$")
        if modPat.match(perms) != None:
            modify = "Y"
        else:
            modify = "N"
        rFlPat = re.compile("^.{0,7}R.{0,7}$")
        if rFlPat.match(perms) != None:
            rFile = "Y"
        else:
            rFile = "N"
        wFlPat = re.compile("^.{0,7}W.{0,7}$")
        if wFlPat.match(perms) != None:
            wFile = "Y"
        else:
            wFile = "N"
        eFlPat = re.compile("^.{0,7}X.{0,7}$")
        if eFlPat.match(perms) != None:
            xFile = "Y"
        else:
            xFile = "N"
        lFlPat = re.compile("^.{0,7}L.{0,7}$")
        if lFlPat.match(perms) != None:
            lFile = "Y"
        else:
            lFile = "N"
        specPat = re.compile("^.{0,7}S.{0,7}$")
        if specPat.match(perms) != None:
            spec = "Y"
        else:
            spec = "N"
        
        sql = "INSERT INTO dbo.F_Drive_Permissions_Update (File_Server, Access_Path, Is_Root_Folder, Is_F_Drive_Folder, User_Group_Fully_Qual_Name, User_Group_Clean_Name, Full_Control, Modify, Read_File, "
        sql = sql + "Write, Read_Execute, List_Folder_Contents, Special, Permissions, Inherited_From_Folder) VALUES ("
        sql = sql + "'" + fSrv + "',"
        sql = sql + "'" + aPath.replace("'","''") + "',"
        sql = sql + "'" + isRootFol + "',"
        sql = sql + "'" + isFDrv + "',"
        sql = sql + "'" + usrGrp.replace("'","''") + "',"
        sql = sql + "'" + acctNm + "',"
        sql = sql + "'" + fControl + "',"
        sql = sql + "'" + modify + "',"
        sql = sql + "'" + rFile + "',"
        sql = sql + "'" + wFile + "',"
        sql = sql + "'" + xFile + "',"
        sql = sql + "'" + lFile + "',"
        sql = sql + "'" + spec + "',"
        sql = sql + "'" + perms + "',"
        sql = sql + "'" + inher + "');"
        print sql
        cursor.execute(sql)
        cursor.execute("Commit;")

cursor.close()
cnn.close()

print "**************************Script Successfully Completed********************************"
