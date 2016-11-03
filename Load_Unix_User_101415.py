import xlrd
import re
import os
import pyodbc

strFlLoc = r"xxxx"
flList = os.listdir(strFlLoc)
strFlNmConv = "SOX Unix Local Users"
flNms = []

for f in flList:
    flNmPat = re.compile("^" + strFlNmConv + r".{1,100}\.xls$")
    if flNmPat.match(f) != None:
        flNms.append(strFlLoc + "\\" + f)

cnn = pyodbc.connect(r"TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxx;")
cursor = cnn.cursor()

for f in flNms:
    with xlrd.open_workbook(f) as wb:
        for ws in wb.sheets():
            for row in range(ws.nrows):
                values = []
                insertVals = []
                for col in range(ws.ncols):
                    values.append(ws.cell(row,col).value)
                if values[0] != "Machine Name":
                    if values[0] != '':
                        machNm = values[0]
                        userNm = values[1]
                        uid = values[2]
                        uDb = values[3]
                        uInfo = values[4]
                        gid = values[6]
                        lastLogin = values[7]
                        homePath = values[8]
                        acctLocked = values[9]
                        if "\n" in values[5]:
                            groupNms = values[5].split("\n")
                            for gNm in groupNms:
                                groupNm = gNm
                                if groupNm != '':
                                    insertVals = [machNm,userNm,uid,uDb,uInfo,groupNm,gid,lastLogin,homePath,acctLocked]
                                    cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                    sql = "INSERT INTO dbo.Unix_User_100915 (Machine_Name,User_Name,User_ID,User_Database,User_Information,Group_Name,Group_ID,Last_Login,Home_Path,Account_Locked) "
                                    sql = sql + "VALUES (" + cleanVals + ");"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
                        else:
                            groupNm = values[5]
                            if groupNm != '':
                                insertVals = [machNm,userNm,uid,uDb,uInfo,groupNm,gid,lastLogin,homePath,acctLocked]
                                cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                sql = "INSERT INTO dbo.Unix_User_100915 (Machine_Name,User_Name,User_ID,User_Database,User_Information,Group_Name,Group_ID,Last_Login,Home_Path,Account_Locked) "
                                sql = sql + "VALUES (" + cleanVals + ");"
                                cursor.execute(sql)
                                cursor.execute("Commit;")
                    else:
                        if "\n" in values[5]:
                            groupNms = values[5].split("\n")
                            for gNm in groupNms:
                                groupNm = gNm
                                if groupNm != '':
                                    insertVals = [machNm,userNm,uid,uDb,uInfo,groupNm,gid,lastLogin,homePath,acctLocked]
                                    cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                    sql = "INSERT INTO dbo.Unix_User_100915 (Machine_Name,User_Name,User_ID,User_Database,User_Information,Group_Name,Group_ID,Last_Login,Home_Path,Account_Locked) "
                                    sql = sql + "VALUES (" + cleanVals + ");"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
                        else:
                            groupNm = values[5]
                            if groupNm != '':
                                insertVals = [machNm,userNm,uid,uDb,uInfo,groupNm,gid,lastLogin,homePath,acctLocked]
                                cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                sql = "INSERT INTO dbo.Unix_User_100915 (Machine_Name,User_Name,User_ID,User_Database,User_Information,Group_Name,Group_ID,Last_Login,Home_Path,Account_Locked) "
                                sql = sql + "VALUES (" + cleanVals + ");"
                                cursor.execute(sql)
                                cursor.execute("Commit;")
cursor.close()
cnn.close()

print "***********************************Script Successfully Completed*************************************"
