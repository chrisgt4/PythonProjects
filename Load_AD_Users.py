import xlrd
import re
import os
import pyodbc

strFlLoc = r"xxxx"
flList = os.listdir(strFlLoc)
strFlNmConv = "SOX AD Users Term Tool"
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
                if values[0] != "Domain/Workgroup Name":
                    if values[0] != '':
                        domain = values[0]
                        machNm = values[1]
                        uNm = values[2]
                        fullNm = values[3]
                        firstNm = values[4]
                        lastNm = values[5]
                        fullyQualNm = values[6]
                        acctPrivLvl = values[7]
                        userType = values[8]
                        acctDesc = values[9]
                        acctDis = values[10]
                        lastLogin = values[11]
                        createDt = values[12]
                        insertVals = [domain,machNm,uNm,fullNm,firstNm,lastNm,fullyQualNm,acctPrivLvl,userType,acctDesc,acctDis,lastLogin,createDt]
                        cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                        sql = "INSERT INTO dbo.AD_User_123115 (Domain_Name,Machine_Name,User_Name,Full_Name,First_Name,Last_Name,Fully_Qual_Name,Acct_Priv_Level"
                        sql = sql + ",User_Type,Account_Desc,Account_Disabled,Last_Login,Create_Date) "
                        sql = sql + "VALUES (" + cleanVals + ");"
                        cursor.execute(sql)
                        cursor.execute("Commit;")
cursor.close()
cnn.close()

print "***********************************Script Successfully Completed*************************************"
