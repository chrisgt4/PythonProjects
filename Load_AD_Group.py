import xlrd
import re
import os
import pyodbc

strFlLoc = r"xxx"
flList = os.listdir(strFlLoc)
strFlNmConv = "Group Documentation_"
flNms = []

for f in flList:
    flNmPat = re.compile("^" + strFlNmConv + r".{1,100}\.xls$")
    if flNmPat.match(f) != None:
        flNms.append(strFlLoc + "\\" + f)

cnn = pyodbc.connect(r"TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxxx;DATABASE=xxx;")
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
                        grpNm = values[2]
                        qualNm = values[3]
                        if "\n" in values[4]:
                            memberNms = values[4].split("\n")
                            for mNm in memberNms:
                                memberNm = mNm
                                if memberNm != '':
                                    if memberNm != '[None Found]':
                                        if ":" in memberNm:
                                            cleanMemNm = memberNm.split(":")
                                            iaGrpMem = cleanMemNm[1]
                                        else:
                                            iaGrpMem = memberNm
                                        if "\\" in memberNm:
                                            try:
                                                iaGrpMems = iaGrpMem.split("\\")
                                                iaGrpMemClean = iaGrpMems[1]
                                            except IndexError:
                                                iaGrpMemClean = memberNm
                                        else:
                                            iaGrpMemClean = memberNm
                                    else:
                                        iaGrpMem = memberNm
                                        iaGrpMemClean = memberNm
                                    insertVals = [domain,machNm,grpNm,qualNm,memberNm,iaGrpMem,iaGrpMemClean]
                                    cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                    sql = "INSERT INTO dbo.Group_To_User_123115 (Domain_Name,Machine_Name,Group_Name,Fully_Qual_Name,Group_Member,IA_Group_Member,IA_Group_Member_Clean) "
                                    sql = sql + "VALUES (" + cleanVals + ");"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
                        else:
                            memberNm = values[4]
                            if memberNm != '':
                                if memberNm != '[None Found]':
                                    if ":" in memberNm:
                                        cleanMemNm = memberNm.split(":")
                                        iaGrpMem = cleanMemNm[1]
                                    else:
                                        iaGrpMem = memberNm
                                    if "\\" in memberNm:
                                        try:
                                            iaGrpMems = iaGrpMem.split("\\")
                                            iaGrpMemClean = iaGrpMems[1]
                                        except IndexError:
                                            iaGrpMemClean = memberNm
                                    else:
                                        iaGrpMemClean = memberNm
                                else:
                                    iaGrpMem = memberNm
                                    iaGrpMemClean = memberNm
                                insertVals = [domain,machNm,grpNm,qualNm,memberNm,iaGrpMem,iaGrpMemClean]
                                cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                sql = "INSERT INTO dbo.Group_To_User_123115 (Domain_Name,Machine_Name,Group_Name,Fully_Qual_Name,Group_Member,IA_Group_Member,IA_Group_Member_Clean) "
                                sql = sql + "VALUES (" + cleanVals + ");"
                                cursor.execute(sql)
                                cursor.execute("Commit;")
                    else:
                        if "\n" in values[4]:
                            memberNms = values[4].split("\n")
                            for mNm in memberNms:
                                memberNm = mNm
                                if memberNm != '':
                                    if memberNm != '[None Found]':
                                        if ":" in memberNm:
                                            cleanMemNm = memberNm.split(":")
                                            iaGrpMem = cleanMemNm[1]
                                        else:
                                            iaGrpMem = memberNm
                                        if "\\" in memberNm:
                                            try:
                                                iaGrpMems = iaGrpMem.split("\\")
                                                iaGrpMemClean = iaGrpMems[1]
                                            except IndexError:
                                                iaGrpMemClean = memberNm
                                        else:
                                            iaGrpMemClean = memberNm
                                    else:
                                        iaGrpMem = memberNm
                                        iaGrpMemClean = memberNm
                                    insertVals = [domain,machNm,grpNm,qualNm,memberNm,iaGrpMem,iaGrpMemClean]
                                    cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                    sql = "INSERT INTO dbo.Group_To_User_123115 (Domain_Name,Machine_Name,Group_Name,Fully_Qual_Name,Group_Member,IA_Group_Member,IA_Group_Member_Clean) "
                                    sql = sql + "VALUES (" + cleanVals + ");"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
                        else:
                            memberNm = values[4]
                            if memberNm != '':
                                if memberNm != '[None Found]':
                                    if ":" in memberNm:
                                        cleanMemNm = memberNm.split(":")
                                        iaGrpMem = cleanMemNm[1]
                                    else:
                                        iaGrpMem = memberNm
                                    if "\\" in memberNm:
                                        try:
                                            iaGrpMems = iaGrpMem.split("\\")
                                            iaGrpMemClean = iaGrpMems[1]
                                        except IndexError:
                                            iaGrpMemClean = memberNm
                                    else:
                                        iaGrpMemClean = memberNm
                                else:
                                    iaGrpMem = memberNm
                                    iaGrpMemClean = memberNm
                                insertVals = [domain,machNm,grpNm,qualNm,memberNm,iaGrpMem,iaGrpMemClean]
                                cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                                sql = "INSERT INTO dbo.Group_To_User_123115 (Domain_Name,Machine_Name,Group_Name,Fully_Qual_Name,Group_Member,IA_Group_Member,IA_Group_Member_Clean) "
                                sql = sql + "VALUES (" + cleanVals + ");"
                                cursor.execute(sql)
                                cursor.execute("Commit;")
cursor.close()
cnn.close()

print "***********************************Script Successfully Completed*************************************"
