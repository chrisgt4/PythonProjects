import xlrd
import re
import os
import pyodbc

strFlLoc = r"G:\Internal Audit\Audit Working Papers\2015\2015023 Operating Systems (Unix) ~Technology~\CAATTs\Bindview\File Permissions"
flList = os.listdir(strFlLoc)
strFlNmConv = "File Perms"
flNms = []

for f in flList:
    flNmPat = re.compile("^.{1,100}" + strFlNmConv + r".{1,100}\.xlsx$")
    if flNmPat.match(f) != None:
        flNms.append(strFlLoc + "\\" + f)

cnn = pyodbc.connect(r"TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=R90GBCW2\SQLEXPRESS;DATABASE=Unix;")
cursor = cnn.cursor()

for f in flNms:
    with xlrd.open_workbook(f) as wb:
        for ws in wb.sheets():
            if ws.name != "Messages":
                for row in range(ws.nrows):
                    values = []
                    insertVals = []
                    for col in range(ws.ncols):
                        values.append(ws.cell(row,col).value)
                    if values[0] != "Machine Name":
                        if values[0] != '':
                            machNm = values[0]
                            flNmDir = values[1]
                            flNmPath = values[2]
                            flOwnr = values[3]
                            flGrp = values[4]
                            flPermTxt = values[5]
                            flPermOct = str(values[6])
                            lastA = values[7]
                            lastC = values[8]
                            lastM = values[9]
                            wExec = values[10]
                            wRead = values[11]
                            wWrite = values[12]
                            insertVals = [machNm,flNmDir,flNmPath,flOwnr,flGrp,flPermTxt,flPermOct,lastA,lastC,lastM,wExec,wRead,wWrite]
                            cleanVals = "'" + "','".join(vals.replace("'","''") for vals in insertVals) + "'"
                            sql = "INSERT INTO dbo.BV_FilePerms_111915 (MachineName,FileNameDirectory,FileNamePath,FileOwner,FileGroup,FilePermissionsText,FilePermissionsOctal"
                            sql = sql + ",LastAccessed,LastChanged,LastModified,WorldExecutable,WorldReadable,WorldWritable) "
                            sql = sql + "VALUES (" + cleanVals + ");"
                            cursor.execute(sql)
                            cursor.execute("Commit;")
cursor.close()
cnn.close()

print "***********************************Script Successfully Completed*************************************"
