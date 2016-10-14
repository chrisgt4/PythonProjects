import ConfigParser

config = ConfigParser.ConfigParser()
config.readfp(open(r'CAP_Backup_Loading_Configs.txt'))

nbFailedJobsTblName = config.get('Data Load Configs','nbFailedJobsTblName')
nbSuccessfulJobsTblName = config.get('Data Load Configs','nbSuccessfulJobsTblName')
strFailedDropLoc = config.get('Data Load Configs','strFailedDropLoc')
strSuccessDropLoc  = config.get('Data Load Configs','strSuccessDropLoc')
strMoveLoc = config.get('Data Load Configs','strMoveLoc')
dataTruncLimit = config.get('Data Load Configs','dataTruncLimit')
iaDbUID = config.get('Data Load Configs','iaDbUID')
iaDbPwd = config.get('Data Load Configs','iaDbPwd')
ePickupLoc = config.get('Data Load Configs','emailPickupLoc')
eDropLoc = config.get('Data Load Configs','emailDropLoc')
wUser = config.get('Data Load Configs','writeUser')
iaDbServer = config.get('Data Load Configs','iaDbServer')
iaDbName = config.get('Data Load Configs','iaDbName')
emailTo = config.get('Data Load Configs','strTo')
emailFrom = config.get('Data Load Configs','strFrom')
emailSubject = config.get('Data Load Configs','strSubject')

def sendEmail(strFrom, strTo, strSubject, strBody, eSrc, eDest):
    import shutil
    f = open('email.txt','w')
    f.write('From: ' + strFrom + '\n')
    f.write('To: ' + strTo + '\n')
    f.write('Subject: ' + strSubject + '\n')
    f.write('\n')
    f.write(strBody)
    f.close()
    shutil.copy(eSrc,eDest)

def checkJobID(uid, pwd, iaDbSvr, iaDbNm, tbl, jID):
    import pyodbc

    cnxn = pyodbc.connect("driver={SQL Server};trusted_connection=Yes;server=" + iaDbSvr + ";database=" + iaDbNm + ";")
    cursor = cnxn.cursor()
    sql = "SELECT t.JobID FROM dbo." + tbl + " AS t WHERE t.JobID = '" + jID + "';"
    cursor.execute(sql)
    row = cursor.fetchone()
    if row == None:
        return None
    else:
        return row.JobID
    cursor.close()
    cnxn.close()

def getMaxFailedJobID(uid, pwd, iaDbSvr, iaDbNm, fTbl):
    import pyodbc

    cnxn = pyodbc.connect("driver={SQL Server};trusted_connection=Yes;server=" + iaDbSvr + ";database=" + iaDbNm + ";")
    cursor = cnxn.cursor()
    sql = "SELECT MAX(f.FailedJobID) As MaxID FROM dbo." + fTbl + " AS f;"
    cursor.execute(sql)
    row = cursor.fetchone()
    return row.MaxID
    cursor.close()
    cnxn.close()

def getMaxSuccessJobID(uid, pwd, iaDbSvr, iaDbNm, sTbl):
    import pyodbc

    cnxn = pyodbc.connect("driver={SQL Server};trusted_connection=Yes;server=" + iaDbSvr + ";database=" + iaDbNm + ";")
    cursor = cnxn.cursor()
    sql = "SELECT MAX(s.SuccessJobID) As MaxID FROM dbo." + sTbl + " AS s;"
    cursor.execute(sql)
    row = cursor.fetchone()
    return row.MaxID
    cursor.close()
    cnxn.close()
    
def getRowCount(csvfile):
    import csv
    counter = 0
    filereader = csv.reader(csvfile, delimiter=',', quotechar='"')
    for row in filereader:
        counter = counter + 1
    csvfile.close()
    return counter

def loadNBFailedJobs(curPath, uid, pwd, wUser, iaDbSvr, iaDbNm, truncLimit, fTbl, sTo, sFrom, sSubj, movePath):
    import csv
    import re
    import os
    import sys
    import pyodbc
    import datetime
    import shutil
    from simplecrypt import encrypt, decrypt
    from binascii import hexlify, unhexlify

    #uid = decrypt("xxxxxx",unhexlify(uid))
    #password = decrypt("xxxxxx",unhexlify(pwd))

    flList = os.listdir(curPath)
    counter = 0
    flNms = []
    
    for f in flList:
        flNmPat = re.compile("^OpsCenter_CAAT_Failed_Jobs_CJC_.{1,100}CSV$")
        if flNmPat.match(f) != None:
            counter = counter + 1
            flNms.append(curPath + f)
    if counter == 0:
        sendEmail(sFrom, sTo, sSubj, "No failed backup files available for loading.", ePickupLoc, eDropLoc)
        return
    try:
        cnxn = pyodbc.connect("driver={SQL Server};trusted_connection=Yes;server=" + iaDbSvr + ";database=" + iaDbNm + ";")
        cursor = cnxn.cursor()

        flNumber = 0
        
        for f in flNms:
            counter = 0
            loadCount = 0
            maxFjId = getMaxFailedJobID(uid, pwd, iaDbSvr, iaDbNm, fTbl)
            exitPat = re.compile("^Report generated on.{1,100}$")
            if maxFjId == None:
                maxFjId = 0
            with open(f, 'rb') as readFl:
                rowcount = getRowCount(readFl) - 5
            with open(f, 'rb') as readFl2:
                if rowcount >= truncLimit:
                    sendEmail(sFrom, sTo, sSubj, f + " has possible truncated data.", ePickupLoc, eDropLoc)
                    readFl.close()
                else:
                    reader = csv.reader(readFl2)
                    for row in reader:
                        counter = counter + 1
                        if counter >= 5:
                            if exitPat.match(str(row[0])) == None:
                                if checkJobID(uid, pwd, iaDbSvr, iaDbNm, fTbl, row[3]) == None:
                                    rightnowstr = str(datetime.datetime.now()).split(".")
                                    if row[2] == ' ':
                                        row[2] = "Blank"
                                    maxFjId = maxFjId + 1
                                    loadCount = loadCount + 1
                                    sql = "INSERT INTO dbo." + fTbl + " (FailedJobID,ClientName,PolicyName,SchedName,JobID,SchedType,ExitStatus,JobStart,JobEnd,StatusCode,WriteUser,WriteDate) " \
                                          "VALUES (" + str(maxFjId) + \
                                          ",'" + row[0] + \
                                          "','" + row[1] + \
                                          "','" + row[2] + \
                                          "','" + row[3] + \
                                          "','" + row[4] + \
                                          "','" + row[5] + \
                                          "','" + row[6] + \
                                          "','" + row[7] + \
                                          "','" + row[8] + \
                                          "','" + wUser + \
                                          "','" + rightnowstr[0] + "');"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
            sendEmail(sFrom, sTo, sSubj, "CAP Backup Data Loading Success.\n" + f + " successfully loaded " + str(loadCount) + " lines of data.", ePickupLoc, eDropLoc)
            readFl2.close()
            shutil.move(f, movePath + flList[flNumber])
            flNumber = flNumber + 1
    except:
        sendEmail(sFrom, sTo, sSubj, "CAP Backup Data Loading Error.\n" + f + " failed loading with error" + str(sys.exc_info()[0]), ePickupLoc, eDropLoc)        

    cursor.close()
    cnxn.close()
    
def loadNBSuccessfulJobs(curPath, uid, pwd, wUser, iaDbSvr, iaDbNm, truncLimit, sTbl, sTo, sFrom, sSubj, movePath):
    import csv
    import re
    import os
    import sys
    import pyodbc
    import datetime
    import shutil
    from simplecrypt import encrypt, decrypt
    from binascii import hexlify, unhexlify

    #uid = decrypt("xxxxxx",unhexlify(uid))
    #password = decrypt("xxxxxx",unhexlify(pwd))

    flList = os.listdir(curPath)
    counter = 0
    flNms = []
    
    for f in flList:
        flNmPat = re.compile("^OpsCenter_CAAT_Successful_Jobs_CJC_.{1,100}CSV$")
        if flNmPat.match(f) != None:
            counter = counter + 1
            flNms.append(curPath + f)
    if counter == 0:
        sendEmail(sFrom, sTo, sSubj, "No failed backup files available for loading.", ePickupLoc, eDropLoc)
        return
    try:
        cnxn = pyodbc.connect("driver={SQL Server};trusted_connection=Yes;server=" + iaDbSvr + ";database=" + iaDbNm + ";")
        cursor = cnxn.cursor()

        flNumber = 0
        
        for f in flNms:
            counter = 0
            loadCount = 0
            maxSjId = getMaxSuccessJobID(uid, pwd, iaDbSvr, iaDbNm, sTbl)
            exitPat = re.compile("^Report generated on.{1,100}$")
            if maxSjId == None:
                maxSjId = 0
            with open(f, 'rb') as readFl:
                rowcount = getRowCount(readFl) - 5
            with open(f, 'rb') as readFl2:
                if rowcount >= truncLimit:
                    sendEmail(sFrom, sTo, sSubj, f + " has possible truncated data.", ePickupLoc, eDropLoc)
                    readFl.close()
                else:
                    reader = csv.reader(readFl2)
                    for row in reader:
                        counter = counter + 1
                        if counter >= 5:
                            if exitPat.match(str(row[0])) == None:
                                if checkJobID(uid, pwd, iaDbSvr, iaDbNm, sTbl, row[3]) == None:
                                    rightnowstr = str(datetime.datetime.now()).split(".")
                                    if row[2] == ' ':
                                        row[2] = "Blank"
                                    maxSjId = maxSjId + 1
                                    loadCount = loadCount + 1
                                    sql = "INSERT INTO dbo." + sTbl + " (SuccessJobID,ClientName,PolicyName,SchedName,JobID,SchedType,ExitStatus,JobStart,JobEnd,StatusCode,WriteUser,WriteDate) " \
                                          "VALUES (" + str(maxSjId) + \
                                          ",'" + row[0] + \
                                          "','" + row[1] + \
                                          "','" + row[2] + \
                                          "','" + row[3] + \
                                          "','" + row[4] + \
                                          "','" + row[5] + \
                                          "','" + row[6] + \
                                          "','" + row[7] + \
                                          "','" + row[8] + \
                                          "','" + wUser + \
                                          "','" + rightnowstr[0] + "');"
                                    cursor.execute(sql)
                                    cursor.execute("Commit;")
            sendEmail(sFrom, sTo, sSubj, "CAP Backup Data Loading Success.\n" + f + " successfully loaded " + str(loadCount) + " lines of data.", ePickupLoc, eDropLoc)
            readFl2.close()
            shutil.move(f, movePath + flList[flNumber])
            flNumber = flNumber + 1
    except:
        sendEmail(sFrom, sTo, sSubj, "CAP Backup Data Loading Error.\n" + f + " failed loading with error" + str(sys.exc_info()[0]), ePickupLoc, eDropLoc)        

    cursor.close()
    cnxn.close()


loadNBFailedJobs(strFailedDropLoc, iaDbUID, iaDbPwd, wUser, iaDbServer, iaDbName, dataTruncLimit, nbFailedJobsTblName, emailTo, emailFrom, emailSubject, strMoveLoc)
loadNBSuccessfulJobs(strSuccessDropLoc, iaDbUID, iaDbPwd, wUser, iaDbServer, iaDbName, dataTruncLimit, nbSuccessfulJobsTblName, emailTo, emailFrom, emailSubject, strMoveLoc)
  
