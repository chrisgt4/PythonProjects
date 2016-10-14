import ConfigParser

config = ConfigParser.ConfigParser()
config.readfp(open(r'CAP_IP_Mapping_Win_Configs.txt'))

loadTblName = config.get('Data Load Configs','loadTblName')
strDropLoc = config.get('Data Load Configs','strDropLoc')
ePickupLoc = config.get('Data Load Configs','emailPickupLoc')
eDropLoc = config.get('Data Load Configs','emailDropLoc')
wUser = config.get('Data Load Configs','writeUser')
iaDbServer = config.get('Data Load Configs','iaDbServer')
iaDbName = config.get('Data Load Configs','iaDbName')
emailTo = config.get('Data Load Configs','strTo')
emailFrom = config.get('Data Load Configs','strFrom')
emailSubject = config.get('Data Load Configs','strSubject')
iaDbUID = config.get('Access Keys','iaDbUID')
uidKey = config.get('Access Keys','uidKey')
uidIv = config.get('Access Keys','uidIv')
iaDbPwd = config.get('Access Keys','iaDbPwd')
pwdKey = config.get('Access Keys','pwdKey')
pwdIv = config.get('Access Keys','pwdIv')

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

def cleanTbl(uid, pwd, iaDbSvr, iaDbNm, tbl):
    import pyodbc

    cnxn = pyodbc.connect("driver={SQL Server};uid=" + uid + ";pwd=" + pwd + ";server=" + iaDbSvr + ";database=" + iaDbNm + ";")
    cursor = cnxn.cursor()
    sql = "DELETE FROM dbo." + tbl + ";"
    cursor.execute(sql)
    cursor.execute("Commit;")
    cursor.close()
    cnxn.close()

def loadIPMappingWin(curPath, uid, uidKey, uidIv, pwd, pwdKey, pwdIv, wUser, iaDbSvr, iaDbNm, tbl, sTo, sFrom, sSubj):
    import csv
    import re
    import os
    import sys
    import pyodbc
    import datetime
    import shutil
    from Crypto.Cipher import AES
    from binascii import hexlify, unhexlify

    try:
        cipher = AES.new(uidKey, AES.MODE_CBC, unhexlify(uidIv))
        username = cipher.decrypt(unhexlify(uid))
        cleanUid = username.rstrip("0") 

        cipher2 = AES.new(pwdKey, AES.MODE_CBC, unhexlify(pwdIv))
        password = cipher2.decrypt(unhexlify(pwd))
        cleanPwd = password.rstrip("%")

        cnxn = pyodbc.connect("driver={SQL Server};uid=" + cleanUid + ";pwd=" + cleanPwd + ";server=" + iaDbSvr + ";database=" + iaDbNm + ";")
        cursor = cnxn.cursor()

        flList = os.listdir(curPath)
        counter = 0
        flNms = []
        
        for f in flList:
            flNmPat = re.compile("^IP_Mapping_Windows.{1,100}csv$")
            if flNmPat.match(f) != None:
                counter = counter + 1
                flNms.append(curPath + '\\' + f)
        if counter == 0:
            sendEmail(sFrom, sTo, sSubj, "No failed backup files available for loading.", ePickupLoc, eDropLoc)
            return

        cleanTbl(cleanUid, cleanPwd, iaDbSvr, iaDbNm, tbl)
        
        cnxn = pyodbc.connect("driver={SQL Server};uid=" + cleanUid + ";pwd=" + cleanPwd + ";server=" + iaDbSvr + ";database=" + iaDbNm + ";")
        cursor = cnxn.cursor()

        flNumber = 0
        
        for f in flNms:
            counter = 0
            loadCount = 0
            with open(f, 'rb') as readFl:
                reader = csv.reader(readFl)
                for row in reader:
                    counter = counter + 1
                    if counter >= 2:
                        loadCount = loadCount + 1
                        sql = "INSERT INTO dbo." + tbl + " (IP_Addr,NetBIOS,DNSName,OS,Domain,AD_Site,Description) " \
                              "VALUES ('" + row[0] + \
                              "','" + row[1] + \
                              "','" + row[2] + \
                              "','" + row[3] + \
                              "','" + row[4] + \
                              "','" + row[5] + \
                              "','" + row[6].replace("'","''") + "');"
                        cursor.execute(sql)
                        cursor.execute("Commit;")
            sendEmail(sFrom, sTo, sSubj, "Windows IP Mapping Data Loading Success.\n" + f + " successfully loaded " + str(loadCount) + " lines of data.", ePickupLoc, eDropLoc)
            readFl.close()
            os.remove(f)
            flNumber = flNumber + 1
        cursor.close()
        cnxn.close()
    except:
        sendEmail(sFrom, sTo, sSubj, "CAP Backup Data Loading Error.\n" + f + " failed loading with error" + str(sys.exc_info()[0]), ePickupLoc, eDropLoc)       

loadIPMappingWin(strDropLoc, iaDbUID, uidKey, uidIv, iaDbPwd, pwdKey, pwdIv, wUser, iaDbServer, iaDbName, loadTblName, emailTo, emailFrom, emailSubject)
