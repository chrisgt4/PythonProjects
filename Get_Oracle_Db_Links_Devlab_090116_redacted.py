# Written by Chris Curtis
# Last updated on 9/1/2016
# Version 1.0
# Script to pull security tables from Oracle databases
# tables include:
# DBA_DB_LINKS

import pypyodbc
import cx_Oracle

uid = 'xxxxxx'
#redacted the password used for security purposes
pwd = 'xxxxxxxxxxxx'

dblist = ['xxxx']
hostlist = ['xxxxx']

sqlCnn = pypyodbc.connect('TRUSTED_CONNECTION=Yes;DRIVER={SQL Server Native Client 10.0};SERVER=xxx;DATABASE=xxx;')
sqlCursor = sqlCnn.cursor()

a = 0
for db in dblist:
    host = hostlist[a]
    sid = dblist[a]
    if db == 'INSP':
        try:
            oraCnn = cx_Oracle.connect(uid + '/' + pwd + '@' + host + ':1526/' + sid)
            oraCursor = oraCnn.cursor()
        except:
            print db + ' - cannot connect'
    else:
        try:
            oraCnn = cx_Oracle.connect(uid + '/' + pwd + '@' + host + ':1521/' + sid)
            oraCursor = oraCnn.cursor()
        except:
            print db + ' - cannot connect'
    try:
        ver = oraCnn.version
        # loop for DBA_DB_LINKS
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_DB_LINKS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_DB_LINKS (DBNAME,VERSION,OWNER,DB_LINK,USERNAME,HOST,CREATED)
                                  VALUES (?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_DB_LINKS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_DB_LINKS (DBNAME,VERSION,OWNER,DB_LINK,USERNAME,HOST,CREATED)
                                  VALUES (?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5))
                sqlCursor.execute("COMMIT;")
                ctr += 1
    except:
        print db + ' - insufficient privilege'
    a += 1
    try:
        oraCursor.close()
        oraCnn.close()
    except:
        print db + ' - not open'

sqlCursor.close()
sqlCnn.close()
print "----------------------------------------Script Successfully Completed-----------------------------------------------------"