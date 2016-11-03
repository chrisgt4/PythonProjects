# Written by Chris Curtis
# Last updated on 6/13/2016
# Version 1.0
# Script to pull security tables from Oracle databases
# tables to pull from
# DBA_USERS_WITH_DEFPWD

import pypyodbc
import cx_Oracle

uid = 'xxxx'
pwd = 'xxxxx'

dblist = ['xxxx']
hostlist = ['xxxxxx']

sqlCnn = pypyodbc.connect('TRUSTED_CONNECTION=Yes;DRIVER={SQL Server Native Client 10.0};SERVER=R90GBCW2\SQLEXPRESS;DATABASE=DBMS2;')
sqlCursor = sqlCnn.cursor()

a = 0
for db in dblist:
    host = hostlist[a]
    sid = dblist[a]
    if db == 'INSP':
        oraCnn = cx_Oracle.connect(uid + '/' + pwd + '@' + host + ':1526/' + sid)
        oraCursor = oraCnn.cursor()
    else:
        oraCnn = cx_Oracle.connect(uid + '/' + pwd + '@' + host + ':1521/' + sid)
        oraCursor = oraCnn.cursor()
    ver = oraCnn.version
    # loop for DBA_USERS_WITH_DEFPWD
    if ver[:2]=='11':
        ctr = 0
        for row in oraCursor.execute('SELECT * FROM sys.DBA_USERS_WITH_DEFPWD'):
            v1 = row[0]
            sqlCursor.execute("""
                              INSERT INTO dbo.DBA_USERS_WITH_DEFPWD (DBNAME,VERSION,USERNAME)
                              VALUES (?,?,?);
                              """,
                              (db,ver,v1))
            sqlCursor.execute("COMMIT;")
            ctr += 1
    elif ver[:2]=='12':
        ctr = 0
        for row in oraCursor.execute('SELECT * FROM sys.DBA_USERS_WITH_DEFPWD'):
            v1 = row[0]
            sqlCursor.execute("""
                              INSERT INTO dbo.DBA_USERS_WITH_DEFPWD (DBNAME,VERSION,USERNAME)
                              VALUES (?,?,?);
                              """,
                              (db,ver,v1))
            sqlCursor.execute("COMMIT;")
            ctr += 1
    a += 1
    oraCursor.close()
    oraCnn.close()

sqlCursor.close()
sqlCnn.close()
print "----------------------------------------Script Successfully Completed-----------------------------------------------------"