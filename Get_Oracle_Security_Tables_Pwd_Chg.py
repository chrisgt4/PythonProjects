# Written by Chris Curtis
# Last updated on 8/1/2016
# Version 1.0
# Script to pull security tables from Oracle databases
# tables include:
# USER$

import pypyodbc
import cx_Oracle

uid = 'xxxx'
#redacted the password used for security purposes
pwd = 'xxx'

dblist = ['xxxx']
hostlist = ['xxxxx']

sqlCnn = pypyodbc.connect('TRUSTED_CONNECTION=Yes;DRIVER={SQL Server Native Client 10.0};SERVER=xxx;DATABASE=xxxx;')
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
    ver = oraCnn.version
    try:
        # loop for DBA_USERS
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_USERS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                v6 = row[5]
                v7 = row[6]
                v8 = row[7]
                v9 = row[8]
                v10 = row[9]
                v11 = row[10]
                v12 = row[11]
                v13 = row[12]
                v14 = row[13]
                v15 = row[14]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_USERS (DBNAME,VERSION,USERNAME,USER_ID,PASSWORD,ACCOUNT_STATUS,LOCK_DATE,EXPIRY_DATE,DEFAULT_TABLESPACE,TEMPORARY_TABLESPACE,CREATED,PROFILE,INITIAL_RSRC_CONSUMER_GROUP,EXTERNAL_NAME,
                                  PASSWORD_VERSIONS,EDITIONS_ENABLED,AUTHENTICATION_TYPE)
                                  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5,v6,v7,v8,v9,v10,v11,v12,v13,v14,v15))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_USERS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                v6 = row[5]
                v7 = row[6]
                v8 = row[7]
                v9 = row[8]
                v10 = row[9]
                v11 = row[10]
                v12 = row[11]
                v13 = row[12]
                v14 = row[13]
                v15 = row[14]
                v16 = row[15]
                v17 = row[16]
                v18 = row[17]
                v19 = row[18]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_USERS (DBNAME,VERSION,USERNAME,USER_ID,PASSWORD,ACCOUNT_STATUS,LOCK_DATE,EXPIRY_DATE,DEFAULT_TABLESPACE,TEMPORARY_TABLESPACE,CREATED,PROFILE,INITIAL_RSRC_CONSUMER_GROUP,EXTERNAL_NAME,
                                  PASSWORD_VERSIONS,EDITIONS_ENABLED,AUTHENTICATION_TYPE,PROXY_ONLY_CONNECT,COMMON,LAST_LOGIN,ORACLE_MAINTAINED)
                                  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5,v6,v7,v8,v9,v10,v11,v12,v13,v14,v15,v16,v17,v18,v19))
                sqlCursor.execute("COMMIT;")
                ctr += 1

        # loop for DBA_PROFILES
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_PROFILES'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_PROFILES (DBNAME,VERSION,PROFILE,RESOURCE_NAME,RESOURCE_TYPE,LIMIT)
                                  VALUES (?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_PROFILES'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_PROFILES (DBNAME,VERSION,PROFILE,RESOURCE_NAME,RESOURCE_TYPE,LIMIT,COMMON)
                                  VALUES (?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5))
                sqlCursor.execute("COMMIT;")
                ctr += 1

        # loop for DBA_TAB_PRIVS
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_TAB_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                v6 = row[5]
                v7 = row[6]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_TAB_PRIVS (DBNAME,VERSION,GRANTEE,OWNER,TABLE_NAME,GRANTOR,PRIVILEGE,GRANTABLE,HIERARCHY)
                                  VALUES (?,?,?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5,v6,v7))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_TAB_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                v6 = row[5]
                v7 = row[6]
                v8 = row[7]
                v9 = row[8]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_TAB_PRIVS (DBNAME,VERSION,GRANTEE,OWNER,TABLE_NAME,GRANTOR,PRIVILEGE,GRANTABLE,HIERARCHY,COMMON,TYPE)
                                  VALUES (?,?,?,?,?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4,v5,v6,v7,v8,v9))
                sqlCursor.execute("COMMIT;")
                ctr += 1

        # loop for DBA_SYS_PRIVS
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_SYS_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_SYS_PRIVS (DBNAME,VERSION,GRANTEE,PRIVILEGE,ADMIN_OPTION)
                                  VALUES (?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_SYS_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_SYS_PRIVS (DBNAME,VERSION,GRANTEE,PRIVILEGE,ADMIN_OPTION,COMMON)
                                  VALUES (?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4))
                sqlCursor.execute("COMMIT;")
                ctr += 1

        # loop for DBA_ROLE_PRIVS
        if ver[:2]=='11':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_ROLE_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_ROLE_PRIVS (DBNAME,VERSION,GRANTEE,GRANTED_ROLE,ADMIN_OPTION,DEFAULT_ROLE)
                                  VALUES (?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v4))
                sqlCursor.execute("COMMIT;")
                ctr += 1
        elif ver[:2]=='12':
            ctr = 0
            for row in oraCursor.execute('SELECT * FROM sys.DBA_ROLE_PRIVS'):
                v1 = row[0]
                v2 = row[1]
                v3 = row[2]
                v4 = row[3]
                v5 = row[4]
                v6 = row[5]
                sqlCursor.execute("""
                                  INSERT INTO dbo.DBA_ROLE_PRIVS (DBNAME,VERSION,GRANTEE,GRANTED_ROLE,ADMIN_OPTION,DEFAULT_ROLE,DELEGATE_OPTION,COMMON)
                                  VALUES (?,?,?,?,?,?,?,?);
                                  """,
                                  (db,ver,v1,v2,v3,v5,v4,v6))
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