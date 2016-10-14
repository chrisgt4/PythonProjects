#Test script to query Active Directory using pyad and putting the data into SQL Server
#Written by Chris Curtis - Internal Audit
#Last Updated on 9/8/16

import ldap
import sys
import pypyodbc
import datetime
import struct

from ldap.controls import SimplePagedResultsControl
from distutils.version import StrictVersion
from optparse import OptionParser

LDAP24API = StrictVersion(ldap.__version__) >= StrictVersion('2.4')
domains = ['xxxxx']

try:
    cnn = pypyodbc.connect(r"TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxx;")
    cursor = cnn.cursor()
    cursor2 = cnn.cursor()

    #cursor.execute("DELETE FROM dbo.AD_Groups;")
    #cursor.execute("Commit;")
    #cursor.execute("DELETE FROM dbo.AD_Grp2UserMap;")
    #cursor.execute("Commit;")
    #cursor.execute("DELETE FROM dbo.AD_Grp2GrpMap;")
    #cursor.execute("Commit;")
except:
    print >> sys.stderr, 'Connection to database failed.'
    sys.exit()

def convert(binary):
    version = struct.unpack('B', binary[0])[0]
    # I do not know how to treat version != 1 (it does not exist yet)
    assert version == 1, version
    length = struct.unpack('B', binary[1])[0]
    authority = struct.unpack('>Q', '\x00\x00' + binary[2:8])[0]
    string = 'S-%d-%d' % (version, authority)
    binary = binary[8:]
    assert len(binary) == 4 * length
    for i in xrange(length):
        value = struct.unpack('<L', binary[4*i:4*(i+1)])[0]
        string += '-%d' % (value)
    return string

def findMaxKey():
    cursor.execute("SELECT MAX(GrpID) AS MaxKey FROM dbo.AD_Groups;")
    for row in cursor.fetchall():
        maxKey = row[0]
    if maxKey == None:
        uCounter = 0
    else:
        uCounter = maxKey
    return uCounter

def getSIDName(sid):
    cursor.execute("SELECT User_Name, User_DN FROM dbo.AD_Users WHERE SID = ?;",(sid,))
    if cursor.rowcount == 0:
        userNm = None
        userDn = None
    else:
        for row in cursor.fetchall():
            userNm = row[0]
            userDn = row[1]
    return userNm, userDn

def getUserName(userdn):
    #crit = (userdn,)
    cursor.execute("SELECT User_Name FROM dbo.AD_Users WHERE User_DN = ?;",(userdn,))
    if cursor.rowcount == 0:
        userNm = None
    else:
        for row in cursor.fetchall():
            userNm = row[0]
    return userNm

def getGrpName(grpNm):
    #crit = (grpNm,)
    cursor.execute("SELECT Group_Name FROM dbo.AD_Groups WHERE Group_DN = ?;",(grpNm,))
    if cursor.rowcount == 0:
        groupNm = None
    else:
        for row in cursor.fetchall():
            groupNm = row[0]
    return groupNm

def conv_AD_time(adtime):
    secs = adtime / 10000000
    conv = secs - 11644473600
    dt = datetime.datetime(2000,1,1,0,0)
    return dt.fromtimestamp(conv)

def create_controls(pagesize):
    if LDAP24API:
        return SimplePagedResultsControl(True,size=page_size,cookie='')
    else:
        return SimplePagedResultsControl(ldap.LDAP_CONTROL_PAGE_OID, True,
                                         (pagesize,''))
def get_pctrls(serverctrls):
    if LDAP24API:
        return [c for c in serverctrls
                if c.controlType == SimplePagedResultsControl.controlType]
    else:
        return [c for c in serverctrls
                if c.controlType == ldap.LDAP_CONTROL_PAGE_OID]

def set_cookie(lc_object,pctrls,pagesize):
    if LDAP24API:
        cookie = pctrls[0].cookie
        lc_object.cookie = cookie
        return cookie
    else:
        est, cookie = pctrls[0].controlValue
        lc_object.controlValue = (pagesize,cookie)
        return cookie
def clean_List(list, key):
    if key in list:
        tempL = []
        tempL = list[key]
        val = ""
        for t in tempL:
            val = val + t
        return val
    else:
        return None

def getGrpID(memberDN):
    cursor.execute("SELECT GrpID FROM dbo.AD_Groups WHERE Group_DN = ?;",(memberDN,))
    if cursor.rowcount == 0:
        grpID = None
    else:
        for row in cursor.fetchall():
            grpID = row[0]
    return grpID

def checkForGrp2Grp(grpID):
    cursor.execute("SELECT * FROM dbo.AD_Grp2UserMap WHERE Member_Name IS NULL AND GrpID = ?;",(grpID,))
    grp2grpList = []
    for row in cursor.fetchall():
        val = row[4]
        grp2grpList.append(val)
    return grp2grpList

def checkForMapping(parentGrp, childGrp, fromGrp):
    cursor.execute("SELECT * FROM dbo.AD_Grp2GrpMap WHERE GrpID = ? AND GrpID_Member = ? AND GrpID_From = ?;",(parentGrp,childGrp,fromGrp))
    if len(cursor.fetchall()) > 0:
        return "Yes"
    else:
        return "No"

def updateProcessedFlag(rowID):
    cursor.execute("UPDATE dbo.AD_Grp2GrpMap SET Processed_Flag = 'Y' WHERE RowID = ?;",(rowID,))
    cursor.execute("Commit;")

for d in domains:
    if d == 'xxx':
        baseDN = "DC=xx,DC=xx,DC=xxx,DC=xx"
    elif d == 'xxx':
        baseDN = "DC=xxx,DC=xxx,DC=xxx"
    elif d == 'xxx':
        baseDN = "DC=xxx,DC=xxx,DC=xxx"
    searchScope = ldap.SCOPE_SUBTREE
    retrieveAttributes = None
    searchFilter = "(&(objectclass=group)(!(objectclass=computer)))"
    page_size = 10
    dNm = d
      
    try:
        ldap.set_option(ldap.OPT_REFERRALS, 0)
        if d == 'xxx':
            l = ldap.initialize("ldap://xxxx:389/")
        elif d == 'xxx':
            l = ldap.initialize("ldap://xxxx:389/")
        elif d == 'xxx':
            l = ldap.initialize("ldap://xxxx:389/")
        l.protocol_version = ldap.VERSION3
        if d == 'xxx':
            username = "CN=xxx,OU=xx,OU=xx,DC=xx,DC=xx,DC=xx,DC=xx"
            password = "xxxx"
        elif d == 'xxx':
            username = "CN=xxx,OU=xxxx,OU=xxx,DC=xxx,DC=xxx,DC=xxx"
            password = "xxx"
        elif d == 'xxx':
            username = "CN=xxx,OU=xxx,OU=xx,DC=xx,DC=xx,DC=xx,DC=xx"
            password = "xxxxx"
        #line below is to use TLS connection for LDAP, which is currently not enabled or allowed
        #l.start_tls_s()
        l.simple_bind_s(username, password)
    except ldap.LDAPError as e:
        print e
        sys.exit()
        
    lc = create_controls(page_size)
    grpID = findMaxKey()

    while True:
        try:
            ldap_result_id = l.search_ext(baseDN, searchScope, searchFilter, retrieveAttributes, serverctrls=[lc])
        except ldap.LDAPError as e:
            print e
            sys.exit()
        try:
            result_type, result_data, rmsgid, serverctrls = l.result3(ldap_result_id)
        except ldap.LDAPError as e:
            print e
            sys.exit()
        for dn, entry in result_data:
            if dn != None:
                grpID = grpID + 1
                sid = convert(clean_List(entry,'objectSid'))
                grpNm = clean_List(entry,"sAMAccountName")
                if grpNm == None:
                    qualNm = None
                else:
                    qualNm = dNm + "\\" + grpNm
                grpDesc = clean_List(entry,"description")
                createDt1 = clean_List(entry,"whenCreated")
                createDt = datetime.datetime.strptime(createDt1[:-3],"%Y%m%d%H%M%S")
                chgDt1 = clean_List(entry,"whenChanged")
                chgDt = datetime.datetime.strptime(chgDt1[:-3],"%Y%m%d%H%M%S")
                wDt = str(datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
                wUser = "CurtisC"
                sql = '''INSERT INTO dbo.AD_Groups (GrpID,Domain_Name,SID,Group_Name,Fully_Qual_Name,Group_DN,Group_Desc,Create_Date,Last_Changed_Date,
                Write_Date,Last_Record_Update,Write_User)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                '''
                cursor.execute(sql,(grpID,dNm,sid,grpNm,qualNm,dn,grpDesc,createDt,chgDt,wDt,wDt,wUser))
                cursor.execute("Commit;")
                members = []
                if "member" in entry:
                    members = entry["member"]
                    for m in members:
                        tArr = m.split(',')
                        tStr = tArr[0].replace('CN=','')
                        mNm = getUserName(m)
                        sql = "INSERT INTO dbo.AD_Grp2UserMap (GrpID,Domain_Name,Member_Name,Member_DN,Write_Date,Last_Record_Update,Write_User) VALUES (?,?,?,?,?,?,?);"
                        cursor.execute(sql,(grpID,dNm,mNm,m,wDt,wDt,wUser))
                        cursor.execute("Commit;")
            else:
                break
        pctrls = get_pctrls(serverctrls)
        if not pctrls:
            print >> sys.stderr, 'Warning: Server ignores RFC 2696 control.'
            break
        cookie = set_cookie(lc, pctrls, page_size)
        if not cookie:
            break
    l.unbind_s()

    cursor.execute("SELECT * FROM dbo.AD_Grp2UserMap WHERE Member_Name IS NULL;")
    for row in cursor.fetchall():
        gID = row[1]
        mDN = row[4]
        gIDMem = getGrpID(mDN)
        wDt = str(datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
        wUser = "CurtisC"
        if gIDMem != None:
            sql = '''INSERT INTO dbo.AD_Grp2GrpMap (GrpID,GrpID_Member,GrpID_From,Write_Date,Last_Record_Update,Write_User)
            VALUES (?,?,?,?,?,?);'''
            cursor2.execute(sql,(gID,gIDMem,gID,wDt,wDt,wUser))
            cursor2.execute("Commit;")
    keepGoing = True
    while keepGoing == True:
        cursor.execute("SELECT * FROM dbo.AD_Grp2GrpMap WHERE Processed_Flag IS NULL;")
        if len(cursor.fetchall()) > 0:
            for row in cursor.fetchall():
                checkGID = row[2]
                g2gList = checkForGrp2Grp(checkGID)
                if len(g2gList) != 0:
                    keepGoing = True
                    break
                else:
                    keepGoing = False
        else:
            keepGoing = False
        if keepGoing == True:
            cursor.execute("SELECT * FROM dbo.AD_Grp2GrpMap WHERE Processed_Flag IS NULL;")
            for row in cursor.fetchall():
                checkGID = row[2]
                g2gList = checkForGrp2Grp(checkGID)
                if len(g2gList) != 0:
                    for g in g2gList:
                        gidVal = getGrpID(g)
                        if gidVal != None and gidVal != row[1] and checkForMapping(row[1],gidVal,checkGID) != "No":
                            wDt = str(datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
                            wUser = "CurtisC"
                            sql = '''INSERT INTO dbo.AD_Grp2GrpMap (GrpID,GrpID_Member,GrpID_From,Write_Date,Last_Record_Update,Write_User)
                            VALUES (?,?,?,?,?,?);'''
                            cursor2.execute(sql,(row[1],gidVal,checkGID,wDt,wDt,wUser))
                            cursor2.execute("Commit;")
                updateProcessedFlag(row[0])
        elif keepGoing == False:
            cursor.execute("SELECT * FROM dbo.AD_Grp2GrpMap WHERE Processed_Flag IS NULL;")
            for row in cursor.fetchall():
                updateProcessedFlag(row[0])
        cursor.execute("SELECT * FROM dbo.AD_Grp2UserMap WHERE Member_Name IS NULL AND Member_DN LIKE 'CN=S-%';")
        for row in cursor.fetchall():
            tArr = []
            tArr = row[4].split(',')
            tStr = tArr[0].replace('CN=','')
            mNm, mDn = getSIDName(tStr)
            rowId = row[0]
            if mNm != None and mDn != None:
                cursor2.execute("UPDATE dbo.AD_Grp2UserMap SET Member_Name = ?, Member_DN = ? WHERE RowID = ?;",(mNm,mDn,rowId))
                cursor2.execute("Commit;")
            else:
                cursor2.execute("UPDATE dbo.AD_Grp2UserMap SET Member_Name = ? WHERE RowID = ?;",(tStr,rowId))
                cursor2.execute("Commit;")
        cursor.execute("DELETE FROM dbo.AD_Grp2UserMap WHERE Member_Name IS NULL;")
        cursor.execute("Commit;")
cursor2.close()
cursor.close()
cnn.close()

print "---------------------------------Script Successfully Run-----------------------------------"
