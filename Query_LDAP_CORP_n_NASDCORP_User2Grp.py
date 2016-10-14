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

#Constant variable definitions
LDAP24API = StrictVersion(ldap.__version__) >= StrictVersion('2.4')
domains = ['NASDCORP','NASDROOT','CORP']

try:
    cnn = pypyodbc.connect(r"TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=ny4wnp01-sql1;DATABASE=IA_DB;")
    cursor = cnn.cursor()

    #cursor.execute("DELETE FROM dbo.AD_Users;")
    #cursor.execute("Commit;")
    #cursor.execute("DELETE FROM dbo.AD_User2GrpMap;")
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
    cursor.execute("SELECT MAX(UserID) AS MaxKey FROM dbo.AD_Users;")
    for row in cursor.fetchall():
        maxKey = row[0]
    if maxKey == None:
        uCounter = 0
    else:
        uCounter = maxKey
    return uCounter

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

for d in domains:
    if d == 'NASDCORP':
        baseDN = "DC=corp,DC=root,DC=nasd,DC=com"
    elif d == 'CORP':
        baseDN = "DC=corp,DC=finra,DC=org"
    elif d == 'NASDROOT':
        baseDN = "DC=root,DC=nasd,DC=com"
    searchScope = ldap.SCOPE_SUBTREE
    retrieveAttributes = None
    searchFilter = "(&(objectclass=user)(!(objectclass=computer)))"
    page_size = 10
    dNm = d
      
    try:
        ldap.set_option(ldap.OPT_REFERRALS, 0)
        if d == 'NASDCORP':
            l = ldap.initialize("ldap://ny4wncpdcp003.corp.root.nasd.com:389/")
        elif d == 'CORP':
            l = ldap.initialize("ldap://ny4-corp-dcp1.corp.finra.org:389/")
        elif d == 'NASDROOT':
            l = ldap.initialize("ldap://rkv-cp-dcp1.root.nasd.com:389/")
        l.protocol_version = ldap.VERSION3
        if d == 'NASDCORP':
            username = "CN=Christopher Curtis,OU=Users,OU=Locations,DC=corp,DC=root,DC=nasd,DC=com"
            password = 'Ch0c0&J0c3lyn'
        elif d == 'CORP':
            username = "CN=Christopher Curtis (IA),OU=Privileged_User_Accounts,OU=Operations,DC=corp,DC=finra,DC=org"
            password = 'Go#9terps'
        elif d == 'NASDROOT':
            username = "CN=Christopher Curtis,OU=Users,OU=Locations,DC=corp,DC=root,DC=nasd,DC=com"
            password = 'Ch0c0&J0c3lyn'
        #line below is to use TLS connection for LDAP, which is currently not enabled or allowed
        #l.start_tls_s()
        l.simple_bind_s(username, password)
    except ldap.LDAPError as e:
        print e
        sys.exit()
        
    lc = create_controls(page_size)
    userID = findMaxKey()

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
                userID = userID + 1
                sid = convert(clean_List(entry,'objectSid'))
                uNm = clean_List(entry,"sAMAccountName")
                if uNm == None:
                    qualNm = None
                else:
                    qualNm = dNm + "\\" + uNm
                fullNm = clean_List(entry,"name")
                fNm = clean_List(entry,"givenName")
                lNm = clean_List(entry,"sn")
                loc = clean_List(entry,"l")
                state = clean_List(entry,"st")
                officeLoc = clean_List(entry,"physicalDeliveryOfficeName")
                acctDesc = clean_List(entry,"description")
                createDt1 = clean_List(entry,"whenCreated")
                createDt = datetime.datetime.strptime(createDt1[:-3],"%Y%m%d%H%M%S")
                chgDt1 = clean_List(entry,"whenChanged")
                chgDt = datetime.datetime.strptime(chgDt1[:-3],"%Y%m%d%H%M%S")
                empID = clean_List(entry,"employeeID")
                uac = int(clean_List(entry,"userAccountControl"))
                #account disabled flags
                if uac == 512 or uac == 522 or uac == 544 or uac == 2080 or uac == 66048 or uac == 590336 or uac == 16843264:
                    acctDis = 'N'
                elif uac == 514 or uac == 546 or uac == 66050:
                    acctDis = 'Y'
                else:
                    acctDis = None
                #pwdNeverExp flags
                if uac == 512 or uac == 514 or uac == 522 or uac == 546 or uac == 2080:
                    pwdNeverExp = 'N'
                elif uac == 66048 or uac == 66050 or uac == 590336 or uac == 16843264:
                    pwdNeverExp = 'Y'
                else:
                    pwdNeverExp = None
                #pwdNotReq flags
                if uac == 546 or uac == 2080:
                    pwdNotReq = 'Y'
                elif uac == 512 or uac == 514 or uac == 522 or uac == 544 or uac == 66048 or uac == 66050 or uac == 590336 or uac == 16843264:
                    pwdNotReq = 'N'
                else:
                    pwdNotReq = None
                if clean_List(entry,"lastLogonTimestamp") != None:
                    lastLogon = conv_AD_time(int(clean_List(entry,"lastLogonTimestamp")))
                else:
                    lastLogon = None
                if clean_List(entry,"accountExpires") == "9223372036854775807" or clean_List(entry,"accountExpires") == "0":
                    acctExp = None
                elif clean_List(entry,"accountExpires") != None:
                    acctExp = conv_AD_time(int(clean_List(entry,"accountExpires")))
                else:
                    acctExp = None
                sql = '''INSERT INTO dbo.AD_Users (UserID,SID,EMPLID,Domain_Name,User_Name,User_DN,Fully_Qual_Name,Full_Name,First_Name,Last_Name,City,State,Office_Location,
                Account_Desc,Create_Date,Last_Changed_Date,Last_Logon_Date,Acct_Exp_Date,User_Acct_Control,Acct_Disabled,Pwd_Never_Expires,Pwd_Not_Req,Write_Date,
                Last_Record_Update,Write_User)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);
                '''
                wDt = str(datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
                wUser = "CurtisC"
                cursor.execute(sql,(userID,sid,empID,dNm,uNm,dn,qualNm,fullNm,fNm,lNm,loc,state,officeLoc,acctDesc,createDt,chgDt,lastLogon,acctExp,uac,acctDis,pwdNeverExp,pwdNotReq,wDt,wDt,wUser))
                cursor.execute("Commit;")
                groups = []
                if "memberOf" in entry:
                    groups = entry["memberOf"]
                    for g in groups:
                        tArr = g.split(',')
                        tStr = tArr[0].replace('CN=','')
                        sql = "INSERT INTO dbo.AD_User2GrpMap (UserID,Domain_Name,Group_Name,Group_DN,Write_Date,Last_Record_Update,Write_User) VALUES (?,?,?,?,?,?,?);"
                        cursor.execute(sql,(userID,dNm,tStr,g,wDt,wDt,wUser))
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

cursor.close()
cnn.close()

print "---------------------------------Script Successfully Run-----------------------------------"
