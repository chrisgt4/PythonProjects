#Script to query Active Directory to pull a list of all computers
#Written by Chris Curtis - Internal Audit
#Last Updated on 4/12/16

import ldap
import sys
import datetime
import csv

from ldap.controls import SimplePagedResultsControl
from distutils.version import StrictVersion
from optparse import OptionParser

#Constant variable definitions
LDAP24API = StrictVersion(ldap.__version__) >= StrictVersion('2.4')
domains = ['NASDCORP']
hdr = ['distinguishedName','name','dNSHostName','description','operatingSystem','operatingSystemVersion','whenCreated','whenChanged','lastLogon','lastLogonTimestamp']

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

ofile = open('Computers.csv','wb')
writer = csv.writer(ofile, delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)
writer.writerow(hdr)

for d in domains:
    if d == 'NASDCORP':
        baseDN = "DC=corp,DC=root,DC=nasd,DC=com"
    searchScope = ldap.SCOPE_SUBTREE
    retrieveAttributes = None
    searchFilter = "(objectclass=computer)"
    page_size = 10
    dNm = d
      
    try:
        ldap.set_option(ldap.OPT_REFERRALS, 0)
        if d == 'NASDCORP':
            l = ldap.initialize("ldap://ny4wncpdcp003.corp.root.nasd.com:389/")
        l.protocol_version = ldap.VERSION3
        if d == 'NASDCORP':
            username = 'username' #redacted
             password = 'password' #redacted
        #line below is to use TLS connection for LDAP, which is currently not enabled or allowed
        #l.start_tls_s()
        l.simple_bind_s(username, password)
    except ldap.LDAPError as e:
        print e
        sys.exit()
        
    lc = create_controls(page_size)
    
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
                cList = []
                disNm = clean_List(entry,"distinguishedName")
                cList.append(disNm)
                cNm = clean_List(entry,"name")
                cList.append(cNm)
                dnsNm = clean_List(entry,"dNSHostName")
                cList.append(dnsNm)
                acctDesc = clean_List(entry,"description")
                cList.append(acctDesc)
                oSys = clean_List(entry,"operatingSystem")
                cList.append(oSys)
                oVer = clean_List(entry,"operatingSystemVersion")
                cList.append(oVer)
                createDt1 = clean_List(entry,"whenCreated")
                createDt = datetime.datetime.strptime(createDt1[:-3],"%Y%m%d%H%M%S")
                cList.append(createDt)
                chgDt1 = clean_List(entry,"whenChanged")
                chgDt = datetime.datetime.strptime(chgDt1[:-3],"%Y%m%d%H%M%S")
                cList.append(chgDt)
                lastLog1 = clean_List(entry,"lastLogon")
                lastLog = datetime.datetime.strptime(chgDt1[:-3],"%Y%m%d%H%M%S")
                cList.append(lastLog)
                lastLogTs1 = clean_List(entry,"lastLogonTimestamp")
                lastLogTs = datetime.datetime.strptime(chgDt1[:-3],"%Y%m%d%H%M%S")
                cList.append(lastLogTs)
                writer.writerow(cList)
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

ofile.close()

print "---------------------------------Script Successfully Run-----------------------------------"
