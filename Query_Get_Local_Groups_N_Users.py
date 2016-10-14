#Script to produce a csv file of local groups and users for specified Windows servers/computers
#Developed by Chris Curtis - Internal Audit
#Last updated on 3/11/16

import win32net
import win32security
import csv
import sys
import socket
import datetime

svr = ['xxxx']
hdr = ['IP_Address','DNS_Name','Group_Name','Group_Member_Qual','Group_Member_Non_Qual','Account_Desc','Full_Name','Last_Logon_Date','Acct_Exp_Date','User_Account_Control','Acct_Disabled','Pwd_Never_Expires', \
       'Pwd_Not_Req','Pwd_Expired','Pwd_Cannot_Change']

def conv_AD_time(adtime):
    secs = adtime / 10000000
    if secs < 11644473600:
        return adtime
    else:
        conv = secs - 11644473600
        dt = datetime.datetime(2000,1,1,0,0)
        return dt.fromtimestamp(conv)

ofile = open('SOX Local Users_100316.csv','wb')
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer.writerow(hdr)

for s in svr:
    resume = 0
    while True:
        dnsNm, aliasL, ipaddrL = socket.gethostbyaddr(s)
        data, totalG, resume = win32net.NetLocalGroupEnum(s,1,resume)
        for group in data:
            mbrresume = 0
            while True:
                memberdata, totalM, mbrresume = win32net.NetLocalGroupGetMembers(s,group['name'],2,resume)
                for member in memberdata:
                    wlist = []
                    wlist.append(s)
                    wlist.append(dnsNm)
                    wlist.append(group['name'])
                    username, domain, type = win32security.LookupAccountSid(s,member['sid'])
                    wlist.append(member['domainandname'])
                    wlist.append(username)
                    try:
                        info = win32net.NetUserGetInfo(s,username,3)
                        wlist.append(info['comment'])
                        wlist.append(info['full_name'])
                        adtime = info['last_logon']
                        if adtime == 0 or adtime == 9223372036854775807:
                            lastlogon = "Never"
                        else:
                            lastlogon = conv_AD_time(adtime)
                        wlist.append(lastlogon)
                        if info['acct_expires'] == 9223372036854775807 or info['acct_expires'] == 0:
                            acctExp = None
                        elif info['acct_expires'] != None:
                            acctExp = conv_AD_time(int(info['acct_expires']))
                        else:
                            acctExp = None
                        wlist.append(acctExp)
                        uac = info['flags']
                        wlist.append(info['flags'])
                        #account disabled flags
                        if uac == 512 or uac == 522 or uac == 544 or uac == 2080 or uac == 66048 or uac == 66049 or uac == 66113 or uac == 590336 or uac == 8389121 or uac == 16843264:
                            acctDis = 'N'
                        elif uac == 514 or uac == 546 or uac == 66050 or uac == 66115 or uac == 8389123 or uac == 8389187:
                            acctDis = 'Y'
                        else:
                            acctDis = None
                        #pwdNeverExp flags
                        if uac == 512 or uac == 514 or uac == 522 or uac == 546 or uac == 2080 or uac == 8389121 or uac == 8389123:
                            pwdNeverExp = 'N'
                        elif uac == 66048 or uac == 66049 or uac == 66050 or uac == 66115 or uac == 590336 or uac == 66113 or uac == 8389187 or uac == 16843264:
                            pwdNeverExp = 'Y'
                        else:
                            pwdNeverExp = None
                        #pwdNotReq flags
                        if uac == 546 or uac == 2080:
                            pwdNotReq = 'Y'
                        elif (uac == 512 or uac == 514 or uac == 522 or uac == 544 or uac == 66048 or uac == 66049 or uac == 66050 or uac == 66113 or uac == 66115 or uac == 590336
                              or uac == 8389121 or uac == 8389123 or uac == 8389187 or uac == 16843264):
                            pwdNotReq = 'N'
                        else:
                            pwdNotReq = None
                        #pwdExpired flags
                        if uac == 8389123 or uac == 8389121 or uac == 8389187:
                            pwdExp = 'Y'
                        elif (uac == 512 or uac == 514 or uac == 522 or uac == 544 or uac == 546 or uac == 2080 or uac == 66048 or uac == 66049 or uac == 66050 or uac == 66113
                              or uac == 66115 or uac == 590336 or uac == 16843264):
                            pwdExp = 'N'
                        else:
                            pwdExp = None
                        #pwdCannotChange flags
                        if uac == 66113 or uac == 66115 or 8389187:
                            pwdCannotChg = 'Y'
                        elif (uac == 512 or uac == 514 or uac == 522 or uac == 544 or uac == 546 or uac == 2080 or uac == 66048 or uac == 66049 or uac == 66050 or uac == 590336
                              or uac == 8389123 or uac == 8389121 or uac == 16843264):
                            pwdCannotChg = 'N'
                        else:
                            pwdCannotChg = None
                        wlist.append(acctDis)
                        wlist.append(pwdNeverExp)
                        wlist.append(pwdNotReq)
                        wlist.append(pwdExp)
                        wlist.append(pwdCannotChg)
                        writer.writerow(wlist)
                    except:
                        writer.writerow(wlist)
                if mbrresume == 0:
                    break
        if not resume:
            break

ofile.close()

print "----------------------Script Successfully Run-----------------------"
            
