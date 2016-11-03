#Script to load Qualys raw data from XML into a SQL Server database
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 6/5/2015
#commented out lines are for writing to a csv instead of to a database, can be recommented back in

import sys
sys.path.append('./Lib/site-packages/')
sys.path.append('F:/Python Projects/')
import os
import re
import csv
import pypyodbc

#define the directory to search for the file
strFlLoc = r"C:\Users\CurtisC\Desktop\May"
flList = os.listdir(strFlLoc)

#find the xml file using regex to parse
for f in flList:
    flNmPat = re.compile(".{1,100}xml")
    if flNmPat.match(f) != None:
        flNm = strFlLoc + "\\" + f

#setup database connection to local SQL Server
cnn = pypyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxx;DATABASE=xxx;")
cursor = cnn.cursor()

#clean the database table
sql = "DELETE FROM dbo.Qualys_Hosts_052015;"
cursor.execute(sql)
cursor.execute("Commit;")
sql = "DELETE FROM dbo.Qualys_Vulns_052015;"
cursor.execute(sql)
cursor.execute("Commit;")
sql = "DELETE FROM dbo.Qualys_Glossary_052015;"
cursor.execute(sql)
cursor.execute("Commit;")

lines = []

#uncomment if writing to csv
#ofile = open('Qualys_Data.csv',"wb")
#writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
#header = ['IP_Addr','DNS_Name','Operating_System','Netbios']
#writer.writerow(header)

#parse the xml file into a list object
with open(flNm, 'r') as f:
    for line in f:
        cleanline = line.strip()
        lines.append(cleanline)

print "finished loading xml to list"
isHost = False
isVuln = False
isGlossary = False
isVulnDet = False
isVendRef = False
isCVEId = False
isBugTraq = False
isAssetGrp = False
isResult = False
qid = "none"
counter = 0
counterv = 0
counterg = 0

#parse that Qualys xml like a bad habit
for l in lines:
    #code below is to insert values to the Qualys_Hosts table
    if l == "<HOST>":
        isHost = True
    elif l == "</HOST>":
        isHost = False
    if isHost == True and l != "</HOST>":
        if l.endswith("</IP>", 0, len(l)) == True:
            L = []
            dns = "none"
            dnsclean = "none"
            osys = "none"
            nbios = "none"
            assetgrps = "none"
            cleanl = l.replace("<IP>","")
            cleanl2 = cleanl.replace("</IP>","")
            L.append(cleanl2)
            ipaddr = cleanl2
        elif l.endswith("</DNS>", 0, len(l)) == True:
            cleanl = l.replace("<DNS><![CDATA[","")
            cleanl2 = cleanl.replace("]]></DNS>","")
            dns = cleanl2
            if dns.find(".", 0, len(dns)) != -1:
                dnslist = dns.split(".")
                dnsclean = dnslist[0]
            elif dns == "none" and nbios != "none":
                dnsclean = nbios
            else:
                dnsclean = dns
        elif l.endswith("</OPERATING_SYSTEM>", 0, len(l)) == True:
            cleanl = l.replace("<OPERATING_SYSTEM><![CDATA[","")
            cleanl2 = cleanl.replace("]]></OPERATING_SYSTEM>","")
            osys = cleanl2
        elif l.endswith("</NETBIOS>", 0, len(l)) == True:
            cleanl = l.replace("<NETBIOS><![CDATA[","")
            cleanl2 = cleanl.replace("]]></NETBIOS>","")
            nbios = cleanl2
        if l == "<ASSET_GROUPS>":
                isAssetGrp = True
        elif l == "</ASSET_GROUPS>":
            isAssetGrp = False
        if isAssetGrp == True and l != "</ASSET_GROUPS>":
            if l.endswith("</ASSET_GROUP_TITLE>", 0, len(l)) == True:
                cleanl = l.replace("<ASSET_GROUP_TITLE><![CDATA[","")
                cleanl2 = cleanl.replace("]]></ASSET_GROUP_TITLE>","")
                if assetgrps == "none":
                    assetgrps = cleanl2
                else:
                    assetgrps = assetgrps + "|" + cleanl2
    if isHost == False and l == "</HOST>":
        counter = counter + 1
        if dns == "none":
            L.append("blank")
        else:
            L.append(dns)
        if dnsclean == "none":
            L.append("blank")
        else:
            L.append(dnsclean)
        if osys == "none":
            L.append("blank")
        else:
            L.append(osys)
        if nbios == "none":
            L.append("blank")
        else:
            L.append(nbios)
        if assetgrps == "none":
            L.append("blank")
        else:
            L.append(assetgrps)
        cleanVals = "'" + "','".join(vals.replace("'","''") for vals in L) + "'"
        sql = "INSERT INTO dbo.Qualys_Hosts_052015 (IP_Addr,DNS_Name,DNS_Clean,Operating_Sys,NETBIOS,Asset_Groups) "
        sql = sql + "VALUES (" + cleanVals + ");"
        cursor.execute(sql)
        cursor.execute("Commit;")
        #writer.writerow(L)
    #code below is to insert values to the Qualys_Vulns table
    if l == "<VULN_INFO>":
        isVuln = True
    elif l == "</VULN_INFO>":
        isVuln = False
        counterv = counterv + 1
    if isVuln == True and l != "</VULN_INFO>":
        if l.endswith("</QID>", 0, len(l)) == True:
            T = []
            qtype = "none"
            cvss = "none"
            status = "none"
            firstdetect = "none"
            lastdetect = "none"
            scanstart = "none"
            scanend = "none"
            timesfound = "none"
            ssl = "none"
            port = "none"
            service = "none"
            protocol = "none"
            cleanlist = l.split('"')
            cleanl = cleanlist[1]
            T.append(ipaddr)
            qid = cleanl
            T.append(qid)
        elif l.endswith("</TYPE>", 0, len(l)) == True:
            cleanl = l.replace("<TYPE>","")
            cleanl2 = cleanl.replace("</TYPE>","")
            qtype = cleanl2
        elif l.endswith("</CVSS_FINAL>", 0, len(l)) == True:
            cleanl = l.replace("<CVSS_FINAL>","")
            cleanl2 = cleanl.replace("</CVSS_FINAL>","")
            cvss = cleanl2
        elif l.endswith("</VULN_STATUS>", 0, len(l)) == True:
            cleanl = l.replace("<VULN_STATUS>","")
            cleanl2 = cleanl.replace("</VULN_STATUS>","")
            status = cleanl2
        elif l.endswith("</FIRST_FOUND>", 0, len(l)) == True:
            cleanl = l.replace("<FIRST_FOUND>","")
            cleanl2 = cleanl.replace("</FIRST_FOUND>","")
            firstdetect = cleanl2
        elif l.endswith("</LAST_FOUND>", 0, len(l)) == True:
            cleanl = l.replace("<LAST_FOUND>","")
            cleanl2 = cleanl.replace("</LAST_FOUND>","")
            lastdetect = cleanl2
        elif l.endswith("</TIMES_FOUND>", 0, len(l)) == True:
            cleanl = l.replace("<TIMES_FOUND>","")
            cleanl2 = cleanl.replace("</TIMES_FOUND>","")
            timesfound = cleanl2
        elif l.endswith("</SSL>", 0, len(l)) == True:
            cleanl = l.replace("<SSL>","")
            cleanl2 = cleanl.replace("</SSL>","")
            ssl = cleanl2
        elif l.endswith("</PORT>", 0, len(l)) == True:
            cleanl = l.replace("<PORT>","")
            cleanl2 = cleanl.replace("</PORT>","")
            port = cleanl2
        elif l.endswith("</SERVICE>", 0, len(l)) == True:
            cleanl = l.replace("<SERVICE>","")
            cleanl2 = cleanl.replace("</SERVICE>","")
            service = cleanl2
        elif l.endswith("</PROTOCOL>", 0, len(l)) == True:
            cleanl = l.replace("<PROTOCOL>","")
            cleanl2 = cleanl.replace("</PROTOCOL>","")
            protocol = cleanl2
    if qid == "qid_45038":
        if l.startswith("<RESULT>", 0, len(l)) == True:
            isResult = True
        elif l.endswith("</RESULT>", 0, len(l)) == True:
            isResult = False
    if isResult == True or l.endswith("</RESULT>", 0, len(l)) == True:
        if l.startswith("Start time: ", 0, len(l)) == True:
            cleanl = l.replace("Start time: ","")
            cleanl2 = cleanl.replace(" GMT","")
            scanstart = cleanl2
        elif l.startswith("End time: ", 0, len(l)) == True:
            cleanl = l.replace("End time: ","")
            cleanl2 = cleanl.replace(" GMT]]></RESULT>","")
            scanend = cleanl2
    if isVuln == False and l == "</VULN_INFO>":
        if qtype == "none":
            T.append("blank")
        else:
            T.append(qtype)
        if cvss == "none":
            T.append("blank")
        else:
            T.append(cvss)
        if status == "none":
            T.append("blank")
        else:
            T.append(status)
        if firstdetect == "none":
            T.append("blank")
        else:
            T.append(firstdetect)
        if lastdetect == "none":
            T.append("blank")
        else:
            T.append(lastdetect)
        if scanstart == "none":
            T.append("blank")
        else:
            T.append(scanstart)
        if scanend == "none":
            T.append("blank")
        else:
            T.append(scanend)
        if timesfound == "none":
            T.append("blank")
        else:
            T.append(timesfound)
        if ssl == "none":
            T.append("blank")
        else:
            T.append(ssl)
        if port == "none":
            T.append("blank")
        else:
            T.append(port)
        if service == "none":
            T.append("blank")
        else:
            T.append(service)
        if protocol == "none":
            T.append("blank")
        else:
            T.append(protocol)
        cleanVals = "'" + "','".join(vals.replace("'","''") for vals in T) + "'"
        sql = "INSERT INTO dbo.Qualys_Vulns_052015 (IP_Addr,QID,QType,CVSS_Score,QStatus,FirstDetect,LastDetect,ScanStart,ScanEnd,TimesFound,SSL,Port,Service,Protocol) "
        sql = sql + "VALUES (" + cleanVals + ");"
        cursor.execute(sql)
        cursor.execute("Commit;")
    #code below is to insert values to the Qualys_Glossary table
    if l == "<GLOSSARY>":
        isGlossary = True
    elif l == "</GLOSSARY>":
        isGlossary = False
    if isGlossary == True and l != "</GLOSSARY>":
        if l.startswith("<VULN_DETAILS id=", 0, len(l)) == True:
            isVulnDet = True
        elif l == "</VULN_DETAILS>":
            isVulnDet = False
            counterg = counterg + 1
        if isGlossary == True and isVulnDet == True and l != "</VULN_DETAILS>":
            if l.endswith("</QID>", 0, len(l)) == True:
                D = []
                title = "none"
                severity = "none"
                category = "none"
                publishdt = "none"
                vendorrefs = "none"
                vendorurls = "none"
                cveids = "none"
                cveurls = "none"
                bugtraqids = "none"
                bugtraqurls = "none"
                cleanlist = l.split('"')
                cleanl = cleanlist[1]
                D.append(cleanl)
            elif l.endswith("</TITLE>", 0, len(l)) == True:
                cleanl = l.replace("<TITLE><![CDATA[","")
                cleanl2 = cleanl.replace("]]></TITLE>","")
                title = cleanl2
            elif l.endswith("</SEVERITY>", 0, len(l)) == True:
                cleanl = l.replace("<SEVERITY>","")
                cleanl2 = cleanl.replace("</SEVERITY>","")
                severity = cleanl2
            elif l.endswith("</CATEGORY>", 0, len(l)) == True:
                cleanl = l.replace("<CATEGORY>","")
                cleanl2 = cleanl.replace("</CATEGORY>","")
                category = cleanl2
            elif l.endswith("</LAST_UPDATE>", 0, len(l)) == True:
                cleanl = l.replace("<LAST_UPDATE>","")
                cleanl2 = cleanl.replace("</LAST_UPDATE>","")
                publishdt = cleanl2
            if l == "<VENDOR_REFERENCE_LIST>":
                isVendRef = True
            elif l == "</VENDOR_REFERENCE_LIST>":
                isVendRef = False
            if isVendRef == True and l != "</VENDOR_REFERENCE_LIST>":
                if l.endswith("</ID>", 0, len(l)) == True:
                    cleanl = l.replace("<ID><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></ID>","")
                    if vendorrefs == "none":
                        vendorrefs = cleanl2
                    else:
                        vendorrefs = vendorrefs + "|" + cleanl2
                if l.endswith("</URL>", 0, len(l)) == True:
                    cleanl = l.replace("<URL><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></URL>","")
                    if vendorurls == "none":
                        vendorurls = cleanl2
                    else:
                        vendorurls = vendorurls + "|" + cleanl2
            if l == "<CVE_ID_LIST>":
                isCVEId = True
            elif l == "</CVE_ID_LIST>":
                isCVEId = False
            if isCVEId == True and l != "</CVE_ID_LIST>":
                if l.endswith("</ID>", 0, len(l)) == True:
                    cleanl = l.replace("<ID><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></ID>","")
                    if cveids == "none":
                        cveids = cleanl2
                    else:
                        cveids = cveids + "|" + cleanl2
                if l.endswith("</URL>", 0, len(l)) == True:
                    cleanl = l.replace("<URL><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></URL>","")
                    if cveurls == "none":
                        cveurls = cleanl2
                    else:
                        cveurls = cveurls + "|" + cleanl2
            if l == "<BUGTRAQ_ID_LIST>":
                isBugTraq = True
            elif l == "</BUGTRAQ_ID>":
                isBugTraq = False
            if isBugTraq == True and l != "</BUGTRAQ_ID>":
                if l.endswith("</ID>", 0, len(l)) == True:
                    cleanl = l.replace("<ID><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></ID>","")
                    if bugtraqids == "none":
                        bugtraqids = cleanl2
                    else:
                        bugtraqids = bugtraqids + "|" + cleanl2
                if l.endswith("</URL>", 0, len(l)) == True:
                    cleanl = l.replace("<URL><![CDATA[","")
                    cleanl2 = cleanl.replace("]]></URL>","")
                    if bugtraqurls == "none":
                        bugtraqurls = cleanl2
                    else:
                        bugtraqurls = bugtraqurls + "|" + cleanl2
        if isGlossary == True and isVulnDet == False and l == "</VULN_DETAILS>":
            counterg = counterg + 1
            if title == "none":
                D.append("blank")
            else:
                D.append(title)
            if severity == "none":
                D.append("blank")
            else:
                D.append(severity)
            if category == "none":
                D.append("blank")
            else:
                D.append(category)
            if publishdt == "none":
                D.append("blank")
            else:
                D.append(publishdt)
            if vendorrefs == "none":
                D.append("blank")
            else:
                D.append(vendorrefs)
            if vendorurls == "none":
                D.append("blank")
            else:
                D.append(vendorurls)
            if cveids == "none":
                D.append("blank")
            else:
                D.append(cveids)
            if cveurls == "none":
                D.append("blank")
            else:
                D.append(cveurls)
            if bugtraqids == "none":
                D.append("blank")
            else:
                D.append(bugtraqids)
            if bugtraqurls == "none":
                D.append("blank")
            else:
                D.append(bugtraqurls)
            cleanVals = "'" + "','".join(vals.replace("'","''") for vals in D) + "'"
            sql = "INSERT INTO dbo.Qualys_Glossary_052015 (QID,Title,Severity,Category,PublishDt,VendorRefs,VendorUrls,CVEIds,CVEUrls,BugtraqIds,BugtraqUrls) "
            sql = sql + "VALUES (" + cleanVals + ");"
            cursor.execute(sql)
            cursor.execute("Commit;")    
#close the file
#f.close()
#ofile.close()

#close the database connection
cursor.close()
cnn.close()

print "Total Hosts: " + str(counter)
print "Total Vulnerabilities: " + str(counterv)
print "Total Glossary: " + str(counterg)
print "*****Script successfully complete*****"
