#Script to load Qualys raw knowledge base data from XML into a SQL Server database
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 6/23/2015
#commented out lines are for writing to a csv instead of to a database, can be recommented back in

import sys
sys.path.append('./Lib/site-packages/')
sys.path.append('F:/Python Projects/')
import os
import re
import csv
import pypyodbc

#define the directory to search for the file
strFlLoc = r"C:\Users\CurtisC\Desktop\kb-2015052000.xml"
flList = os.listdir(strFlLoc)

#find the xml file using regex to parse
for f in flList:
    flNmPat = re.compile(".{1,100}xml")
    if flNmPat.match(f) != None:
        flNm = strFlLoc + "\\" + f

#setup database connection to local SQL Server
cnn = pypyodbc.connect("TRUSTED_CONNECTION=Yes;DRIVER={SQL Server};SERVER=xxxxx;DATABASE=xxxxx;")
cursor = cnn.cursor()

#clean the database table
sql = "DELETE FROM dbo.Qualys_KB_052015;"
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
isVuln = False
isBothDiag = False
isDiag = False
isBothConseq = False
isConseq = False
isBothSol = False
isSol = False
isSwList = False
isExploit = False
isVendRef = False
isCVEId = False
isBugTraq = False
counterk = 0

#parse that Qualys xml like a bad habit
for l in lines:
    #code below is to insert values to the Qualys_KB table
    if l == "<VULN>":
        isVuln = True
    elif l == "</VULN>":
        isVuln = False
    if isVuln == True and l != "</VULN>":
        if l.endswith("</QID>", 0, len(l)) == True:
            K = []
            qid = "none"
            qtype = "none"
            severity = "none"
            title = "none"
            category = "none"
            lastmoddt = "none"
            publishdt = "none"
            diagnosis = "none"
            consequence = "none"
            solution = "none"
            swproduct = "none"
            swvendor = "none"
            exploitrefs = "none"
            exploitdescs = "none"
            exploiturls = "none"
            vendorrefs = "none"
            vendorurls = "none"
            cveids = "none"
            cveurls = "none"
            bugtraqids = "none"
            bugtraqurls = "none"
            cleanl = l.replace("<QID>","")
            cleanl2 = cleanl.replace("</QID>","")
            qid = "qid_" + str(cleanl2)
            K.append(qid)
        elif l.endswith("</VULN_TYPE>", 0, len(l)) == True:
            cleanl = l.replace("<VULN_TYPE>","")
            cleanl2 = cleanl.replace("</VULN_TYPE>","")
            qtype = cleanl2
        elif l.endswith("</SEVERITY>", 0, len(l)) == True:
            cleanl = l.replace("<SEVERITY>","")
            cleanl2 = cleanl.replace("</SEVERITY>","")
            severity = cleanl2
        elif l.endswith("</TITLE>", 0, len(l)) == True:
            cleanl = l.replace("<TITLE><![CDATA[","")
            cleanl2 = cleanl.replace("]]></TITLE>","")
            title = cleanl2
        elif l.endswith("</CATEGORY>", 0, len(l)) == True:
            cleanl = l.replace("<CATEGORY>","")
            cleanl2 = cleanl.replace("</CATEGORY>","")
            category = cleanl2
        elif l.endswith("</LAST_SERVICE_MODIFICATION_DATETIME>", 0, len(l)) == True:
            cleanl = l.replace("<LAST_SERVICE_MODIFICATION_DATETIME>","")
            cleanl2 = cleanl.replace("</LAST_SERVICE_MODIFICATION_DATETIME>","")
            lastmoddt = cleanl2
        elif l.endswith("</PUBLISHED_DATETIME>", 0, len(l)) == True:
            cleanl = l.replace("<PUBLISHED_DATETIME>","")
            cleanl2 = cleanl.replace("</PUBLISHED_DATETIME>","")
            publishdt = cleanl2
        if l.startswith("<DIAGNOSIS>", 0, len(l)) == True and l.endswith("</DIAGNOSIS>", 0, len(l)) == True:
            isBothDiag = True
        else:
            isBothDiag = False
        if l.startswith("<DIAGNOSIS>", 0, len(l)) == True and l.endswith("</DIAGNOSIS>", 0, len(l)) != True:
            isDiag = True
        elif l.endswith("</DIAGNOSIS>", 0, len(l)) == True and l.startswith("<DIAGNOSIS>", 0, len(l)) != True:
            isDiag = False
        if isBothDiag == True:
            cleanl = l.replace("<DIAGNOSIS><![CDATA[","")
            cleanl2 = cleanl.replace("]]></DIAGNOSIS>","")
            diagnosis = cleanl2
        elif isDiag == True and l.endswith("</DIAGNOSIS>", 0, len(l)) != True:
            cleanl = l.replace("<DIAGNOSIS><![CDATA[","")
            cleanl2 = cleanl.replace("]]></DIAGNOSIS>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if diagnosis == "none":
                diagnosis = cleanl6
            else:
                diagnosis = diagnosis + " " + cleanl6
        elif isDiag == False and l.endswith("</DIAGNOSIS>", 0, len(l)) == True:
            cleanl = l.replace("<DIAGNOSIS><![CDATA[","")
            cleanl2 = cleanl.replace("]]></DIAGNOSIS>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if diagnosis == "none":
                diagnosis = cleanl6
            else:
                diagnosis = diagnosis + " " + cleanl6
        if l.startswith("<CONSEQUENCE>", 0, len(l)) == True and l.endswith("</CONSEQUENCE>", 0, len(l)) == True:
            isBothConseq = True
        else:
            isBothConseq = False
        if l.startswith("<CONSEQUENCE>", 0, len(l)) == True and l.endswith("</CONSEQUENCE>", 0, len(l)) != True:
            isConseq = True
        elif l.endswith("</CONSEQUENCE>", 0, len(l)) == True and l.startswith("<CONSEQUENCE>", 0, len(l)) != True:
            isConseq = False
        if isBothConseq == True:
            cleanl = l.replace("<CONSEQUENCE><![CDATA[","")
            cleanl2 = cleanl.replace("]]></CONSEQUENCE>","")
            consequence = cleanl2
        elif isConseq == True and l.endswith("</CONSEQUENCE>", 0, len(l)) != True:
            cleanl = l.replace("<CONSEQUENCE><![CDATA[","")
            cleanl2 = cleanl.replace("]]></CONSEQUENCE>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if consequence == "none":
                consequence = cleanl6
            else:
                consequence = consequence + " " + cleanl6
        elif isConseq == False and l.endswith("</CONSEQUENCE>", 0, len(l)) == True:
            cleanl = l.replace("<CONSEQUENCE><![CDATA[","")
            cleanl2 = cleanl.replace("]]></CONSEQUENCE>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if consequence == "none":
                consequence = cleanl6
            else:
                consequence = consequence + " " + cleanl6
        if l.startswith("<SOLUTION>", 0, len(l)) == True and l.endswith("</SOLUTION>", 0, len(l)) == True:
            isBothSol = True
        else:
            isBothSol = False
        if l.startswith("<SOLUTION>", 0, len(l)) == True and l.endswith("</SOLUTION>", 0, len(l)) != True:
            isSol = True
        elif l.endswith("</SOLUTION>", 0, len(l)) == True and l.startswith("<SOLUTION>", 0, len(l)) != True:
            isSol = False
        if isBothSol == True:
            cleanl = l.replace("<SOLUTION><![CDATA[","")
            cleanl2 = cleanl.replace("]]></SOLUTION>","")
            solution = cleanl2
        elif isSol == True and l.endswith("</SOLUTION>", 0, len(l)) != True:
            cleanl = l.replace("<SOLUTION><![CDATA[","")
            cleanl2 = cleanl.replace("]]></SOLUTION>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if solution == "none":
                solution = cleanl6
            else:
                solution = solution + " " + cleanl6
        elif isSol == False and l.endswith("</SOLUTION>", 0, len(l)) == True:
            cleanl = l.replace("<SOLUTION><![CDATA[","")
            cleanl2 = cleanl.replace("]]></SOLUTION>","")
            cleanl3 = cleanl2.replace("<P>","")
            cleanl4 = cleanl3.replace("<BR>","")
            cleanl5 = cleanl4.replace("<B>","")
            cleanl6 = cleanl5.replace("</B>","")
            if solution == "none":
                solution = cleanl6
            else:
                solution = solution + " " + cleanl6
        if l == "<SOFTWARE_LIST>":
            isSwList = True
        elif l == "</SOFTWARE_LIST>":
            isSwList = False
        if isSwList == True and l != "</SOFTWARE_LIST>":
            if l.endswith("</PRODUCT>", 0, len(l)) == True:
                cleanl = l.replace("<PRODUCT><![CDATA[","")
                cleanl2 = cleanl.replace("]]></PRODUCT>","")
                if swproduct == "none":
                    swproduct = cleanl2
                else:
                    swproduct = swproduct + " | " + cleanl2
            if l.endswith("</VENDOR>", 0, len(l)) == True:
                cleanl = l.replace("<VENDOR><![CDATA[","")
                cleanl2 = cleanl.replace("]]></VENDOR>","")
                if swvendor == "none":
                    swvendor = cleanl2
                else:
                    swvendor = swvendor + " | " + cleanl2
        if l == "<EXPLT_LIST>":
            isExploit = True
        elif l == "</EXPLT_LIST>":
            isExploit = False
        if isExploit == True and l != "</EXPLT_LIST>":
            if l.endswith("</REF>", 0, len(l)) == True:
                cleanl = l.replace("<REF><![CDATA[","")
                cleanl2 = cleanl.replace("]]></REF>","")
                if exploitrefs == "none":
                    exploitrefs = cleanl2
                else:
                    exploitrefs = exploitrefs + " | " + cleanl2
            if l.endswith("</DESC>", 0, len(l)) == True:
                cleanl = l.replace("<DESC><![CDATA[","")
                cleanl2 = cleanl.replace("]]></DESC>","")
                if exploitdescs == "none":
                    exploitdescs = cleanl2
                else:
                    exploitdescs = exploitdescs + " | " + cleanl2
            if l.endswith("</LINK>", 0, len(l)) == True:
                cleanl = l.replace("<LINK><![CDATA[","")
                cleanl2 = cleanl.replace("]]></LINK>","")
                if exploiturls == "none":
                    exploiturls = cleanl2
                else:
                    exploiturls = exploiturls + " | " + cleanl2
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
                    vendorrefs = vendorrefs + " | " + cleanl2
            if l.endswith("</URL>", 0, len(l)) == True:
                cleanl = l.replace("<URL><![CDATA[","")
                cleanl2 = cleanl.replace("]]></URL>","")
                if vendorurls == "none":
                    vendorurls = cleanl2
                else:
                    vendorurls = vendorurls + " | " + cleanl2
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
                    cveids = cveids + " | " + cleanl2
            if l.endswith("</URL>", 0, len(l)) == True:
                cleanl = l.replace("<URL><![CDATA[","")
                cleanl2 = cleanl.replace("]]></URL>","")
                if cveurls == "none":
                    cveurls = cleanl2
                else:
                    cveurls = cveurls + " | " + cleanl2
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
                    bugtraqids = bugtraqids + " | " + cleanl2
            if l.endswith("</URL>", 0, len(l)) == True:
                cleanl = l.replace("<URL><![CDATA[","")
                cleanl2 = cleanl.replace("]]></URL>","")
                if bugtraqurls == "none":
                    bugtraqurls = cleanl2
                else:
                    bugtraqurls = bugtraqurls + " | " + cleanl2
    if isVuln == False and l == "</VULN>":
        counterk = counterk + 1
        if qtype == "none":
            K.append("blank")
        else:
            K.append(qtype)
        if severity == "none":
            K.append("blank")
        else:
            K.append(severity)
        if title == "none":
            K.append("blank")
        else:
            K.append(title)
        if category == "none":
            K.append("blank")
        else:
            K.append(category)
        if lastmoddt == "none":
            K.append("blank")
        else:
            K.append(lastmoddt)
        if publishdt == "none":
            K.append("blank")
        else:
            K.append(publishdt)
        if diagnosis == "none":
            K.append("blank")
        else:
            K.append(diagnosis)
        if consequence == "none":
            K.append("blank")
        else:
            K.append(consequence)
        if solution == "none":
            K.append("blank")
        else:
            K.append(solution)
        if swproduct == "none":
            K.append("blank")
        else:
            K.append(swproduct)
        if swvendor == "none":
            K.append("blank")
        else:
            K.append(swvendor)
        if exploitrefs == "none":
            K.append("blank")
        else:
            K.append(exploitrefs)
        if exploitdescs == "none":
            K.append("blank")
        else:
            K.append(exploitdescs)
        if exploiturls == "none":
            K.append("blank")
        else:
            K.append(exploiturls)
        if vendorrefs == "none":
            K.append("blank")
        else:
            K.append(vendorrefs)
        if vendorurls == "none":
            K.append("blank")
        else:
            K.append(vendorurls)
        if cveids == "none":
            K.append("blank")
        else:
            K.append(cveids)
        if cveurls == "none":
            K.append("blank")
        else:
            K.append(cveurls)
        if bugtraqids == "none":
            K.append("blank")
        else:
            K.append(bugtraqids)
        if bugtraqurls == "none":
            K.append("blank")
        else:
            K.append(bugtraqurls)
        cleanVals = "'" + "','".join(vals.replace("'","''") for vals in K) + "'"
        sql = "INSERT INTO dbo.Qualys_KB_052015 (QID,QType,Severity,Title,Category,LastModDt,PublishDt,Diagnosis,Consequence,Solution,SWProduct,SWVendor,ExploitRefs,ExploitDescs,ExploitUrls,"
        sql = sql + "VendorRefs,VendorUrls,CVEIds,CVEUrls,BugtraqIds,BugtraqUrls) "
        sql = sql + "VALUES (" + cleanVals + ");"
        cursor.execute(sql)
        cursor.execute("Commit;")    
#close the file
#f.close()
#ofile.close()

#close the database connection
cursor.close()
cnn.close()

print "Total Knowledge Base: " + str(counterk)
print "*****Script successfully complete*****"
