#GetWinIPMapping script - connect to SCCM reporting server and grab the IP Address mapping information for Windows servers
# - written by Chris Curtis - Internal Audit
# - last revision on 8/26/2015
# - version 1.0 of the module
# - version 1.0 release notes
#     - initial development of script

import os
import sys
import pyodbc
import csv
import shutil

cnn = pyodbc.connect(r"trusted_connection=Yes;driver={SQL Server};server=xxxx;database=xxxx")
cursor = cnn.cursor()

ofile = open('IP_Mapping_Windows.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['IP_Addr','NetBIOS','DNSName','OS','Domain','AD_Site','Description']
writer.writerow(header)

sql = """
    SELECT 
	DISTINCT ip.IP_Addresses0 AS IP_Addr
	, s.Netbios_Name0 AS NetBIOS
	, s.Name0 AS Name
	, s.Operating_System_Name_and0 AS OS
	, s.Resource_Domain_OR_Workgr0 AS Domain
	, s.AD_Site_Name0 AS AD_Site
	, s.description0 AS Description
    FROM dbo.System_DISC s
    LEFT OUTER JOIN dbo.System_IP_Address_ARR ip
    ON ip.ItemKey = s.ItemKey
    WHERE s.Operating_System_Name_and0 NOT LIKE '%workstation%'
    AND ip.IP_Addresses0 IS NOT NULL
    ORDER BY s.Operating_System_Name_and0;
"""

cursor.execute(sql)
rows = cursor.fetchall()

for r in rows:
    R = []
    R.append(r.IP_Addr)
    R.append(r.NetBIOS)
    R.append(r.Name)
    R.append(r.OS)
    R.append(r.Domain)
    R.append(r.AD_Site)
    R.append(r.Description)
    writer.writerow(R)

ofile.close()
shutil.move(ofile.name,"\\\\xxxxx" + ofile.name)

cursor.close()
cnn.close()

print "****************************************Script Successfully Run*********************************************************"
