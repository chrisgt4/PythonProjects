
import boto3
import httplib
import json
import sys
import os
import csv
import datetime

os.environ['AWS_ACCESS_KEY_ID'] = 'xxxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxxx'
os.environ['AWS_SESSION_TOKEN'] = 'xxxx'
os.environ['https_proxy'] = 'https://xxxxxx:8080'
os.environ['http_proxy'] = 'http://xxxxxx:8080'

regs = ['us-east-1','us-west-2']

ofile = open('AWS_All_Sec_Grps.csv',"wb")
ofile2 = open('AWS_All_Sec_Grps_InRules.csv',"wb")
ofile3 = open('AWS_All_Sec_Grps_OutRules.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer2 = csv.writer(ofile2, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
writer3 = csv.writer(ofile3, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['GroupID','OwnerID','Region','VPCId','GroupName','Description','Tags']
header2 = ['GroupID','GroupName','IpProtocol','FromPort','ToPort','UserIdGroupPairs','IpRanges']
header3 = ['GroupID','GroupName','IpProtocol','FromPort','ToPort','UserIdGroupPairs','IpRanges']
writer.writerow(header)
writer2.writerow(header2)
writer3.writerow(header3)

#ec2 = boto3.resource('ec2',reg)

for r in regs:
    client = boto3.client('ec2',r)
    resp = client.describe_security_groups()
    secgroups = []
    secgroups = resp['SecurityGroups']
    for s in secgroups:
        vals = []
        gid = s['GroupId']
        oid = s['OwnerId']
        vpcid = s['VpcId']
        gnm = s['GroupName']
        desc = s['Description']
        inrules = s['IpPermissions']
        if len(inrules) != 0:
            for i in inrules:
                inrulevals = []
                #print i
                inrulevals.append(gid)
                inrulevals.append(gnm)
                inrulevals.append(i['IpProtocol'])
                if 'FromPort' in i:
                    inrulevals.append(i['FromPort'])
                else:
                    inrulevals.append('None')
                if 'ToPort' in i:
                    inrulevals.append(i['ToPort'])
                else:
                    inrulevals.append('None')
                if 'UserIdGroupPairs' in i:
                    if len(i['UserIdGroupPairs']) > 0:
                        #print i['UserIdGroupPairs']
                        ugpairs = []
                        ugpairs = i['UserIdGroupPairs']
                        ugpairvals = []
                        for ug in ugpairs:
                            ugpairvals.append(ug['UserId'] + ':' + ug['GroupId'])
                    else:
                        ugpairvals = '[None]'
                else:
                    ugpairvals = '[None]'
                inrulevals.append(ugpairvals)

                if 'IpRanges' in i:
                    if len(i['IpRanges']) > 0:
                        ipr = []
                        ipr = i['IpRanges']
                        ipvals = []
                        for ip in ipr:
                            ipvals.append(ip['CidrIp'])
                    else:
                        ipvals = '[None]'
                else:
                    ipvals = '[None]'
                inrulevals.append(ipvals)
                writer2.writerow(inrulevals)
        outrules = s['IpPermissionsEgress']
        if len(outrules) != 0:
            for i in outrules:
                outrulevals = []
                #print i
                outrulevals.append(gid)
                outrulevals.append(gnm)
                outrulevals.append(i['IpProtocol'])
                if 'FromPort' in i:
                    outrulevals.append(i['FromPort'])
                else:
                    outrulevals.append('None')
                if 'ToPort' in i:
                    outrulevals.append(i['ToPort'])
                else:
                    outrulevals.append('None')
                if 'UserIdGroupPairs' in i:
                    if len(i['UserIdGroupPairs']) > 0:
                        #print i['UserIdGroupPairs']
                        ugpairs = []
                        ugpairs = i['UserIdGroupPairs']
                        ugpairvals = []
                        for ug in ugpairs:
                            ugpairvals.append(ug['UserId'] + ':' + ug['GroupId'])
                    else:
                        ugpairvals = '[None]'
                else:
                    ugpairvals = '[None]'
                outrulevals.append(ugpairvals)
                if 'IpRanges' in i:
                    if len(i['IpRanges']) > 0:
                        ipr = []
                        ipr = i['IpRanges']
                        ipvals = []
                        for ip in ipr:
                            ipvals.append(ip['CidrIp'])
                    else:
                        ipvals = '[None]'
                else:
                    ipvals = '[None]'
                outrulevals.append(ipvals)
                writer3.writerow(inrulevals)
        tags = []
        tvals = []
        if 'Tags' in s:
            tags = s['Tags']
            for t in tags:
                tvals.append(t['Key'] + '||' + t['Value'])
        else:
            tvals = '[None]'
        vals.append(gid)
        if oid == '510199193688':
            vals.append('finra-awsprod')
        else:
            vals.append(oid)
        vals.append(r)
        vals.append(vpcid)
        vals.append(gnm)
        vals.append(desc)
        vals.append(tvals)
        writer.writerow(vals)

ofile.close()
ofile2.close()
ofile3.close()

print "================Script Successful================="
