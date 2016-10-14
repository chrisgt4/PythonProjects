
import boto3
import httplib
import json
import sys
import os
import csv
from datetime import datetime
import time

os.environ['AWS_ACCESS_KEY_ID'] = 'xxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxx'
os.environ['AWS_SESSION_TOKEN'] = 'x'
os.environ['https_proxy'] = 'xxxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

ofile = open('AWS_EMR_TestingInfo_092216.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['ID','Region','Name','State','CreationTime','Ec2KeyPairName','Ec2SubnetId','Ec2AvailabilityZone','IamInstanceProfile','AutoTerminate','TerminationProtected',
          'VisibleToAllUsers','EMRRole',
          'Cost Center',
          'PTC_MANAGED',
          'Owner',
          'Purpose',
          'CLUSTER_ID',
          'AGS',
          'SDLC',
          'Name',
          'Creator',
          'Provider',
          'Component',
          'Role',
          'Sonpar',
          'GIT_COMMIT',
          'GIT_BRANCH']
writer.writerow(header)

regs = ['us-east-1','us-west-2']
knowntags = ['Cost Center',
'PTC_MANAGED',
'Owner',
'Purpose',
'CLUSTER_ID',
'AGS',
'SDLC',
'Name',
'Creator',
'Provider',
'Component',
'Role',
'Sonpar',
'GIT_COMMIT',
'GIT_BRANCH'
]

for r in regs:
    client = boto3.client("emr",r)
    paginator = client.get_paginator('list_clusters')
    #response_iterator = paginator.paginate(CreatedAfter=datetime(2016,8,16))
    response_iterator = paginator.paginate(ClusterStates=['RUNNING'])
    for page in response_iterator:
        clusters = []
        clusters = page['Clusters']
        for c in clusters:
            vals = []
            emrid = c['Id']
            vals.append(emrid)
            vals.append(r)
            vals.append(c['Name'])
            vals.append(c['Status']['State'])
            vals.append(c['Status']['Timeline']['CreationDateTime'])
            cdetails = client.describe_cluster(ClusterId=emrid)
            time.sleep(.25)
            vals.append(cdetails['Cluster']['Ec2InstanceAttributes']['Ec2KeyName'])
            vals.append(cdetails['Cluster']['Ec2InstanceAttributes']['Ec2SubnetId'])
            vals.append(cdetails['Cluster']['Ec2InstanceAttributes']['Ec2AvailabilityZone'])
            vals.append(cdetails['Cluster']['Ec2InstanceAttributes']['IamInstanceProfile'])
            vals.append(cdetails['Cluster']['AutoTerminate'])
            vals.append(cdetails['Cluster']['TerminationProtected'])
            vals.append(cdetails['Cluster']['VisibleToAllUsers'])
            vals.append(cdetails['Cluster']['ServiceRole'])
            if 'Tags' in cdetails['Cluster']:
                    taglist = []
                    tagvals = []
                    taglist = cdetails['Cluster']['Tags']
                    for kt in knowntags:
                        for t in taglist:
                            if kt == t['Key']:
                                vals.append(t['Value'])
                                isfound = True
                                break
                            elif t['Key'] not in knowntags:
                                break
                                print "Need to update tags before running"
                            else:
                                isfound = False
                        if isfound == False:
                            vals.append('None')
            writer.writerow(vals)