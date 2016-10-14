
import boto3
import httplib
import json
import sys
import os
import csv
import datetime

os.environ['AWS_ACCESS_KEY_ID'] = 'xxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xx'
os.environ['AWS_SESSION_TOKEN'] = 'xxxx'
os.environ['https_proxy'] = 'xxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

regs = ['us-east-1','us-west-2']

ofile = open('AWS_EMR_Tags_Rpt_092216.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['Tags']
writer.writerow(header)

for r in regs:

    client = boto3.client("emr",r)
    paginator = client.get_paginator('list_clusters')
    response_iterator = paginator.paginate(ClusterStates=['STARTING','BOOTSTRAPPING','RUNNING','WAITING'])
    for page in response_iterator:
        clusters = []
        clusters = page['Clusters']
        for c in clusters:
            emrid = c['Id']
            cdetails = client.describe_cluster(ClusterId=emrid)
            if 'Tags' in cdetails['Cluster']:
                taglist = []
                taglist = cdetails['Cluster']['Tags']
                for t in taglist:
                    vals = []
                    vals.append(t['Key'])
                    writer.writerow(vals)
ofile.close()

print "================Script Successful================="