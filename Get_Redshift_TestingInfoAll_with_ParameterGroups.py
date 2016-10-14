
import boto3
import httplib
import json
import sys
import os
import csv
from datetime import datetime
import time

os.environ['AWS_ACCESS_KEY_ID'] = 'xxxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxxx'
os.environ['AWS_SESSION_TOKEN'] = 'xxxx'
os.environ['https_proxy'] = 'xxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

knowntags = ['Cost Center',
             'Owner',
             'AGS',
             'SDLC',
             'Purpose',
             'workload-type'
             ]
ofile = open('AWS_Redshift_AllTestingInfo_093016.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['ID','Region','AvailabilityZone','NodeType','Status','ClusterCreateTime','MasterUsername','DBName','VpcID','MaintenanceWindow',
          'Nodes','ClusterVersion','PubliclyAccessible','Encrypted','LoggingEnabled','ParameterGroups',
          'Cost Center',
          'Owner',
          'AGS',
          'SDLC',
          'Purpose',
          'workload-type']
writer.writerow(header)

regs = ['us-east-1','us-west-2']

for r in regs:
    client = boto3.client('redshift',r)
    paginator = client.get_paginator('describe_clusters')
    response_iterator = paginator.paginate()
    for page in response_iterator:
        clusters = []
        clusters = page['Clusters']
        for c in clusters:
            vals = []
            rsid = c['ClusterIdentifier']
            vals.append(rsid)
            vals.append(r)
            vals.append(c['AvailabilityZone'])
            vals.append(c['NodeType'])
            vals.append(c['ClusterStatus'])
            vals.append(c['ClusterCreateTime'])
            vals.append(c['MasterUsername'])
            vals.append(c['DBName'])
            vals.append(c['VpcId'])
            vals.append(c['PreferredMaintenanceWindow'])
            vals.append(c['NumberOfNodes'])
            vals.append(c['ClusterVersion'])
            vals.append(c['PubliclyAccessible'])
            vals.append(c['Encrypted'])
            logval = client.describe_logging_status(ClusterIdentifier=rsid)
            vals.append(logval['LoggingEnabled'])
            if 'ClusterParameterGroups' in c:
                pglist = []
                pgvals = []
                pglist = c['ClusterParameterGroups']
                for pg in pglist:
                    presponse = client.describe_cluster_parameters(ParameterGroupName=pg['ParameterGroupName'])
                    plist = presponse['Parameters']
                    for p in plist:
                        if p['ParameterName']=='require_ssl':
                            pval = p['ParameterValue']
                            pgvals.append(pg['ParameterGroupName'] + ' : ' + pg['ParameterApplyStatus'] + ' : require_ssl=' + pval)
                        if p['ParameterName']=='enable_user_activity_logging':
                            pval = p['ParameterValue']
                            pgvals.append(pg['ParameterGroupName'] + ' : ' + pg['ParameterApplyStatus'] + ' : enable_user_activity_logging=' + pval)
                vals.append(pgvals)
            if 'Tags' in c:
                    taglist = []
                    tagvals = []
                    taglist = c['Tags']
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