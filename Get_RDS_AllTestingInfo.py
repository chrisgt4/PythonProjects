
import boto3
import httplib
import json
import sys
import os
import csv
from datetime import datetime
import time

os.environ['AWS_ACCESS_KEY_ID'] = 'xxxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxx'
os.environ['AWS_SESSION_TOKEN'] = 'xxxx'
os.environ['https_proxy'] = 'xxxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

def makerdsarn(region,instance):
    account = '510199193688'
    service = 'rds'
    resource = 'db'
    arn = 'arn:aws:' + service + ':' + region + ':' + account + ':' + resource + ':' + instance
    return arn

knowntags = ['Cost Center',
             'Owner',
             'workload-type',
             'AGS',
             'SDLC',
             'Purpose',
             'OWNER',
             'aws:cloudformation:stack-name',
             'aws:cloudformation:stack-id',
             'aws:cloudformation:logical-id']

ofile = open('AWS_RDS_Information_091316.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['ID','AvailabilityZone','Region','InstanceClass','Engine','EngineVersion','Status','CreateTime','MasterUsername','DBName','PreferredMaintenanceWindow','MultiAZ',
          'PubliclyAccessible','Encrypted',
          'Cost Center',
          'Owner',
          'workload-type',
          'AGS',
          'SDLC',
          'Purpose',
          'OWNER',
          'aws:cloudformation:stack-name',
          'aws:cloudformation:stack-id',
          'aws:cloudformation:logical-id'
          ]
writer.writerow(header)

regs = ['us-east-1','us-west-2']

for r in regs:
    client = boto3.client('rds',r)
    paginator = client.get_paginator('describe_db_instances')
    response_iterator = paginator.paginate()
    for page in response_iterator:
        dbs = []
        dbs = page['DBInstances']
        for d in dbs:
            vals = []
            rdsid = d['DBInstanceIdentifier']
            vals.append(rdsid)
            vals.append(d['AvailabilityZone'])
            vals.append(r)
            vals.append(d['DBInstanceClass'])
            vals.append(d['Engine'])
            vals.append(d['EngineVersion'])
            vals.append(d['DBInstanceStatus'])
            vals.append(d['InstanceCreateTime'])
            vals.append(d['MasterUsername'])
            if 'DBName' in d:
                vals.append(d['DBName'])
            else:
                vals.append('None')
            vals.append(d['PreferredMaintenanceWindow'])
            vals.append(d['MultiAZ'])
            vals.append(d['PubliclyAccessible'])
            vals.append(d['StorageEncrypted'])
            dbarn = makerdsarn(r,rdsid)
            tagdict = client.list_tags_for_resource(ResourceName=dbarn)
            if 'TagList' in tagdict:
                taglist = []
                tagvals = []
                taglist = tagdict['TagList']
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