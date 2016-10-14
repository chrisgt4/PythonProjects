
import boto3
import httplib
import json
import sys
import os
import csv
import datetime

os.environ['AWS_ACCESS_KEY_ID'] = 'xxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxx'
os.environ['AWS_SESSION_TOKEN'] = 'xxxx'
os.environ['https_proxy'] = 'xxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

regs = ['us-east-1','us-west-2']

ofile = open('AWS_Tags_Rpt_091216.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['Tags']
writer.writerow(header)

for r in regs:

    client = boto3.client("ec2",r)
    paginator = client.get_paginator('describe_instances')
    response_iterator = paginator.paginate()
    for page in response_iterator:
        reservations = []
        reservations = page['Reservations']
        for r in reservations:
            for i in r['Instances']:
                if 'Tags' in i:
                    taglist = []
                    taglist = i['Tags']
                    for t in taglist:
                        vals = []
                        vals.append(t['Key'])
                        writer.writerow(vals)
ofile.close()

print "================Script Successful================="