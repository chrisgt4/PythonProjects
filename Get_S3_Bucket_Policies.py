
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
os.environ['http_proxy'] = 'xxxx:8080'

ofile = open('AWS_S3_Bucket_Policy.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['BucketName','AccountName','Permission','GranteeFull']
writer.writerow(header)

s3 = boto3.resource('s3')
for bucket in s3.buckets.all():
    aclobj = s3.BucketAcl(bucket.name)
    grantlist = []
    grantlist = aclobj.grants
    for g in grantlist:
        vals = []
        vals.append(bucket.name)
        if 'Grantee' in g:
            g2 = g['Grantee']
            if 'DisplayName' in g2:
                vals.append(g['Grantee']['DisplayName'])
            else:
                vals.append('None')
        else:
            vals.append('None')
        if 'Permission' in g:
            vals.append(g['Permission'])
        else:
            vals.append('None')
        vals.append(g['Grantee'])
        writer.writerow(vals)
ofile.close()

print "================Script Successful================="