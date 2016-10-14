
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
knowntags = ['Owner',
'Name',
'CLUSTER_ID',
'PTC_MANAGED',
'SDLC',
'aws:elasticmapreduce:instance-group-role',
'aws:elasticmapreduce:job-flow-id',
'Cost Center',
'Purpose',
'AGS',
'Creator',
'Component',
'NagiosSNS',
'aws:autoscaling:groupName',
'role',
'ConfigBucket',
'Provider',
'env_profile',
'StackId',
'GIT_BRANCH',
'GIT_COMMIT',
'Sonpar',
'Puppetenv',
'aws:cloudformation:stack-id',
'aws:cloudformation:logical-id',
'aws:cloudformation:stack-name',
'Contact',
'IAM Role',
'Deployment',
'ZKInstanceID',
'ZKGroupID',
'Reference',
'Cost Center ',
'Dummy Change',
'rds-endpoint',
'CF Update Dummy',
'SLDC'
]

ofile = open('AWS_Hadoop_EC2_Rpt_092916.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['InstanceID','State','Instance_Profile','AMI','InstanceType','KeyPairName','LaunchTime','Monitoring','AvailabilityZone'
          ,'PrivateDnsName','PrivateIpAddress','PublicDnsName','PublicIpAddress','SecurityGroups','SubnetId','VpcId',
            'Owner',
            'Name',
            'CLUSTER_ID',
            'PTC_MANAGED',
            'SDLC',
            'aws:elasticmapreduce:instance-group-role',
            'aws:elasticmapreduce:job-flow-id',
            'Cost Center',
            'Purpose',
            'AGS',
            'Creator',
            'Component',
            'NagiosSNS',
            'aws:autoscaling:groupName',
            'role',
            'ConfigBucket',
            'Provider',
            'env_profile',
            'StackId',
            'GIT_BRANCH',
            'GIT_COMMIT',
            'Sonpar',
            'Puppetenv',
            'aws:cloudformation:stack-id',
            'aws:cloudformation:logical-id',
            'aws:cloudformation:stack-name',
            'Contact',
            'IAM Role',
            'Deployment',
            'ZKInstanceID',
            'ZKGroupID',
            'Reference',
            'Cost Center ',
            'Dummy Change',
            'rds-endpoint',
            'CF Update Dummy',
            'SLDC']
writer.writerow(header)

for r in regs:

    client = boto3.client("ec2",r)
    paginator = client.get_paginator('describe_instances')
    response_iterator = paginator.paginate(Filters=[{'Name':'instance-state-name','Values':['running']}])
    for page in response_iterator:
        reservations = []
        reservations = page['Reservations']
        for r in reservations:
            for i in r['Instances']:
                vals = []
                vals.append(i['InstanceId'])
                vals.append(i['State']['Name'])
                if 'IamInstanceProfile' in i:
                    iamipvals = []
                    iamipvals.append(i['IamInstanceProfile']['Id'] + '|' + i['IamInstanceProfile']['Arn'])
                    vals.append(iamipvals)
                else:
                    vals.append(['None'])
                vals.append(i['ImageId'])
                vals.append(i['InstanceType'])
                if 'KeyName' in i:
                    vals.append(i['KeyName'])
                else:
                    vals.append('None')
                vals.append(i['LaunchTime'])
                vals.append(i['Monitoring']['State'])
                vals.append(i['Placement']['AvailabilityZone'])
                if 'PrivateDnsName' in i:
                    vals.append(i['PrivateDnsName'])
                else:
                    vals.append('None')
                if 'PrivateIpAddress' in i:
                    vals.append(i['PrivateIpAddress'])
                else:
                    vals.append('None')
                if 'PublicDnsName' in i:
                    vals.append(i['PublicDnsName'])
                else:
                    vals.append('None')
                if 'PublicIpAddress' in i:
                    vals.append(i['PublicIpAddress'])
                else:
                    vals.append('None')
                if 'SecurityGroups' in i:
                    secgrps = []
                    secgrpvals = []
                    secgrps = i['SecurityGroups']
                    for sg in secgrps:
                        secgrpvals.append(sg['GroupId'] + '|' + sg['GroupName'])
                    vals.append(secgrpvals)
                else:
                    vals.append('None')
                if 'SubnetId' in i:
                    vals.append(i['SubnetId'])
                else:
                    vals.append('None')
                if 'VpcId' in i:
                    vals.append(i['VpcId'])
                else:
                    vals.append('None')
                if 'Tags' in i:
                    taglist = []
                    tagvals = []
                    taglist = i['Tags']
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

ofile.close()

print "================Script Successful================="