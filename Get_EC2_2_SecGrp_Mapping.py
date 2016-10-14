
import boto3
import httplib
import json
import sys
import os
import csv
import datetime

os.environ['AWS_ACCESS_KEY_ID'] = 'xxx'
os.environ['AWS_SECRET_ACCESS_KEY'] = 'xxx'
os.environ['AWS_SESSION_TOKEN'] = 'xxx'
os.environ['https_proxy'] = 'xxx:8080'
os.environ['http_proxy'] = 'xxx:8080'

regs = ['us-east-1','us-west-2']

ofile = open('AWS_All_EC2_2_SecGrp_Mapping_082416.csv',"wb")
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
header = ['InstanceID','State','Instance_Profile','AMI','InstanceType','KeyPairName','LaunchTime','Monitoring','AvailabilityZone'
          ,'PrivateDnsName','PrivateIpAddress','PublicDnsName','PublicIpAddress','SecurityGroups','SubnetId','VpcId','Tags']
writer.writerow(header)

for r in regs:
    reservations = boto3.client("ec2",r).describe_instances()["Reservations"]
    for reservation in reservations:
        for i in reservation["Instances"]:
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
                for t in taglist:
                    tagvals.append(t['Key'] + '|' + t['Value'])
                vals.append(tagvals)
            else:
                vals.append('None')
            writer.writerow(vals)

ofile.close()

print "================Script Successful================="