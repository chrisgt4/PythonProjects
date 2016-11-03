#Script to pull all Qualys QIDs from the API into a XML file
#Written by Chris Curtis - Internal Audit
#Initial version 1.0; dated 6/19/2015
#commented out lines are for writing to a csv instead of to a database, can be recommented back in


import sys
sys.path.append('./Lib/site-packages/')
sys.path.append('F:/Python Projects/')
import lxml.objectify
from lxml.builder import E
import qualysapi

qcnn = qualysapi.connect('config.txt')

call = '/api/2.0/fo/asset/host/'
parameters = {'action': 'list', 'ips': '10.0.0.10-10.0.0.11'}
xml_output = qcnn.request(call, parameters)
