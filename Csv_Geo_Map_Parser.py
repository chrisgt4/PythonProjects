# Base code for csv file parsers
# Written by Chris Curtis - Internal Audit
# Last Updated April 12, 2016

import re
import sys
import csv

sourceFl = r"xxxx"

hdr = ['geobin','latitude','longitude','session_count']

runningLoopCtr = 0
wFile = open('Geo_Code_Parsed.csv','wb')
writer = csv.writer(wFile,delimiter=',', quotechar='"',quoting=csv.QUOTE_ALL)

with open(sourceFl, 'r') as readFl:
    writer.writerow(hdr)
    reader = csv.reader(readFl)
    reader.next()
    for line in reader:
        vals = []
        sessCnt = 0
        for (k,v) in enumerate(line):
            if k >= 3:
                if v != '' and v is not None:
                    sessCnt = sessCnt + int(v)
        vals.append(line[0])
        vals.append(line[1])
        vals.append(line[2])
        vals.append(sessCnt)
        writer.writerow(vals)

# close the file
readFl.close()
wFile.close()

print "----------------------------Script Successfully Run--------------------------------------"

