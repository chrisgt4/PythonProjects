# Base code for csv file parsers
# Written by Chris Curtis - Internal Audit
# Last Updated April 12, 2016

import re
import sys
import csv

sourceFl = r"xxxxx"

runningLoopCtr = 0
wFile = open('Carbon_Black_Splunk_Log.csv','wb')
writer = csv.writer(wFile,delimiter=',', quotechar='"',quoting=csv.QUOTE_ALL)

with open(sourceFl, 'r') as readFl:
    reader = csv.reader(readFl)
    for line in reader:
        runningLoopCtr += 1
        if runningLoopCtr <= 100:
            writer.writerow(line)
        else:
            break

# close the file         
readFl.close()
wFile.close()

print "----------------------------Script Successfully Run--------------------------------------"

