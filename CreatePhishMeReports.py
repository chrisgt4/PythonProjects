# Script to parse PhishMe scenario reports to DPP format for executive consumption
# Written by Chris Curtis, Internal Audit
# Last updated on 3/21/16
# Initial Version 1.0

import sys
import csv
import re
from Tkinter import *
from tkMessageBox import *
from tkFileDialog import askopenfilename

root = Tk()

while True:
    sourceFl = askopenfilename()
    if sourceFl == '':
        root.destroy()
        sys.exit()
    else:
        flNmChk = re.compile(r"^.+\.csv")
        if flNmChk.match(sourceFl) == None:
            print "Incorrect file type, please select a .csv file"
            showerror("Load Error","Incorrect file type, please select a .csv file")
        else:
            root.destroy()
            break

lineCtr = 0
hasHdrCnt = 0
hdrCtr = 0
with open(sourceFl,"r") as readFl:
    reader = csv.reader(readFl)
    for data in reader:
        lineCtr = lineCtr + 1
        if lineCtr == 1:
            phishType = None
            for hdrNm in data:
                if hdrNm == "Clicked Link?":
                    phishType = "Click Only"
                elif hdrNm == "Viewed Education?":
                    phishType = "Attachment"
            if phishType == None:
                print "The selected file is not a PhishMe report"
                root = Tk()
                showerror("Error","The selected file is not a PhishMe report")
                root.destroy()
                sys.exit()
            if phishType == "Click Only":
                for hdrNm in data:
                    if hdrNm == "Opened Email?":
                        hasHdrCnt = hasHdrCnt + 1
                        openedEmailHdr = hdrCtr
                    elif hdrNm == "Clicked Link?":
                        hasHdrCnt = hasHdrCnt + 1
                        clickedLinkHdr = hdrCtr
                    elif hdrNm == "Reported Phish?":
                        hasHdrCnt = hasHdrCnt + 1
                        reportedPhishHdr = hdrCtr
                    hdrCtr = hdrCtr + 1
                if hasHdrCnt != 3:
                    print "The format of the PhishMe report has changed, please contact Chris Curtis to update the program"
                    root = Tk()
                    showerror("Error","The format of the PhishMe report has changed, please contact Chris Curtis to update the program")
                    root.destroy()
                    sys.exit()
            elif phishType == "Attachment":
                for hdrNm in data:
                    if hdrNm == "Opened Email?":
                        hasHdrCnt = hasHdrCnt + 1
                        openedEmailHdr = hdrCtr
                    elif hdrNm == "Viewed Education?":
                        hasHdrCnt = hasHdrCnt + 1
                        viewedEducHdr = hdrCtr
                    elif hdrNm == "Reported Phish?":
                        hasHdrCnt = hasHdrCnt + 1
                        reportedPhishHdr = hdrCtr
                    hdrCtr = hdrCtr + 1
                if hasHdrCnt != 3:
                    print "The format of the PhishMe report has changed, please contact Chris Curtis to update the program"
                    root = Tk()
                    showerror("Error","The format of the PhishMe report has changed, please contact Chris Curtis to update the program")
                    root.destroy()
                    sys.exit()
            else:
                print "The selected file is not a PhishMe report"
                root = Tk()
                showerror("Error","The selected file is not a PhishMe report")
                root.destroy()
                sys.exit()
            if phishType == "Click Only":
                rptFl1 = open('ImmediatelyRecognizedAndReported.csv','wb')
                rptFl2 = open('OpenedEmailAndReported.csv','wb')
                rptFl3 = open('OpenedEmailAndDidNotReport.csv','wb')
                rptFl4 = open('ClickedOnLinkAndReported.csv','wb')
                rptFl5 = open('ClickedOnLinkButDidNotReport.csv','wb')
                rptFl6 = open('SummaryStats.csv','wb')
                writer1 = csv.writer(rptFl1, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer2 = csv.writer(rptFl2, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer3 = csv.writer(rptFl3, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer4 = csv.writer(rptFl4, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer5 = csv.writer(rptFl5, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer6 = csv.writer(rptFl6, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                # write the column headers to the file
                writer1.writerow(data)
                writer2.writerow(data)
                writer3.writerow(data)
                writer4.writerow(data)
                writer5.writerow(data)
                rpt1Count = 0
                rpt2Count = 0
                rpt3Count = 0
                rpt4Count = 0
                rpt5Count = 0
            elif phishType == "Attachment":
                rptFl1 = open('ImmediatelyRecognizedAndReported.csv','wb')
                rptFl2 = open('OpenedEmailAndReported.csv','wb')
                rptFl3 = open('OpenedEmailAndDidNotReport.csv','wb')
                rptFl4 = open('OpenedAttachmentAndReported.csv','wb')
                rptFl5 = open('OpenedAttachmentButDidNotReport.csv','wb')
                rptFl6 = open('SummaryStats.csv','wb')
                writer1 = csv.writer(rptFl1, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer2 = csv.writer(rptFl2, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer3 = csv.writer(rptFl3, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer4 = csv.writer(rptFl4, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer5 = csv.writer(rptFl5, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                writer6 = csv.writer(rptFl6, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
                # write the column headers to the file
                writer1.writerow(data)
                writer2.writerow(data)
                writer3.writerow(data)
                writer4.writerow(data)
                writer5.writerow(data)
                rpt1Count = 0
                rpt2Count = 0
                rpt3Count = 0
                rpt4Count = 0
                rpt5Count = 0
        # start of logic to break up the file into separate csv files
        if phishType == "Click Only":
            openedEmail = data[openedEmailHdr]
            clickedLink = data[clickedLinkHdr]
            reportedPhish = data[reportedPhishHdr]
            if openedEmail == "No" and clickedLink == "No" and reportedPhish == "Yes":
                writer1.writerow(data)
                rpt1Count = rpt1Count + 1
            elif openedEmail == "Yes" and clickedLink == "No" and reportedPhish == "Yes":
                writer2.writerow(data)
                rpt2Count = rpt2Count + 1
            elif openedEmail == "Yes" and clickedLink == "No" and reportedPhish == "No":
                writer3.writerow(data)
                rpt3Count = rpt3Count + 1
            elif clickedLink == "Yes" and reportedPhish == "Yes":
                writer4.writerow(data)
                rpt4Count = rpt4Count + 1
            elif openedEmail == "No" and clickedLink == "Yes" and reportedPhish == "No":
                writer5.writerow(data)
                rpt5Count = rpt5Count + 1
        elif phishType == "Attachment":
            openedEmail = data[openedEmailHdr]
            viewedEduc = data[viewedEducHdr]
            reportedPhish = data[reportedPhishHdr]
            if openedEmail == "No" and viewedEduc == "No" and reportedPhish == "Yes":
                writer1.writerow(data)
                rpt1Count = rpt1Count + 1
            elif openedEmail == "Yes" and viewedEduc == "No" and reportedPhish == "Yes":
                writer2.writerow(data)
                rpt2Count = rpt2Count + 1
            elif openedEmail == "Yes" and viewedEduc == "No" and reportedPhish == "No":
                writer3.writerow(data)
                rpt3Count = rpt3Count + 1
            elif viewedEduc == "Yes" and reportedPhish == "Yes":
                writer4.writerow(data)
                rpt4Count = rpt4Count + 1
            elif openedEmail == "No" and viewedEduc == "Yes" and reportedPhish == "No":
                writer5.writerow(data)
                rpt5Count = rpt5Count + 1
    if phishType == "Click Only":
        writer6.writerow(["Immediately recognized and reported: " + str(rpt1Count)])
        writer6.writerow(["Opened email and reported: " + str(rpt2Count)])
        writer6.writerow(["Opened email but did not report: " + str(rpt3Count)])
        writer6.writerow(["Clicked on link and reported: " + str(rpt4Count)])
        writer6.writerow(["Clicked on link but did not report: " + str(rpt5Count)])
    elif phishType == "Attachment":
        writer6.writerow(["Immediately recognized and reported: " + str(rpt1Count)])
        writer6.writerow(["Opened email and reported: " + str(rpt2Count)])
        writer6.writerow(["Opened email but did not report: " + str(rpt3Count)])
        writer6.writerow(["Opened attachment and reported: " + str(rpt4Count)])
        writer6.writerow(["Opened attachment but did not report: " + str(rpt5Count)])

readFl.close()
rptFl1.close()
rptFl2.close()
rptFl3.close()
rptFl4.close()
rptFl5.close()
rptFl6.close()

print "----------------------Script Successfully Run-----------------------------"
