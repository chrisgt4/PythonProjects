import os
import csv

rootdir = r"xxx"

ofile = open('ATL_Files_Not_Scanned.csv',"wb")
writer = csv.writer(ofile, delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)
header = ['Full_Path','File_Name','File_Size_Bytes','File_Size_MB']
writer.writerow(header)

for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        L = []
        L.append(os.path.join(subdir,file))
        L.append(file)
        L.append(os.path.getsize(os.path.join(subdir,file)))
        L.append(int(os.path.getsize(os.path.join(subdir,file)))/1048576)
        writer.writerow(L)

ofile.close()
        
print "**************************Script Successfully Completed********************************"
