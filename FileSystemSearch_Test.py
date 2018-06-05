#package for scripting with the operating system
import os
#package for scripting with csv files
import csv
#package for dealing with time
import time
#packages for dealing with the windows security api
import win32security

#define the root directory for the script to begin the search
rootdir = r"C:/Users/F402316/Documents/"
#define a csv file object - important to note that wb is now w in Python3
ofile = open('FileShareDetails.csv', "w", newline='')
#define a csv writer object to write rows to the csv file object we just created
writer = csv.writer(ofile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
#write the header row to the csv file
writer.writerow(['Full_Path', 'File_Name', 'File_Size_Bytes', 'File_Size_MB', 'Create_Date', 'Last_Modified',
                 'Owner_Name'])

#loop through all directories and sub directories and files within the root directory specfied above
for subdir, dirs, files in os.walk(rootdir):
    #loop through all files identified
    for file in files:
        #define a list object for storage of the rows to insert into the csv file
        L = []
        #add the full file path
        L.append(os.path.join(subdir, file))
        #add the file name
        L.append(file)
        #add the file size in bytes
        L.append(os.path.getsize(os.path.join(subdir, file)))
        #add the file size in megabytes
        L.append(int(os.path.getsize(os.path.join(subdir, file))) / 1048576)
        #add the file creation date
        L.append(time.ctime(os.path.getctime(os.path.join(subdir, file))))
        #add the last modified date of the file
        L.append(time.ctime(os.path.getmtime(os.path.join(subdir, file))))
        #define a security descriptor object
        sd = win32security.GetFileSecurity(os.path.join(subdir, file), win32security.OWNER_SECURITY_INFORMATION)
        #get the file owner's SID
        owner_sid = sd.GetSecurityDescriptorOwner()
        #lookup the SID in Active Directory to resolve the name and domain
        name, domain, type = win32security.LookupAccountSid(None, owner_sid)
        #add the domain and name of the owner of the file
        L.append(domain + "\\" + name)
        #write list as a row to the csv file
        writer.writerow(L)
#close the csv file
ofile.close()
print("**************************Script Successfully Completed********************************")
