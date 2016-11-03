# DATA620 Assignment 12.1
# Section 9043
# Written by Chris Curtis
# Last Updated April 19, 2016

# -------------------------------------------------------------------------------------------------------------------------
# ASSUMPTIONS
# Program will do the following:
# - analyze a csv file containing over 7000 publicly released emails for Hilary Clinton
# - attempt to remove extraneous lines from the raw text of the email data that can skew the word counts
# - attempt to remove punctuation characters from the first and last character of a word since these cause the words
#   to register as a new word that can skew the counts
# - produce the total email counts in the data set used across all years of coverage
# - identify unique words used in the body of the email and count the number of teams each was used
# - exclude words found in the badWords list from the overall analysis
# - excludes words that are blank or null
# - excludes emails that have a blank or null date sent
# - produce an output of the top 30 words used and the number of times used for each year
# - excludes the years 2008 and 2014 due to the low population count in the dataset (1 and 2 respectively)
# - produce two csv files of the results including the top 30 words w/frequency and all words w/frequency
# -------------------------------------------------------------------------------------------------------------------------

# --------------------------------------------------------BEGIN PROGRAM----------------------------------------------------
# module for reading and writing csv files
import csv

# defines the file used in analysis
sourceFl = r"/Users/christophercurtis/Documents/Masters Degree/DATA620/Week 12/output/Emails.csv"
# defins a list of punctuation characters to be used to strip from the first and last character of the words
puncChars = ["'",'"',':',';',',','?','!','(',')','[',']','.','`','^','_','-','‘','„','*','\\','/','&','#','%','•','$']
# defines the list of words to exclude from the frequency analysis
badWords = ['the','to','of','and','a','in','with','as','that','u.s.','unclassified','department','state','release','full',
            'no','is','this','on','for','he','we','are','was','i','by','from','his','have','will','at','be','it','our',
            'has','an','you','not','they','but','who','their','new','these','','also','part','been','b6','were','2011','or',
            'its','may','would','according','all','can','about','had','what','out','-','there','said','so','2012','one',
            'very','more','other','do','no.','case','date','doc','•','while','during','if','—','am','which','she','us',
            'when','her','up','pm','just','your','after','2010','than','my','any','me','some','like','could','should',
            'get','which','over','well','2009','see','know','time','need','him','want','last','how','--','now','think',
            'two','them','only','make','most','>','against','same','between','into','being','did',"i'm",'many','told',
            'through','w','going','much','those','next','march','into','even','good','first','where','before','take',
            'h','because','sent','both','made','(source','source,','let','way','back','still','such',"it's",'back',
            'asked','then',"don't",'subject','go','come','work','believes','former','september','"the','working','under',
            're:','hope','talk','call','since','called','today','please','email','say','fyi','u.s','mr','here','week',
            'tomorrow','years','1','below','best','re']

# initialize an empty dictionary for all unique words identified
allWords = {}

# populate a dictionary to store the email count by year
countByYear = {'2008':0,'2009':0,'2010':0,'2011':0,'2012':0,'2014':0}
# open the file in read-only mode; denoting different lines using the new lines Python character;
# important to note that the file is opened with utf-8 encoding instead of ascii to handle special characters
with open(sourceFl, newline='\n', encoding='utf-8') as readFl:
    # define an object to read the csv file
    reader = csv.reader(readFl)
    # skip the first line that contains the column headers
    columns = next(reader)
    # loop through all lines in the reader object
    for data in reader:
        # store the date sent field in a string variable
        dtSent = data[6]
        # strip the first 4 characters that denotes the year
        eYear = dtSent[:4]
        # populate a string variable with the raw text data
        eText = data[21]
        # exclude emails where the year is blank or null
        if eYear != '' and eYear != None:
            # increment the dictionary storing the email counts by year
            countByYear[eYear] += 1
            # split the raw text into a list using all line feed / carriage return characters
            lines = eText.splitlines()
            # loop through all the lines
            for l in lines:
                # skip lines that are blank or null
                if l != None or line != '':
                    # exclude extraneous lines found in the raw text field of the email that can skew the counts
                    # BEGIN LINE EXCLUSIONS
                    if l.startswith('UNCLASSIFIED') == True:
                        continue
                    elif l.startswith('U.S. Department of State') == True:
                        continue
                    elif l.startswith('STATE DEPT. - PRODUCED TO HOUSE SELECT BENGHAZI COMM.') == True:
                        continue
                    elif l.startswith('SUBJECT TO AGREEMENT ON SENSITIVE INFORMATION & REDACTIONS.') == True:
                        continue
                    elif l.startswith('RELEASE IN') == True:
                        continue
                    elif l.startswith('B6') == True:
                        continue
                    elif l.startswith('B5') == True:
                        continue
                    elif l.startswith('Original Message') == True:
                        continue
                    elif l.startswith('Case No.') == True:
                        continue
                    elif l.startswith('Case No') == True:
                        continue
                    elif l.startswith('Doc No.') == True:
                        continue
                    elif l.startswith('Doc No') == True:
                        continue
                    elif l.startswith('Date:') == True:
                        continue
                    elif l.startswith('Subject:') == True:
                        continue
                    elif l.startswith('From:') == True:
                        continue
                    elif l.startswith('Sent:') == True:
                        continue
                    elif l.startswith('To:') == True:
                        continue
                    elif l.startswith('CONFIDENTIAL') == True:
                        continue
                    elif l.startswith('For:') == True:
                        continue
                    elif l.startswith('Cc:') == True:
                        continue
                    # END LINE EXCLUSIONS
                    # split the current line into separate words
                    words = l.split(" ")
                    # loop through all the words
                    for w in words:
                        # convert the words to lower case for conformity in the counts
                        # Python is case sensitive when interpreting strings
                        word = w.lower()
                        # strip punctuation characters from the beginning or ending of the word
                        # only if the word is not blank or null and not in the list of punctuation characters
                        # otherwise skip the word
                        # has a try except in case it encounters a word that is comprised only of punctuation characters
                        # if such a word is found program will skip it
                        if word != '' and word != None and word not in puncChars:
                            try:
                                while word[-1] in puncChars and len(word) != 0:
                                    if word[-1] in puncChars:
                                        word = word[:-1]
                                while word[0] in puncChars and len(word) != 0:
                                    if word[0] in puncChars:
                                        word = word[1:]
                            except:
                                continue
                        else:
                            continue
                        # populate the first level of the dictionary with keys that denote the different years
                        # if the year does not already exist in the dictionary keys
                        if eYear not in allWords:
                            # if this is the first time the year is entered as a key
                            # populate a nested dictionary with the word and start the count at 1
                            allWords[eYear] = {word:1}
                        # if the year does exist
                        else:
                            # check to see if the word is already in the nested dictionary storing counts of unique words by year
                            if word not in allWords[eYear]:
                                # if the word is unique, add a new entry for the word and start the count at 1
                                allWords[eYear][word] = 1
                            # if the word already exists
                            else:
                                # increment the running counter for the word by 1
                                allWords[eYear][word] += 1

# print out the count of emails by year stored in the dictionary
print('Email Count by Years')
for d in countByYear:
    print(d + ' : ' + str(countByYear[d]))
print('\r')

# initialize a variables for adding a record ID to the csv output files
ctr2 = 0

# create a new csv file for writing the top 30 words identified in the emails by year
with open('ClintonWordAnalysisTop30.csv','w', newline='\n') as csvfile:
    # create a second csv file for writing the frequency of all words identified in the emails by year
    csvfile2 = open('ClintonWordAnalysisAll.csv','w',newline='\n', encoding='utf-8')
    # specify the header fields
    hdr = ['RecordID','Year','Word','Frequency']
    # define two objects to write to the csv files
    writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    writer2 = csv.writer(csvfile2, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # write the header to the csv files
    writer.writerow(hdr)
    writer2.writerow(hdr)
    # loop through all the year keys found in the dictionary
    for eYear in allWords:
        # exclude years 2014 and 2008 due to the low population of emails in these years
        if eYear not in ('2014','2008'):
            # define a list that pulls the words and counts from the nested dictionary only if the word is not excluded
            # the dictionary was converted to a list for easy sorting
            lst = [(allWords[eYear][word], word) for word in allWords[eYear] if word not in badWords]
            # sort the list in descending order
            lst.sort()
            lst.reverse()
            # start a counter 
            ctr = 0
            # print the out of the top 30 words by year and write to the first csv output file
            print(eYear + ' Top 30 Words')
            for count, word in lst[:30]:
                ctr += 1
                print(str(ctr) + '. ' + word + ' : ' + str(count))
                writer.writerow([str(ctr),eYear,word,str(count)])
            print('\r')
            # write the output of the counts of all words into the second csv output file
            for count, word in lst:
                ctr2 += 1
                writer2.writerow([str(ctr2),eYear,word,str(count)])

# close all the files
csvfile.close()
csvfile2.close()
readFl.close()

print('----------------Script Successfully Run-----------------')

# ---------------------------------------------------END PROGRAM-------------------------------------------------------
