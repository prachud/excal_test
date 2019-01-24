import os
import time
import zipfile
from zipfile import ZipFile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime


# @author Daniel Vernon - daniel.vernon@wipro.com
# If any bugs are found within the code please contact the above email before changing
# The code is not modular so any changes made may break the whole code base
# TODO: Add some more error checking

def main(path,test):
    filePath = os.path.join(path,"Load")
    logLocal = os.path.join(path,"Log")
    # Where the default email is stored
    defaultFile = os.path.join(path,"Resources","defaultEmail.txt")
    print (defaultFile)
    # Where all the files to be processed are stored
    fileList = os.listdir(filePath)

    fileNames = set()
    # Ignores the email files - Tries and finds the XML files
    for file in fileList:
        if "_Email" in file or ".zip" in file or "EMAIL" in file:
            pass
        else:
            fileNames.add(file.split('.')[0])
    # If no file names found - then don't run past here
    if fileNames == set():
        print ("No valid xml or email files found in the system - please check and rerun script")
        return set()
    else:
        fileList = []
        for file in fileNames:
            fileList.append((file + ".xls",file + "_EMAIL.txt",file))

        for file in fileList:
            attachFile = file[0]
            attachFilePath = os.path.join(filePath,attachFile)
            textFile = os.path.join(filePath,file[1])
            groupID = file[2]
            zipFileX = groupID +'.zip'
            zipFilePath = os.path.join(filePath,zipFileX)

            # Zip the file that is going to be attached
            fileZipped = zipdir(filePath,attachFile,groupID)
            if not fileZipped:
                return False

            # Get the email list and fill in the correct text
            emailList,emailText,subjectText = getEmailAdd(textFile,groupID,defaultFile,test)
            print (emailList,emailText)
            # Email the data to the correct recipients
            emailSent = sendEmails(emailList,emailText,subjectText,zipFileX,zipFilePath,test)

            if not emailSent:
                print ("Email failed")
                return False

            # Write to the log (Email list, Time, Name of file) - Each month has a new log
            writeToLog(emailList,groupID,logLocal)
            # Delete used tmp (attachFile and emailFile
            if fileZipped:
                processed = os.path.join(path, "Processed", zipFileX)
                if emailSent:
                    zipMoved = False
                    while (!zipMoved):
                    try:
                        os.rename(zipFilePath, processed)
                    except:
                        print ("File already exsits - added new version")
                    os.remove(textFile)
                    os.remove(os.path.join(attachFilePath))

                else:
                    os.remove(os.path.join(path, "Processed", zipFileX))

        return fileNames

def zipdir(path,file,groupID):
    try:
        fileName = "{}.zip".format(groupID)
        os.chdir(path)
        zipf = ZipFile(fileName, 'w',zipfile.ZIP_DEFLATED)
        zipf.write(file)
        zipf.close()
        return True
    except:
        print ("Script failed to zip xls file - Check naming convention (File should just contain ID only) e.g. LGU9F6NF.xls")
        return False

def getEmailAdd(IDfile,groupID,defaultFile,test):
    file_object = open(IDfile, "r")

    text = file_object.read()
    emailList = (text.split("\n"))
    emailString = ""
    for email in emailList:
        if emailString != "":
            emailString += ','
        emailString += email

    file_object = open(defaultFile, "r")
    emailText = file_object.read()

    if test:
        subjectText = "AUTOMATION TEST MAIL PLEASE IGNORE {}".format(groupID)
    else:
        subjectText = "npower Excalibur report {}".format(groupID)


    return emailString,emailText,subjectText

def sendEmails(mailString,mailBody,mailSubject,zipFileX,zipFilePath,test):
    try:
        server = smtplib.SMTP('bridgehead.npower.com', 25)
    except:
        print ("Email server could not be connected")
        return False

    msg = MIMEMultipart()
    message = mailBody
    msg['From'] = "excalibur.no.reply@npower.com"
    msg['To'] = mailString
    msg['Subject'] = mailSubject

    msg.attach(MIMEText(message, 'plain'))

    attachment = open(zipFilePath, "rb")

    p = MIMEBase('application', 'zip')
    # To change the payload into encoded form
    p.set_payload((attachment).read())
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % zipFileX)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # send the message via the server set up earlier.
    try:
        print ("Sending email..")
        server.send_message(msg)
    except:
        print ("Email failed to send - Check email ID's")
        return False

    del msg
    server.quit()
    return True

def writeToLog(emailList,groupID,logLocal):
    fileList = os.listdir(logLocal)
    today = datetime.today()
    month = str(today.month)
    year = str(today.year)
    if len(month) == 1:
        month = '0' + month
    file = month+year+'.txt'
    filePath = os.path.join(logLocal,file)
    stringToWrite = "Timestamp: " + str(today) + ", File ID: " + groupID + ", Recipiants: " + emailList + "\n"
    if file in fileList:
        f = open(filePath,"a+")
        f.write(stringToWrite)
        f.close()
    else:
        f = open(filePath, "w+")
        f.write(stringToWrite)
        f.close()
    print ("Log entry added")



if __name__ == '__main__':
    start = time.time()
    test = False
    # Strucutre is Excal Large -> (Load,Log,Processed,Resources,tmp)
    # Hard code the path in now as it won't change

    path = "\\\\ud1.utility\\gsa\\kfcmet\\MBC\\REPORTS\\EXCAL-LARGE\\"
    if test:
        path = "C:\\Users\\Daniel\\Documents\\EXCAL-LARGE\\"
    result = main(path,test)
    if result != set():
        duration = time.time() - start
        print (("Script completed in {} seconds - The following ID's were processed:").format(round(duration,4)))
        print (result)
