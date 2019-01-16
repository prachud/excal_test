import os
from utils import zipFile, readEmailAdd,sendEmail
import time

# @author Daniel Vernon
# Script Process:
# Get file names from directory tmp being stored in
# Zip the xml file
# Get the email addresses for that file from the txt file
# Send the default email to the people on the email list
# Repeat the process until no tmp remain
# Delete the tmp that have been sent once confirmation of email being send has occurred
# Move the zip files


def main(path,test):
    email = ["praveen.vattikonda@innogy.com"]
    filePath = os.path.join(path,"tmp")
    defaultFile = os.path.join(path,"resources","defaultEmail.txt")
    fileList = os.listdir(filePath)

    fileNames = set()
    for file in fileList:
        if "_Email" in file or ".zip" in file:
            pass
        else:
            fileNames.add(file.split('.')[0])
    if fileNames == set():
        print ("No valid xml or email files found in the system - please check and rerun script")
        return set()
    else:
        fileList = []
        for file in fileNames:
            fileList.append((file + ".xls",file + "_Email.txt",file))

        for file in fileList:
            attachFile = file[0]
            attachFilePath = os.path.join(filePath,attachFile)
            textFile = os.path.join(filePath,file[1])
            groupID = file[2]
            zipFileX = groupID +'.zip'
            zipFilePath = os.path.join(filePath,zipFileX)

            # Zip the file that is going to be attached
            zipFile.zipdir(filePath,attachFile,groupID)
            # Get the email list and fill in the correct text
            emailList,emailText,subjectText = readEmailAdd.getEmailAdd(textFile,groupID,defaultFile,test)
            # Email the data to the correct recipients
            sendEmail.sendEmails(emailList,emailText,subjectText,zipFileX,zipFilePath,test)
            # Delete used tmp (attachFile and emailFile
            if not test:
                os.remove(textFile)
                os.remove(os.path.join(attachFilePath))
                os.remove(zipFilePath)

        return fileNames




if __name__ == '__main__':
    start = time.time()
    test = True
    full_path = os.path.realpath(__file__)
    path = (os.path.dirname(os.path.dirname(full_path)))
    # Run the powershell script that Naveen wrote to reduce the size of the file
   # import subprocess
    #powershellPath = "C:\\Users\da390967\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Windows PowerShell\powershell.exe"
    #test = "C:\Windows\System32\WindowsPowerShell\\v1.0\powershell.exe"
    #subprocess.call([test,"C:\\Users\\da390967\\PycharmProjects\\xmlRead\\xmlRead\\ExcelSaveScript.ps1"])
    path = os.path.join(path,"xmlRead")
    result = main(path,test)
    if result != set():
        duration = time.time() - start
        print (("Script completed in {} seconds - The following ID's were processed:").format(round(duration,4)))
        print (result)
