

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

    #subjectText = "npower Excalibur report {}".format(groupID)
    if test:
        subjectText = "AUTOMATION TEST MAIL PLEASE IGNORE {}".format(groupID)
    else:
        subjectText = "npower Excalibur report {}".format(groupID)


    return emailString,emailText,subjectText
