import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def sendEmails(mailString,mailBody,mailSubject,zipFileX,zipFilePath,test):


    if test:
        mailString = "praveen.vattikonda@innogy.com"
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
    server.send_message(msg)

    del msg
    server.quit()

