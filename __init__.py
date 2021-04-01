# coding: utf-8
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

global Amazon_Email

module = GetParams('module')

if module == "connection":

    conx = ""
    try:
        Amazon_Email = {}
        email_ = GetParams('email_')
        pass_ = GetParams('pass_')
        host_ = GetParams('host_')
        port = 587
        user_ = GetParams('user_')
        var_ = GetParams('var_')

        Amazon_Email["email"] = email_
        Amazon_Email["pass"] = pass_
        Amazon_Email["host"] = host_
        Amazon_Email["port"] = port
        Amazon_Email["user"] = user_

        server = smtplib.SMTP(Amazon_Email["host"], Amazon_Email["port"])
        server.ehlo()
        server.starttls()
        # stmplib docs recommend calling ehlo() before & after starttls()
        #server.ehlo()
        server.login(Amazon_Email["user"], Amazon_Email["pass"])
        Amazon_Email["server"] = server
        conx = True
    except:
        PrintException()
        conx = False

    SetVar(var_, conx)

if module == "sendEmail":

    Amazon_Email["server"] = smtplib.SMTP(Amazon_Email["host"], Amazon_Email["port"])
    Amazon_Email["server"].starttls()
    Amazon_Email["server"].login(Amazon_Email["user"], Amazon_Email["pass"])

    to_ = GetParams("to")
    subject = GetParams("subject")
    body_ = GetParams("msg")
    cc = GetParams('cc')
    attached_file = GetParams('path')
    files = GetParams('folder')
    filenames = []

    try:

        msg = MIMEMultipart()
        msg['From'] = Amazon_Email["email"]
        msg['To'] = to_
        msg['Cc'] = cc
        msg['Subject'] = subject

        if cc:

            toAddress = to_.split(",") + cc.split(",")
        else:
            toAddress = to_.split(",")

        if not body_:
            body_ = ""
        body = body_
        msg.attach(MIMEText(body, 'html'))

        if files:
            for f in os.listdir(files):
                f = os.path.join(files, f)
                filenames.append(f)

            if filenames:
                for file in filenames:
                    filename = os.path.basename(file)
                    attachment = open(file, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload((attachment).read())
                    attachment.close()
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                    msg.attach(part)

        else:
            if attached_file:
                if os.path.exists(attached_file):
                    filename = os.path.basename(attached_file)
                    attachment = open(attached_file, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload((attachment).read())
                    attachment.close()
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                    msg.attach(part)

        text = msg.as_string()
        Amazon_Email["server"].sendmail(Amazon_Email["email"], toAddress, text)
        # server.close()

    except Exception as e:
        PrintException()
        raise e
