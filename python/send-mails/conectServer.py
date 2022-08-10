import os
import smtplib

def getServer():
    print("connecting to server...")

    server = smtplib.SMTP_SSL(os.environ['SMTP_SERVER'], os.environ['SMTP_PORT'])
    server.login(os.environ["EMAIL_ADDRESS"], os.environ["EMAIL_PASSWORD"])

    print("done")

    return server