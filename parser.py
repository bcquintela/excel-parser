#!/usr/bin/python

from cgitb import html
from datetime import date, datetime
import pandas as pd
import sys
import smtplib
from email.utils import formataddr
from email.mime.text import MIMEText

row = int(sys.argv[1])

sender = 'sender@domain'
receiver = pd.read_excel('file.xlsx', skiprows=row - 1, usecols="E", nrows=1, header=None, names=["Value"]).iloc[0]["Value"]

text = pd.read_excel('file.xlsx', skiprows=row - 1, usecols="M", nrows=1, header=None, names=["Value"]).iloc[0]["Value"]

msg = MIMEText(text)

msg['Subject'] = 'Subject'
msg['From'] = formataddr(('Sender', sender))
msg['To'] = receiver

password = 'password'

print("Sending email to " + receiver + " ...")

with smtplib.SMTP('mail.domain', 587) as server:
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(sender, password)
    server.sendmail(sender, receiver, msg.as_string())
    server.quit()
    print("Email successfully sent!")
    print(datetime.now())