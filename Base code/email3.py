# -*- coding: utf-8 -*-
"""
Created on Wed Aug  1 14:54:10 2018

@author: willi
"""


import mimetypes
import email
import email.mime.application
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib

# Create a text/plain message
msg = email.mime.multipart.MIMEMultipart()
msg['Subject'] = 'Trial'
msg['From'] = 'trendspptx@gmail.com'
msg['To'] = 'summi010@umn.edu'

# The main body is just another attachment
body = email.mime.text.MIMEText("""Hello, how are you? """)
msg.attach(body)

# Input the file location, including the file name and type
filename='C:\\Users\\willi\\Google Drive\\msba6120_ Statistics\\datafiles\\Smoking.xlsx'
fp=open(filename,'rb')

# edit subtype to replicate the document type
att = email.mime.application.MIMEApplication(fp.read(),_subtype="docx")
fp.close()
att.add_header('Content-Disposition','attachment',filename=filename)
msg.attach(att)

s = smtplib.SMTP('smtp.gmail.com')
s.starttls()

"""Setting up email: Need to change GMAIL settings to allow "less secure apps"
        https://www.google.com/settings/security/lesssecureapps"""

# Your login information
s.login('trendspptx','MSBA2019h!5')

# Email information (from address, [recipient email addresses])
s.sendmail('williamisummitt@gmail.com',['summi010@umn.edu','kowal230@umn.edu'], msg.as_string())
s.quit()


"""
Sources: 
    http://naelshiab.com/tutorial-send-email-python/
    https://docs.python.org/2/library/email.html#module-email
    http://pc2solution.blogspot.com/2014/02/python-send-mail-with-attachment-any_12.html"""