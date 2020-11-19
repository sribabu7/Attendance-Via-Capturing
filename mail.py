import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime
import os
mail_content = '''Hello,
This is a test mail.
In this mail we are sending some attachments.
The mail is sent using Python SMTP library.
Thank You
'''
#The mail addresses and password
recmail=input('enter your mail address:')
print('Sending Mail to '+ recmail)
sender_address = 'dumpa.16.it@anits.edu.in'
sender_pass = 'sribabu_1998'
receiver_address = recmail
#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'A test mail sent by Python. It has an attachment.'
#xlsx to csv conversion
data_xls = pd.read_excel(os.getcwd()+"/firebase/attendance_files/attendance"+str(datetime.now().date())+'.xls', 'class1', index_col=None)
	#'/firebase/attendance_files/attendance'+str(datetime.now().date())+'.xls'
data_xls.to_csv('attendence.csv', encoding='utf-8')
#The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
attach_file_name = 'attendence.csv'
payload = MIMEBase('application', 'octate-stream')
payload.set_payload(open("attendence.csv","rb").read())
encoders.encode_base64(payload) #encode the attachment
#add payload header with filename
payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
message.attach(payload)
#Create SMTP session for sending the mail
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent Successfully')
