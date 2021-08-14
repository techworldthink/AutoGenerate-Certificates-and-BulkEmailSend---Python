import csv
import os
import smtplib
import imghdr
from email.message import EmailMessage

#Sender details
#If using Gmail, make sure to turn on "Less secure app access"
email = 'xyz@xyz.com' 
password = 'password'
fromName = "Your Name"

#variables to count the number of successful mails
sent = 0
failed = 0

#give the name of csv file containing the list of those who are receiving the email.
with open ('List_emails/participants.csv','r') as plist:  
    plist_read=csv.DictReader(plist)
    for line in plist_read:
        toName=line['name']
        toEmail=line['email']
        certID=line['CID']
        attachment=line['certFileName']
        
        #Email Details
        msg = EmailMessage()
        msg['Subject'] = 'EMAIL SUBJECT'
        msg['From'] = fromName + "<" + email + ">"
        msg['To'] = toEmail

        body = 'Hi '+toName+', \n\nHope you are doing well! Certificates Attached.\n\nRegards,\nMr.X'
        msg.set_content(body)

        #Attachments
        with open (attachment,'rb') as f:
            file_data=f.read()
            file_name=f.name
        msg.add_attachment(file_data,maintype='application',subtype='octet-stream',filename=file_name)
        
        #Send Mail
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(email, password)
                smtp.send_message(msg)
            print("\nMail sent to ",toName,"(",toEmail,")","\nFile Attached:",attachment)
            sent+=1
        except:
            print("Error! : Mail not Sent to ",toName,"    ",toEmail)
            failed+=1

print()
print("REPORT")
print("Successful Mails:",sent)
print("Failed:",failed)
print()
