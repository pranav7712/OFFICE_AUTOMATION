# import all required modules

import smtplib
from getpass import getpass
from email.message import EmailMessage

# define sender's email and password

senders_email = input("Enter your E-mail Id: ")
senders_password = getpass("Enter you password ")

receivers_mail = ['shubham26032000@gmail.com', 'shubham2000.03.26@gmail.com', 'shubham2000.09.26@gmail.com']
sub = ['TEST', 'TEST2', 'FINAL TEST']
attach_files = ['1.jpg', '2.jpg', '3.jpg']

zipped = zip(receivers_mail, sub, attach_files)

# create loop

for (a, b, c) in zipped:
    msg = EmailMessage()
    with open(c, 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg['From'] = senders_email
    msg['To'] = a
    msg['Subject'] = b
    msg.set_content(f'Hello ! How are you??')
    msg.add_attachment(file_data, maintype='img', subtype='imghdr', filename=f.name)

    # login and send the mails
    with smtplib.SMTP_SSL('smtp.gmail.com', '465') as smtp:
        smtp.login(senders_email, senders_password)
        smtp.send_message(msg)

print("All mails sent!")