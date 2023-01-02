import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
def send_mail(user,date,begin,duration,to):
    mail_content = f"Hello {user},\nYou'll have courses the {date} at {begin}, the duration is {duration} hours"
    #The mail addresses and password
    sender_address = 'sender123@gmail.com'
    sender_pass = 'xxxxxxxx'
    receiver_address = 'receiver567@gmail.com'
    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = f'Courses the {date} at {begin}'   #The subject line
    #The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'plain'))
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login("esilvdevinci@gmail.com", "gftwuelxhhacvton") #login with mail_id and password
    text = message.as_string()
    session.sendmail("esilvdevinci@gmail.com", to, text)
    session.quit()
    print('Mail Sent')