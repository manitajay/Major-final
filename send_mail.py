
import os
from email.message import EmailMessage 
import ssl
import smtplib


email_receiver='ajaychuruwala@gmail.com'


def send_mail(subject , body , email_receiver):

    # subject='Warning mail for attendance'
    # body="please attend classes"
    
    email_sender='bhakarajay360@gmail.com'
    email_password='rzwvwkmsbbpseuxa'

    em=EmailMessage()
    em['From']=email_sender
    em['To']=email_receiver

    em['Subject']=subject
    em.set_content(body)


    context=ssl.create_default_context()

    # em['From']=email_sender
    # em['To']=email_receiver
    # em['Subject']=subject
    # em.set.content(body)


    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
        smtp.login(email_sender,email_password)
        smtp.sendmail(email_sender,email_receiver,em.as_string())







