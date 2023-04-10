import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import time

def send_email(certificate_filename, certificate_recipients, email_user, email_password, smtp_server_url, imap_server_url):

    msg = MIMEMultipart()
    msg['Subject'] = 'Mesquita Hnos - Envío de certificados de puertas cortafuego'
    msg['From'] = email_user
    msg['To'] = certificate_recipients

    with open('./templates/email_body.txt', 'r') as f:
        body = f.read()
    
    msg.attach(MIMEText(body))

    with open(certificate_filename, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='pdf')
        attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(certificate_filename))
        msg.attach(attachment)

    with smtplib.SMTP(smtp_server_url, 587) as smtp_server:
        smtp_server.set_debuglevel(1)
        smtp_server.starttls()
        smtp_server.login(email_user, email_password)
        smtp_server.sendmail(email_user, certificate_recipients.split(';'), msg.as_string())

    with imaplib.IMAP4(imap_server_url, 143) as imap_server:
        imap_server.starttls()
        imap_server.login(email_user, email_password)
        imap_server.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_string().encode('utf8'))
        imap_server.logout()