# System
import smtplib, os

# Mailer
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE, formatdate

# ENV
from dotenv import load_dotenv

load_dotenv()

mailer = smtplib.SMTP('smtp.gmail.com', port=587)
mailer.starttls()

mailer.login(os.getenv('SMTP_MAIL'), os.getenv('SMTP_PASS'))


msg = MIMEMultipart()
msg['From'] = os.getenv('SMTP_MAIL')
msg['To'] = COMMASPACE.join('awe30some@gmail.com')
msg['Date'] = formatdate(localtime=True)
msg['Subject'] = 'Shop mail report'

results = open('results.txt', 'rb')

attachment = MIMEApplication(results.read())
attachment['Content-Disposition'] = 'attachment; filename="results.txt"'

msg.attach(attachment)

mailer.sendmail(msg['From'], msg['To'], msg.as_string())
