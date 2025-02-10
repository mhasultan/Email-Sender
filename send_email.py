import pandas as pd
import os
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging
import time
import sys
import imgkit
from PIL import Image
import string
from random import randint, choices
import base64
import random

# Set the path to the wkhtmltopdf executable
path_wkhtmltoimage = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe'
config = imgkit.config(wkhtmltoimage=path_wkhtmltoimage)

# Set up logging
logging.basicConfig(filename='mail.log', level=logging.DEBUG)

# Total number of emails to send
totalSend = 1
if len(sys.argv) > 1:
    totalSend = int(sys.argv[1])

# Read CSV files
emaildf = pd.read_csv('gmail.csv', encoding='utf-8')
contactsData = pd.read_csv('contacts.csv', encoding='utf-8')
subjects = pd.read_csv('subjects.csv', encoding='utf-8')
bodies = ['body.txt']
From = ["jonny smith#", "john smith#", "anny smith#", "adam smith#"]

# Gmail API credentials
credentials_path = 'credentials.json'
token_path = 'token.json'
scopes = ['https://www.googleapis.com/auth/gmail.send']

# Replace placeholders in content
def replace_placeholders(content):
    placeholders = {
        '$RTX': ''.join(choices(string.ascii_uppercase, k=5)) + '-' + ''.join(choices(string.ascii_uppercase, k=3)) + '-' + str(randint(1000000000, 9999999999)) + '-' + str(randint(10000000, 99999999)),
        '$SNM': random.choice(["John", "Emma", "Olivia", "Michael", "Sophia", "James", "Isabella", "Benjamin", "sia", "Alexander"])
    }
    for placeholder, value in placeholders.items():
        content = content.replace(placeholder, value)
    return content

# Create a Gmail API service object
def create_gmail_service():
    flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
    creds = flow.run_local_server(port=0)
    with open(token_path, 'w', encoding='utf-8') as token:
        token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

# Send an email using the Gmail API
def send_mail(service, name, email, emailId, password, bodyFile, subjectWord, fromName):
    message = MIMEMultipart()
    subject = replace_placeholders(subjectWord)
    message['Subject'] = subject
    message['From'] = emailId
    message['To'] = email
    
    with open(bodyFile, 'r', encoding='utf-8') as f:
        body = replace_placeholders(f.read())
    
    message.attach(MIMEText(body, 'plain'))
    
    html = open('html_code.html', 'r', encoding='utf-8').read()
    html = replace_placeholders(html)
    temp_html = 'temp_invoice.html'
    with open(temp_html, 'w', encoding='utf-8') as f:
        f.write(html)
    
    jpg_file = f"invoice_{randint(10000000000, 99999999999)}.jpg"
    imgkit.from_file(temp_html, jpg_file, config=config)
    pdf_file = jpg_file.replace('.jpg', '.pdf')
    
    with Image.open(jpg_file) as img:
        img.convert('RGB').save(pdf_file)
    os.remove(temp_html)
    os.remove(jpg_file)
    
    with open(pdf_file, 'rb') as f:
        payload = MIMEBase('application', 'octet-stream', Name=pdf_file)
        payload.set_payload(f.read())
        encoders.encode_base64(payload)
        payload.add_header('Content-Disposition', 'attachment', filename=pdf_file)
        message.attach(payload)
    
    try:
        service.users().messages().send(userId='me', body={'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}).execute()
        os.remove(pdf_file)
        print(f"Email successfully sent to {email} from {emailId}")
        logging.info(f"Sent to {email} by {emailId} successfully : {totalSend}")
    except Exception as e:
        print(f"Error sending email to {email} by {emailId}")
        logging.exception(f"Error sending email to {email} by {emailId}")
        print(str(e))

def start_mail_system(service):
    j, k, l, m = 0, 0, 0, 0
    for i in range(len(contactsData)):
        if j >= len(emaildf):
            j = 0
        send_mail(
            service,
            contactsData.iloc[i]['name'],
            contactsData.iloc[i]['email'],
            emaildf.iloc[j]['email'],
            emaildf.iloc[j]['password'],
            bodies[k],
            subjects.iloc[l]['subject'],
            From[m]
        )
        j = (j + 1) % len(emaildf)
        k = (k + 1) % len(bodies)
        l = (l + 1) % len(subjects)
        m = (m + 1) % len(From)

def remove_email(emailId):
    df = pd.read_csv('gmail.csv', encoding='utf-8')
    df = df[df['email'] != emailId]
    df.to_csv('gmail.csv', index=False, encoding='utf-8')
    print(f"{emailId} removed from gmail.csv")
    logging.info(f"{emailId} removed from gmail.csv")

try:
    service = create_gmail_service()
    start_mail_system(service)
except KeyboardInterrupt:
    print("\n\nCode stopped by user")
