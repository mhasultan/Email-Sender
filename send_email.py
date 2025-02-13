import pandas as pd
import os
import base64
import logging
from logging.handlers import RotatingFileHandler
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from PIL import Image
import imgkit
from concurrent.futures import ThreadPoolExecutor
import threading

# Set path for wkhtmltoimage
path_wkhtmltoimage = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe'
config = imgkit.config(wkhtmltoimage=path_wkhtmltoimage)

# Logging setup with both file and console handlers
logger = logging.getLogger('EmailSender')
logger.setLevel(logging.DEBUG)

file_handler = RotatingFileHandler('mail.log', maxBytes=10485760, backupCount=5)
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler()
console_formatter = logging.Formatter('%(message)s')
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

CREDENTIALS_PATH = 'credentials.json'
TOKEN_PATH = 'token.json'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

BATCH_SIZE = 50
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds

def create_gmail_service():
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        if not creds.expired:
            return build('gmail', 'v1', credentials=creds)

    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_PATH, 'w', encoding='utf-8') as token:
        token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def generate_pdf(html_content, unique_id):
    temp_html = f'temp_invoice_{unique_id}.html'
    jpg_file = f'invoice_{unique_id}.jpg'
    pdf_file = jpg_file.replace('.jpg', '.pdf')
    
    try:
        with open(temp_html, 'w', encoding='utf-8') as f:
            f.write(html_content)
        options = {'quality': 85, 'width': 1024, 'enable-local-file-access': None}
        imgkit.from_file(temp_html, jpg_file, config=config, options=options)
        with Image.open(jpg_file) as img:
            img.convert('RGB').save(pdf_file, optimize=True)
        return pdf_file
    except Exception as e:
        logger.error(f"Error generating PDF: {str(e)}")
        raise
    finally:
        for file in [temp_html, jpg_file]:
            if os.path.exists(file):
                os.remove(file)

def send_single_mail(service, contact_data, email_data, body_content, subject, from_name, success_count, failure_count):
    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = email_data['email']
    message['To'] = contact_data['email']
    message['Reply-To'] = email_data['email']
    message.attach(MIMEText(body_content, 'plain'))
    unique_id = threading.get_ident()
    pdf_file = None
    
    try:
        with open('html_code.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        pdf_file = generate_pdf(html_content, unique_id)
        
        with open(pdf_file, 'rb') as f:
            payload = MIMEBase('application', 'octet-stream', Name=pdf_file)
            payload.set_payload(f.read())
            encoders.encode_base64(payload)
            payload.add_header('Content-Disposition', 'attachment', filename=pdf_file)
            message.attach(payload)
        
        for attempt in range(MAX_RETRIES):
            try:
                service.users().messages().send(
                    userId='me',
                    body={'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
                ).execute()
                success_count.append(1)
                print(f"✅ Email successfully sent to {contact_data['email']} from {email_data['email']}")
                logger.info(f"Email successfully sent to {contact_data['email']} from {email_data['email']}")
                return
            except Exception as e:
                if attempt == MAX_RETRIES - 1:
                    failure_count.append(1)
                    print(f"❌ Failed to send email to {contact_data['email']} - {str(e)}")
                    logger.error(f"Failed to send email to {contact_data['email']} - {str(e)}")
                time.sleep(RETRY_DELAY)
    finally:
        if pdf_file and os.path.exists(pdf_file):
            os.remove(pdf_file)

def process_batch(service, batch_data, email_data, body_content, subjects, from_names, success_count, failure_count):
    with ThreadPoolExecutor(max_workers=BATCH_SIZE) as executor:
        futures = []
        for i, contact in enumerate(batch_data):
            futures.append(executor.submit(
                send_single_mail,
                service,
                contact,
                email_data[i % len(email_data)],
                body_content,
                subjects[i % len(subjects)],
                from_names[i % len(from_names)],
                success_count,
                failure_count
            ))
        for future in futures:
            try:
                future.result()
            except Exception as e:
                logger.error(f"Batch processing error: {str(e)}")

def start_mail_system(service):
    try:
        email_data = pd.read_csv('gmail.csv').to_dict('records')
        contacts_data = pd.read_csv('contacts.csv').to_dict('records')
        subjects = pd.read_csv('subjects.csv')['subject'].tolist()
        from_names = ["jonny smith", "john smith", "anny smith", "adam smith"]
        
        print(f"Starting email campaign with {len(contacts_data)} emails.")
        with open('body.txt', 'r') as f:
            body_content = f.read()
        
        success_count = []
        failure_count = []
        
        for i in range(0, len(contacts_data), BATCH_SIZE):
            process_batch(service, contacts_data[i:i + BATCH_SIZE], email_data, body_content, subjects, from_names, success_count, failure_count)
            time.sleep(2)
        
        print(f"\nEmail campaign completed! ✅ Sent: {len(success_count)}, ❌ Failed: {len(failure_count)}")
    except Exception as e:
        logger.error(f"Error during email campaign: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        service = create_gmail_service()
        start_mail_system(service)
    except KeyboardInterrupt:
        print("\nCode stopped by user")
    except Exception as e:
        logger.error(f"Fatal error: {str(e)}")
        raise
