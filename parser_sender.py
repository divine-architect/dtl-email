import os
import pickle
import re
import pandas as pd
from base64 import urlsafe_b64encode
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError

# gapi setting
SCOPES = ['https://mail.google.com/']
OUR_EMAIL = '<my college email, removed for privacy>'

# regex i hate regex
email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

def gmail_authenticate():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh()
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def build_message(recipient, subject, body_html):
    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = OUR_EMAIL
    message['To'] = recipient
    message.attach(MIMEText(body_html, 'html'))
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_messages(service, messages):
    sent_count = 0
    for idx, message in enumerate(messages, start=1):
        try:
            service.users().messages().send(userId="me", body=message).execute()
            sent_count += 1
            print(f"Sent email {sent_count}/{len(messages)} successfully.")
        except HttpError as error:
            print(f"Failed to send email {sent_count}: {error}")
            continue  # Skip to the next message if there's an error

if __name__ == "__main__":
    # do not ask how I obtained this
    file_path = 'names.xlsx'
    all_sheets = pd.read_excel(file_path, sheet_name='Form responses 1')  # Load only the specified sheet

    # custom html template to send msg
    with open("sender.html", "r") as file:
        body_template = file.read()

    
    service = gmail_authenticate()
    
    
    subject = "Please fill our Design Thinking Lab (DTL) form"
    messages = []
    
   # more parsing
    for index, row in all_sheets.iterrows():
        name = row['STUDENT NAME ']
        email = row['Student Mail Id(RVCE Mail ID)']
        
        
        if not re.match(email_pattern, str(email)):
            print(f"Invalid email format for {name}: {email}. Skipping.")
            continue
        
        
        body_html = body_template.replace("{name}", name)
        message = build_message(email, subject, body_html)
        messages.append(message)
    
    
    send_messages(service, messages)
    print("Script completed successfully!")
