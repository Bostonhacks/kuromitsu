import pandas as pd
import os
import time
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import argparse
from tqdm import tqdm
import pickle
from pathlib import Path

import base64
from email.message import EmailMessage

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


# Gmail API scopes
SCOPES = [
            'https://www.googleapis.com/auth/gmail.compose'
        ]

SENDER_NAME = "BostonHacks"

def read_data_file(file_path):
    """Read data from CSV or Excel file"""
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please use .csv, .xlsx, or .xls")

# https://developers.google.com/workspace/gmail/api/quickstart/python
def get_gmail_service():
    """Get authenticated Gmail API service"""
    creds = None
    token_path = Path('token.pickle')
    
    # Load credentials from file if it exists
    if token_path.exists():
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    
    # If credentials don't exist or are invalid, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Look for credentials.json in current directory
            credentials_path = Path('credentials.json')
            if not credentials_path.exists():
                raise FileNotFoundError(
                    "credentials.json not found. Download it from Google Cloud Console "
                    "and save it to the current directory."
                )
            
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)


def send_email(service, sender, recipient, subject, body, attachments=None):
    """Send an email using Gmail API with optional attachments"""
    msg = MIMEMultipart()
    msg['From'] = f"{SENDER_NAME} <{sender}>"
    msg['To'] = recipient
    msg['Subject'] = subject
    
    # Attach body text
    msg.attach(MIMEText(body, 'html'))
    
    # Attach files if provided
    if attachments:
        for file_path in attachments:
            if os.path.exists(file_path):
                with open(file_path, 'rb') as file:
                    part = MIMEApplication(file.read(), Name=os.path.basename(file_path))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                    msg.attach(part)
    
    # Encode the message
    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
    
    # Send the message
    service.users().messages().send(userId='me', body={'raw': raw_message}).execute()


def send_batch_emails(data_file, email_column, template_file=None, subject=None, 
                      attachments=None, test_mode=False, limit=None, delay=2):
    """Send batch emails using data from a file"""
    
    # Read the data file
    df = read_data_file(data_file)
    
    # Ensure email column exists
    if email_column not in df.columns:
        raise ValueError(f"Email column '{email_column}' not found in the data file")
    
    # Read email template if provided
    template = None
    if template_file:
        with open(template_file, 'r') as file:
            template = file.read()
    
    # Get Gmail API service
    print("Authenticating with Gmail API...")
    try:
        service = get_gmail_service()
        # Get sender email from OAuth profile
        profile = service.users().getProfile(userId='me').execute()
        sender = profile['emailAddress']
        
        if not sender.endswith('@gmail.com'):
            print(f"Warning: Authenticated as {sender}, not inbox.bostonhacks@gmail.com")
            confirm = input("Continue with this email? (yes/no): ").lower()
            if confirm != 'yes':
                print("Operation cancelled")
                return
    except Exception as e:
        print(f"Authentication failed: {str(e)}")
        return
    

    
    # Set default subject if not provided
    if not subject:
        subject = "Message from BostonHacks"
    
    # Determine email batch size
    emails_to_send = df.head(limit) if limit else df
    total = len(emails_to_send)
    
    print(f"Preparing to send {total} emails from {sender}")
    if test_mode:
        print("TEST MODE: Emails will not actually be sent")
    
    # Confirm before sending
    confirm = input(f"Send {total} emails? (yes/no): ").lower()
    if confirm != 'yes':
        print("Operation cancelled")
        return
    
    # Process each recipient
    success_count = 0
    for index, row in tqdm(emails_to_send.iterrows(), total=total, desc="Sending emails"):
        recipient = row[email_column]
        
        # Skip invalid emails
        if not isinstance(recipient, str) or '@' not in recipient:
            print(f"Skipping invalid email: {recipient}")
            continue
        
        # Personalize email if template has placeholders
        email_body = template
        if template:
            # Replace any column name in curly braces with the value from the row
            for col in df.columns:
                placeholder = f"{{{col}}}"
                if placeholder in template:
                    email_body = email_body.replace(placeholder, str(row[col]))
        else:
            email_body = f"Hello {recipient},\n\nThis is a message from BostonHacks."
        
        try:
            if not test_mode:
                send_email(service, sender, recipient, subject, email_body, attachments)
                success_count += 1
                # Add delay to avoid hitting send limits
                if index < total - 1:  # No need to delay after the last email
                    time.sleep(delay)
            else:
                # In test mode, just print what would be sent
                print(f"Would send to: {recipient}")
                success_count += 1
        except Exception as e:
            print(f"Error sending to {recipient}: {str(e)}")
    
    print(f"\nCompleted: {success_count} of {total} emails {'would be ' if test_mode else ''}sent successfully")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send batch emails from a CSV or Excel file")
    parser.add_argument("data_file", help="Path to CSV or Excel file with recipient data")
    parser.add_argument("email_column", help="Column name containing email addresses")
    parser.add_argument("--template", help="Path to HTML email template file")
    parser.add_argument("--subject", help="Email subject line")
    parser.add_argument("--attachments", nargs='+', help="Paths to files to attach")
    parser.add_argument("--test", action="store_true", help="Run in test mode without sending emails")
    parser.add_argument("--limit", type=int, help="Limit number of emails to send")
    parser.add_argument("--delay", type=float, default=2, help="Delay between emails in seconds (default: 2)")
    
    args = parser.parse_args()
    
    send_batch_emails(
        args.data_file, 
        args.email_column,
        args.template, 
        args.subject, 
        args.attachments,
        args.test,
        args.limit,
        args.delay
    )