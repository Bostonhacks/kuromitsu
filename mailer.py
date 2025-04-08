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
import concurrent.futures
import threading

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
            'https://www.googleapis.com/auth/gmail.compose',
            "https://www.googleapis.com/auth/gmail.settings.basic",
            "https://www.googleapis.com/auth/gmail.readonly"
        ]

SENDER_NAME = "BostonHacks"

# local storage for gmail services
thread_local = threading.local()

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

def get_thread_gmail_service():
    """Get Gmail API service for current thread"""
    if not hasattr(thread_local, 'gmail_service'):
        # print("Creating new Gmail service for thread")
        thread_local.gmail_service = get_gmail_service()
    return thread_local.gmail_service

def send_email(service, sender, recipient, subject, body, attachments=None, reply_to=None):
    """Send an email using Gmail API with optional attachments"""
    msg = MIMEMultipart()
    msg['From'] = f"{SENDER_NAME} <{sender}>"
    msg['To'] = recipient
    msg['Subject'] = subject

    # Set Reply-To header if provided
    if reply_to:
        msg["Reply-To"] = reply_to
    
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
    try:

        response = service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        message_id = response.get("id")

        if message_id:
            # # done with threading so can wait for response
            # time.sleep(1)

            # message = service.users().messages().get(userId="me", id=message_id).execute()
            # labels = message.get("labelIds", [])

            # print(labels)

            # if "SENT" in labels:
            #     return True, f"Email sent successfully to {recipient}"
            # else:
            #     return False, f"Email to {recipient} was created but not confirmed as sent"

            return True, f"Email sent successfully to {recipient}"
            # print(response)

        else:
            return False, f"No message ID when sending to {recipient}"
    
    

        # message_id = response.get("id")

        # if message_id:
        #     message = service.users().messages().get(userId='me', id=message_id).execute()

        #     labels = message.get("labelIds", [])
        #     if "SENT" not in labels:
        #         raise Exception("Email not sent successfully")
        #     return True, f"Email sent successfully to {recipient} with message ID: {message_id}"

    except HttpError as error:
        # Handle specific Gmail API errors
        if error.resp.status == 400:
            # This could be a malformed email address or other request issue
            return False, f"Invalid request: {str(error)}"
        elif error.resp.status == 403:
            return False, f"Permission denied: {str(error)}"
        elif error.resp.status == 404:
            return False, f"Resource not found: {str(error)}"
        elif error.resp.status == 429:
            return False, "Rate limit exceeded. Try again later."
        else:
            return False, f"Gmail API error: {str(error)}"
    except Exception as e:
        # Handle other exceptions
        return False, f"An error occurred: {str(e)}"

        


def process_email(args):
    """Process single email for concurrent execution"""
    index, row, email_column, template, df_columns, sender, subject, attachments, reply_to, test_mode = args

    recipient = row[email_column]

    # Skip invalid emails
    if not isinstance(recipient, str) or '@' not in recipient:
        print(f"Skipping invalid email: {recipient}")
        return (index, False, f"Invalid email: {recipient}")
    
    # Personalize email if template has placeholders
    email_body = template
    if template:
        # Replace any column name in curly braces with the value from the row
        for col in df_columns:
            placeholder = f"{{{col}}}"
            if placeholder in template:
                email_body = email_body.replace(placeholder, str(row[col]))
    else:
        email_body = f"Hello {recipient},\n\nThis is a message from BostonHacks."

    try:
        if not test_mode:
            service = get_thread_gmail_service()
            success, message = send_email(service=service,
                       sender=sender,
                       recipient=recipient,
                       subject=subject,
                       body=email_body,
                       attachments=attachments,
                       reply_to=reply_to
                       )
            
            if success: 
                return (index, True, recipient)
            else:
                return (index, False, message)

        else:
            return (index, True, f"Would send to: {recipient}")
    except Exception as e:
        return (index, False, f"Error sending to {recipient}: {str(e)}")



def send_batch_emails(data_file, email_column, template_file=None, subject=None, reply_to=None, send_as=None,
                      attachments=None, test_mode=False, limit=None, delay=6, max_workers=10, batch_size=10):
    """Send batch emails using data from a file"""
    
    # Read the data file
    df = read_data_file(data_file)

    # create output directory
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    timestamp = time.strftime("%Y%m%d-%H%M%S")
    output_subdir = Path("output", f"output_{timestamp}")
    output_subdir.mkdir(exist_ok=True)
    
    output_file_all = output_subdir / f"output.csv"
    output_file_failure = output_subdir / f"output_failure.csv"
    
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
        # print(profile)
        sender = profile['emailAddress']
        
        if not sender.endswith('@gmail.com'):
            print(f"Warning: Authenticated as {sender}, not inbox.bostonhacks@gmail.com")
            confirm = input("Continue with this email? (yes/no): ").lower()
            if confirm != 'yes':
                print("Operation cancelled")
                return None, None
            
        # if using an alias, check if alias is valid
        if send_as:
            addresses = service.users().settings().sendAs().list(userId="me").execute()
            # print(addresses)
            send_as_addresses = addresses.get("sendAs", [])

            valid_aliases = [sendAs["sendAsEmail"] for sendAs in send_as_addresses]

            if send_as not in valid_aliases:
                print(f"{send_as} is not a valid alias to send as")
                print(f"Valid send-as addresses: {", ".join(valid_aliases)}")
                return None, None
            
            sender = send_as
        
    except Exception as e:
        print(f"Authentication failed: {str(e)}")
        return None, None
    

    
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
    failure_count = 0

    df_columns = df.columns.tolist()

    # results df
    results_df = emails_to_send.copy()
    results_df["success"] = False
    results_df["timestamp"] = None
    results_df["message"] = ""
    results_df["subject"] = subject

    # arg list for each email
    email_tasks = [
        (index, row, email_column, template, df_columns, sender, subject, attachments, reply_to, test_mode)
        for index, row in emails_to_send.iterrows()
    ]

    """threaded version"""
    batches = [email_tasks[i:i + batch_size] for i in range(0, len(email_tasks), batch_size)]

    # just for debugging
    # for batch in batches:
    #     print(len(batch))

    # return None, None

    pbar = tqdm(total=total, desc="Sending emails")

    for batch in batches:
        # execute email tasks in batch with threadpool
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(process_email, args) for args in batch]

            for future in concurrent.futures.as_completed(futures):
                index, success, message = future.result()
                if success:
                    success_count += 1
                    results_df.at[index, "success"] = True
                    results_df.at[index, "message"] = message
                else:
                    failure_count += 1
                    results_df.at[index, "success"] = False
                    results_df.at[index, "message"] = message
                    print(f"Error sending to {index}: {message}")
                pbar.update(1)

                results_df.at[index, "timestamp"] = time.strftime("%Y-%m-%d %H:%M:%S")

        # Add delay to avoid hitting limits
        if len(batches) > 1:
            time.sleep(delay)

    pbar.close()
    print(f"\nCompleted: {success_count} of {total} emails {'would be ' if test_mode else ''}sent successfully")
    if failure_count > 0:
        print(f"Failed: {failure_count} emails")

    # save results to CSV
    results_df.to_csv(output_file_all, index=False)
    results_df[~results_df["success"]].to_csv(output_file_failure, index=False)
    return output_file_all, output_file_failure

    
    """synchronous version"""
    # for index, row in tqdm(emails_to_send.iterrows(), total=total, desc="Sending emails"):
    #     recipient = row[email_column]
        
    #     # Skip invalid emails
    #     if not isinstance(recipient, str) or '@' not in recipient:
    #         print(f"Skipping invalid email: {recipient}")
    #         continue
        
    #     # Personalize email if template has placeholders
    #     email_body = template
    #     if template:
    #         # Replace any column name in curly braces with the value from the row
    #         for col in df.columns:
    #             placeholder = f"{{{col}}}"
    #             if placeholder in template:
    #                 email_body = email_body.replace(placeholder, str(row[col]))
    #     else:
    #         email_body = f"Hello {recipient},\n\nThis is a message from BostonHacks."
        
    #     try:
    #         if not test_mode:
    #             send_email(service, sender, recipient, subject, email_body, attachments)
    #             success_count += 1
    #             # Add delay to avoid hitting send limits
    #             if index < total - 1:  # No need to delay after the last email
    #                 time.sleep(delay)
    #         else:
    #             # In test mode, just print what would be sent
    #             print(f"Would send to: {recipient}")
    #             success_count += 1
    #     except Exception as e:
    #         print(f"Error sending to {recipient}: {str(e)}")

    # print(f"\nCompleted: {success_count} of {total} emails {'would be ' if test_mode else ''}sent successfully")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send batch emails from a CSV or Excel file")
    parser.add_argument("data_file", help="Path to CSV or Excel file with recipient data")
    parser.add_argument("email_column", help="Column name containing email addresses")
    parser.add_argument("--send-as", help="Send-as email address. Also serves as reply-to unless specified otherwise. Must be a valid alias on gmail account")
    parser.add_argument("--reply-to", help="Reply-To email address")
    parser.add_argument("--template", help="Path to HTML email template file")
    parser.add_argument("--subject", help="Email subject line")
    parser.add_argument("--attachments", nargs='+', help="Paths to files to attach")
    parser.add_argument("--test", action="store_true", help="Run in test mode without sending emails")
    parser.add_argument("--limit", type=int, help="Limit number of emails to send")
    parser.add_argument("--delay", type=float, default=6, help="Delay between emails in seconds (default: 6)")
    parser.add_argument("--workers", type=int, default=10, help="Number of concurrent workers (default: 10)")
    parser.add_argument("--batch-size", type=int, default=10, help="Number of emails to send in each batch (default: 10)") 

    args = parser.parse_args()
    
    output_all, output_errors = send_batch_emails(
        data_file=args.data_file,
        email_column=args.email_column,
        template_file=args.template,
        subject=args.subject,
        reply_to=args.reply_to,
        send_as=args.send_as,
        attachments=args.attachments,

        test_mode=args.test,
        limit=args.limit,
        delay=args.delay,
        max_workers=args.workers,
        batch_size=args.batch_size
    )

    if output_all or output_errors:
        print(f"Output files containing success/failure emails: {output_all}")
        print(f"Output files containing failed emails: {output_errors}")