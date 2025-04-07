# Kuromitsu
A CLI tool for mass mailing templated emails. By default, it adds a 20 second delay in between batches of 50 sent emails to avoid hitting Gmail API's rate limit. You change change this delay or batch size with `--delay` and `--batch-size` respectively. 

Currently all emails are sent with name "BostonHacks" (even though you can change the send-from email address). 


# Getting started
Requirements: Python 3

1) Create a virtual env (optional but recommended)
```bash
python -m venv .venv
source ./.venv/bin/Activate # on mac/linux
.\.venv\Scripts\Activate.ps1 # on Windows
```
Ensure you're using the virtual env by invoking `pip -V`. Check the path to be within your current directory

2) Install requirements
   1) `pip install -r requirements.txt`
3) Copy `credentials.json` with either your own Google OAuth Client credentials or provided to you by a tech team member.
4) Add a .csv file to the `/data` folder with data of the emails and other fields you wish to use to send the email.
   1) You can check `/data/example.csv` for an example. You can add/remove columns as needed.
5) Create a HTMl template (or use the example) that you wish to use to send as an email. 
   1) Anytime the wildcard `{ variable }` is used, you must add that as a column to a csv file. So if { name } is used in the HTMl template, add a `name` column to the .csv file
6) Invoke mailer.py with the following 
```bash
usage: mailer.py [-h] [--send-as SEND_AS] [--reply-to REPLY_TO]
                 [--template TEMPLATE] [--subject SUBJECT]
                 [--attachments ATTACHMENTS [ATTACHMENTS ...]] [--test]
                 [--limit LIMIT] [--delay DELAY] [--workers WORKERS]
                 [--batch-size BATCH_SIZE]
                 data_file email_column

Send batch emails from a CSV or Excel file

positional arguments:
  data_file             Path to CSV or Excel file with recipient data
  email_column          Column name containing email addresses

options:
  -h, --help            show this help message and exit
  --send-as SEND_AS     Send-as email address. Also serves as reply-to unless
                        specified otherwise. Must be a valid alias on gmail account
  --reply-to REPLY_TO   Reply-To email address
  --template TEMPLATE   Path to HTML email template file
  --subject SUBJECT     Email subject line
  --attachments ATTACHMENTS [ATTACHMENTS ...]
                        Paths to files to attach
  --test                Run in test mode without sending emails
  --limit LIMIT         Limit number of emails to send
  --delay DELAY         Delay between emails in seconds (default: 20)
  --workers WORKERS     Number of concurrent workers (default: 20)
  --batch-size BATCH_SIZE
                        Number of emails to send in each batch (default: 50)
```
An example
```bash
python mailer.py <path/to/data.csv> <email_column_name> --send-as "contact@bostonhacks.org" --subject "BostonHacks 2025 Updates" --template <path/to/template.html> --attachments <path/to/file> <path/to/file2> --delay 10
```
This will open the data.csv with all the emails and their associated information, read the email column, set the subject of the email, use the HTMl template specified, and attach the files. Emails will be sent in batches with a delay of 10 seconds in between them to avoid rate limits.

**You might have to set a longer delay depending on the usage of the Gmail API. If sending many emails, either lower the batch size or increase the delay between batches to avoid hitting usage limits**.

7) Login via Google. You must use `inbox.bostonhacks@gmail.com` or `contact.bostonhacks@gmail.com` if using the credentials provided by the Tech Team. 
   1) Ignore the "Unverified app" warning. If you are using this externally, you should be using your own Google client anyway.
8) Any subsequent invocation of the script will automatically use the account you logged in. To "logout", delete the "token.pickle" local file in the same directory.

## Arguments
Run `python mailer.py -h` to get help page

