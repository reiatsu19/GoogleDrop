# GoogleDrop: Email Attachment to Google Drive
# This script connects to a Gmail inbox, filters recent emails based on subject and date, 
# downloads PDF attachments (such as invoices), and uploads them to Google Drive using Python. 
# It demonstrates a practical cloud automation workflow combining email parsing, 
# file filtering, and cloud storage integration via Google Drive API.

# Email handling
import imaplib
import email
import getpass
import os
import glob
from email import policy
from email.parser import BytesParser

# Google Drive API (official v3)
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Setup login credentials
gmail_user = "anildev7111@gmail.com"
gmail_pass = getpass.getpass("Enter your app password: ")

# Connecting to gmail IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Login using IMAP
imap.login(gmail_user, gmail_pass)

# Make local folder for file
save_folder = "invoices"
os.makedirs(save_folder, exist_ok=True)

# Selecting folder to filter through
imap.select("INBOX")

# Defining specific email cc address
to_address = "anildev7111@gmail.com"
subject_keyword = "Invoice"
since_date = "05-Jun-2025"  # Format must be DD-MMM-YYYY for IMAP

# Set search criteria to TO address
search_criteria = f'(TO "{to_address}" SUBJECT "{subject_keyword}" SINCE "{since_date}")'

# Launch search
status, messages = imap.search(None, search_criteria)

# Decode results
emails = messages[0].split()

# Track how many have PDFs
pdf_count = 0
saved_count = 0
skipped_count = 0

# Loop through all matching emails
for email_id in emails:
    # Fetch the raw email
    res, msg_data = imap.fetch(email_id, "(RFC822)")
    if res != 'OK':
        print(f"Failed to fetch email ID {email_id}")
        continue

    # Parse email
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email, policy=policy.default)

    # Optional: Show subject for tracking
    # print(f"\n📧 Subject: {msg['Subject']}")

    # Walk through the email parts to find attachments
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            filename = part.get_filename()
            if filename and filename.lower().endswith(".pdf"):
                filepath = os.path.join(save_folder, filename)
                if os.path.exists(filepath):
                    skipped_count += 1
                else:
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                        saved_count += 1
                        pdf_count += 1

# After loop finishes
print(f"Summary:")
print(f"PDFs saved = {saved_count}")
print(f"PDFs skipped = {skipped_count}")
print(f"Emails processed = {len(emails)}")

# Connects the script to Google Drive using the official Google Drive API v3.
# After authenticating the Google account, it searches for (or creates) an “Invoices” folder in Google Drive. 
# The script then scans the local invoices folder for all PDF files downloaded from Gmail. 
# Each file is checked against existing files in the Drive folder to avoid duplicates, 
# and only new PDFs are uploaded to Google Drive.

# Authenticate and build Drive API service
auth.authenticate_user()
drive_service = build('drive', 'v3')

# Create or get the "Invoices" folder in Google Drive
folder_name = "Invoices"
# Search for the folder in Drive (ignore trashed files)
query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
items = results.get('files', [])

# Use existing folder if found, otherwise create a new one
if items:
    folder_id = items[0]['id']
    print(f"Found existing folder: {folder_name} (ID: {folder_id})")
else:
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = drive_service.files().create(body=file_metadata, fields='id').execute()
    folder_id = folder.get('id')
    print(f"Created new folder: {folder_name} (ID: {folder_id})")

# Upload PDFs to Google Drive without duplicates
upload_folder = "invoices"
pdf_files = glob.glob(f"{upload_folder}/*.pdf")
uploaded_count = 0
skipped_count = 0

for file_path in pdf_files:
    file_name = os.path.basename(file_path)

    # Check if file already exists in target Drive folder
    check_query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
    check_results = drive_service.files().list(q=check_query, spaces='drive', fields='files(id)').execute()

    if check_results.get('files'):
        print(f"Skipped (duplicate found): {file_name}")
        skipped_count += 1
        continue

    # Upload file to the folder
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path, mimetype='application/pdf')
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    print(f"Uploaded: {file_name} (ID: {uploaded_file.get('id')})")
    uploaded_count += 1

print(f"\nUploaded PDFs: {uploaded_count}")
print(f"Skipped duplicates: {skipped_count}")
