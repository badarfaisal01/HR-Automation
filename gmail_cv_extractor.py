from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
import os
import email
import pandas as pd

# === CONFIG ===
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_FILE = 'credentials.json'   # your OAuth credentials
DOWNLOAD_FOLDER = 'resumes'             # folder to save CVs

# Create download folder if not exists
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

# === AUTHENTICATE ===
def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

# === FETCH EMAILS ===
def fetch_emails(service):
    results = service.users().messages().list(userId='me', q="has:attachment").execute()
    messages = results.get('messages', [])
    if not messages:
        print("No emails with attachments found.")
        return []

    print(f"Found {len(messages)} emails with attachments.")
    return messages

# === DOWNLOAD ATTACHMENTS ===
def download_attachments(service, messages):
    data = []
    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
        payload = msg_data['payload']
        headers = payload.get('headers', [])
        subject = sender = date = ''
        for h in headers:
            if h['name'] == 'Subject':
                subject = h['value']
            if h['name'] == 'From':
                sender = h['value']
            if h['name'] == 'Date':
                date = h['value']

        parts = payload.get('parts', [])
        for part in parts:
            filename = part.get('filename')
            if filename and ('data' in part['body'] or 'attachmentId' in part['body']):
                attachment_id = part['body'].get('attachmentId')
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=msg['id'], id=attachment_id
                ).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'])
                path = os.path.join(DOWNLOAD_FOLDER, filename)
                with open(path, 'wb') as f:
                    f.write(file_data)
                print(f"Downloaded: {filename}")
                data.append({'Sender': sender, 'Subject': subject, 'Date': date, 'File': filename})
    return data

# === MAIN ===
if __name__ == '__main__':
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)
    messages = fetch_emails(service)
    if messages:
        records = download_attachments(service, messages)
        if records:
            df = pd.DataFrame(records)
            df.to_excel('email_attachments.xlsx', index=False)
            print("✅ Email attachments metadata saved to email_attachments.xlsx")
