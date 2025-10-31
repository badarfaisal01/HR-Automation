import os
import re
import base64
import pandas as pd
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from PyPDF2 import PdfReader
import docx2txt

# ============================
# CONFIG
# ============================
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '11Gj6Z_U2_uj9SuPSkQuzx1g3Nc80QLYFzynzoaXI1jo'  # <-- your sheet ID
SHEET_NAME = 'Form Responses 1'
SHEET_RANGE = f'{SHEET_NAME}!A:G'  # matches columns A-G

DOWNLOAD_FOLDER = 'resumes'

EXPECTED_SKILLS = {'Python', 'Java', 'SQL', 'Machine Learning', 'AI', 'Data Analysis', 'Communication'}

# ============================
# AUTHENTICATION
# ============================
def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


# ============================
# GMAIL HANDLER
# ============================
def fetch_and_download_cvs(service):
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    results = service.users().messages().list(userId='me', q="has:attachment").execute()
    messages = results.get('messages', [])
    files = []

    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
        parts = msg_data['payload'].get('parts', [])

        for part in parts:
            filename = part.get('filename')
            if not filename:
                continue
            if 'attachmentId' in part['body']:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(
                    userId='me', messageId=msg['id'], id=att_id).execute()
                data = base64.urlsafe_b64decode(att['data'])
                filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                with open(filepath, 'wb') as f:
                    f.write(data)
                files.append(filepath)
                print(f"📄 Downloaded: {filename}")
    return files


# ============================
# CV PARSER
# ============================
def extract_text_from_file(filepath):
    text = ""
    if filepath.lower().endswith('.pdf'):
        reader = PdfReader(filepath)
        for page in reader.pages:
            text += page.extract_text() or ""
    elif filepath.lower().endswith('.docx'):
        text = docx2txt.process(filepath)
    return text


def parse_cv_data(text):
    # Extract name
    name_match = re.search(r'([A-Z][a-z]+(?:\s[A-Z][a-z]+)*)', text)
    name = name_match.group(1) if name_match else "Unknown"

    # Extract email
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    email = email_match.group(0) if email_match else "Not Found"

    # Extract experience
    exp_match = re.search(r'(\d+)\s*(?:years|yrs)', text, re.IGNORECASE)
    experience = exp_match.group(1) + " years" if exp_match else "Not Mentioned"

    # Try detecting field/role
    field_match = re.search(r'(developer|engineer|analyst|manager|designer)', text, re.IGNORECASE)
    field = field_match.group(1).capitalize() if field_match else "General"

    # Extract skills
    found_skills = {skill for skill in EXPECTED_SKILLS if skill.lower() in text.lower()}
    missing_skills = EXPECTED_SKILLS - found_skills
    unexpected_skills = set(re.findall(r'\b[A-Z][a-zA-Z]+\b', text)) - EXPECTED_SKILLS

    return {
        'Name': name,
        'Email': email,
        'Field': field,
        'Experience': experience,
        'Missing Skills': ', '.join(missing_skills),
        'Unexpected Skills': ', '.join(list(unexpected_skills)[:5])  # limit for clarity
    }


# ============================
# GOOGLE SHEETS UPLOADER
# ============================
def upload_to_google_sheet(service, data):
    values = [[
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Timestamp
        d['Name'],
        d['Email'],
        d['Field'],
        d['Experience'],
        d['Missing Skills'],
        d['Unexpected Skills']
    ] for d in data]

    body = {'values': values}
    sheet_service = service.spreadsheets().values()
    sheet_service.append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()

    print(f"✅ Uploaded {len(values)} candidates to Google Sheets!")


# ============================
# MAIN
# ============================
if __name__ == '__main__':
    creds = authenticate()

    # Gmail
    gmail_service = build('gmail', 'v1', credentials=creds)
    files = fetch_and_download_cvs(gmail_service)

    if not files:
        print("No CVs found in Gmail attachments.")
        exit()

    # Parse CVs
    parsed_data = []
    for file in files:
        text = extract_text_from_file(file)
        parsed_data.append(parse_cv_data(text))

    # Save locally
    df = pd.DataFrame(parsed_data)
    df.to_excel('parsed_cv_data.xlsx', index=False)
    print("💾 Saved extracted CV data to parsed_cv_data.xlsx")

    # Upload to Google Sheets
    sheet_service = build('sheets', 'v4', credentials=creds)
    upload_to_google_sheet(sheet_service, parsed_data)
