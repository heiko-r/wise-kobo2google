# Python script to upload a file to Google Drive

import json
import google.oauth2.credentials # pip3 install google-auth
from googleapiclient.discovery import build # pip3 install google-api-python-client
from googleapiclient.http import MediaIoBaseUpload

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

GOOGLE_DRIVE_FOLDER_ID = '1zzWDH0Iltzp6cU9jqacF9MmLR3NT2SSx'

'''
Function to upload a file to Google Drive
'''
def upload_file(filename, fh):
    # Set up Google API authentication
    with open('google_tokens.json', 'r') as tokenfile:
        google_tokens = json.load(tokenfile)
    creds = google.oauth2.credentials.Credentials(
        None,
        refresh_token=google_tokens['refresh_token'],
        token_uri=google_tokens['token_uri'],
        client_id=google_tokens['client_id'],
        client_secret=google_tokens['client_secret']
    )
    drive_service = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': filename,
        'parents': [GOOGLE_DRIVE_FOLDER_ID]
    }
    media = MediaIoBaseUpload(fh, mimetype='application/pdf')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
    debug_msg(f'File ID: { file.get("id") }')
    return file

'''
Prints verbose debug message.
'''
def debug_msg(msg):
    if g_EnableDebugMsg:
        print(msg)
