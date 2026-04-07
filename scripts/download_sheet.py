"""
Download the Google Sheet as an Excel file using a service account.
"""

import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SPREADSHEET_ID = '10p6B4A-PiXYisHUvz3G3qIX4Uzlz_RYO'
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), '..', 'data', 'cohorts_data_preprocessed.xlsx')


def main():
    # Load credentials from environment variable or file
    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_KEY')
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        raise RuntimeError('GOOGLE_SERVICE_ACCOUNT_KEY environment variable not set')

    # Use Drive API to download the file
    # This file is an uploaded .xlsx (not a native Google Sheet), so use get_media
    drive_service = build('drive', 'v3', credentials=creds)

    request = drive_service.files().get_media(fileId=SPREADSHEET_ID)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"Download progress: {int(status.progress() * 100)}%")

    with open(OUTPUT_PATH, 'wb') as f:
        f.write(fh.getvalue())

    print(f"Downloaded sheet to: {OUTPUT_PATH}")


if __name__ == '__main__':
    main()
