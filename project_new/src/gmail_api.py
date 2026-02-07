from __future__ import annotations

import base64
from pathlib import Path
from email.mime.text import MIMEText

from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
# src/drive_api.py の2階層上が project/
PROJECT_DIR = Path(__file__).resolve().parents[1]
CREDENTIALS_PATH = PROJECT_DIR / "config" / "credentials.json"
TOKEN_PATH = PROJECT_DIR / "config" / "ttoken_gmail.json"


def get_gmail_service():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def send_gmail_api(to: str, subject: str, body: str) -> None:
    service = get_gmail_service()

    msg = MIMEText(body, "plain", "utf-8")
    msg["To"] = to
    msg["Subject"] = subject

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print("✅ Gmail APIでメール送信完了")
