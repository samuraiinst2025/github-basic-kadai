from __future__ import annotations

import io
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


# 読み書きするので drive スコープ（最も簡単）
SCOPES = ["https://www.googleapis.com/auth/drive"]

# src/drive_api.py の2階層上が project/
PROJECT_DIR = Path(__file__).resolve().parents[1]
CREDENTIALS_PATH = PROJECT_DIR / "config" / "credentials.json"
TOKEN_PATH = PROJECT_DIR / "config" / "token_drive.json"


def get_drive_service():
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

    return build("drive", "v3", credentials=creds)


def _download_request_to_path(request, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with io.FileIO(out_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return out_path


def _resolve_shortcut(file_id: str) -> Tuple[str, str]:
    """
    ショートカットなら実体IDとmimeTypeを返す。ショートカットでなければ元のまま返す。
    """
    service = get_drive_service()
    meta = service.files().get(
        fileId=file_id,
        fields="id,mimeType,shortcutDetails(targetId,targetMimeType)",
    ).execute()

    mime = meta.get("mimeType")
    if mime == "application/vnd.google-apps.shortcut":
        target_id = meta["shortcutDetails"]["targetId"]
        target_mime = meta["shortcutDetails"].get("targetMimeType") or ""
        # 念のため target の mimeType を取り直す
        meta2 = service.files().get(fileId=target_id, fields="id,mimeType").execute()
        return target_id, meta2.get("mimeType", target_mime)

    return file_id, mime


def download_file(file_id: str, out_path: Path) -> Path:
    """
    Drive fileId -> ローカル保存
    - バイナリ（xlsx/json/pdf等）：get_media
    - Google Docs Editors（Sheets/Docs/Slides）：export_media
    - Shortcut：実体へ解決してから判定
    """
    service = get_drive_service()

    resolved_id, mime = _resolve_shortcut(file_id)

    if mime == "application/vnd.google-apps.folder":
        raise ValueError(f"このIDはフォルダです（ファイルIDを指定してください）: {file_id}")

    export_map = {
        "application/vnd.google-apps.spreadsheet": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xlsx",
        ),
        "application/vnd.google-apps.document": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".docx",
        ),
        "application/vnd.google-apps.presentation": (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pptx",
        ),
    }

    if mime in export_map:
        export_mime, ext = export_map[mime]
        request = service.files().export_media(fileId=resolved_id, mimeType=export_mime)
        if out_path.suffix.lower() != ext:
            out_path = out_path.with_suffix(ext)
        return _download_request_to_path(request, out_path)

    request = service.files().get_media(fileId=resolved_id)
    return _download_request_to_path(request, out_path)


def upload_overwrite(file_id: str, local_path: Path, mime_type: str) -> None:
    """
    ローカルファイル -> Driveの同じfileIdへ上書き
    """
    service = get_drive_service()
    media = MediaFileUpload(str(local_path), mimetype=mime_type, resumable=True)
    service.files().update(fileId=file_id, media_body=media).execute()


def list_files_in_folder(folder_id: str) -> List[Dict]:
    """
    フォルダ配下のファイル一覧（ショートカット含む）
    """
    service = get_drive_service()
    q = f"'{folder_id}' in parents and trashed = false"

    files: List[Dict] = []
    page_token = None
    while True:
        resp = service.files().list(
            q=q,
            fields="nextPageToken, files(id,name,mimeType,modifiedTime,shortcutDetails(targetId,targetMimeType))",
            pageToken=page_token,
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return files
