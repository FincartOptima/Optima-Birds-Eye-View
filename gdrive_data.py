"""Fetch the Tradebook / Client File / Account Statement from a shared Google
Drive folder, so the webpage can show live data without anyone manually
uploading files. Whoever manages the data just replaces the same-named file
in the Drive folder; the site picks up the new version on its next refresh.

Auth: a Google Cloud service account. The folder must be shared with the
service account's email as Viewer for the trade-file reads above — but the
Mailing tab's client-email storage (client_emails.json, in this same
folder) needs the service account to be able to WRITE too, which needs the
folder shared as Editor (or Content Manager) instead of Viewer. Same
service account either way, just a higher role on the same folder.

Credentials, in order of precedence:
  1. GDRIVE_SERVICE_ACCOUNT_JSON env var — the full JSON key content (used on
     Render, where secrets are set as env vars rather than files).
  2. GDRIVE_SERVICE_ACCOUNT_FILE env var — a path to the JSON key file.
  3. credentials/gdrive_service_account.json next to this file (local dev).

Folder ID, in order of precedence:
  1. GDRIVE_FOLDER_ID env var.
  2. The folder ID this integration was first set up against (fallback).
"""
from __future__ import annotations

import io
import json
import os
import time
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# Full (not read-only) scope — needed so the client-email JSON file can be
# written once the folder is shared as Editor. Requesting the broader scope
# is harmless even while the folder is still Viewer-only; the actual write
# still gets rejected at the Drive permission layer until that's bumped.
SCOPES = ['https://www.googleapis.com/auth/drive']
_DEFAULT_KEY_PATH = Path(__file__).resolve().parent / "credentials" / "gdrive_service_account.json"
_DEFAULT_FOLDER_ID = "1ofUZU0HeEkq3F-A1O_WtBrufYouARPc8"

# Fixed file names expected inside the Drive folder (matched case-insensitively,
# ignoring any extension the file may or may not have).
FILE_ALIASES = {
    "tradebook": "tradebook",
    "client_file": "client file",
    "account_statement": "account statement",
}
_EXT_BY_MIME = {
    "text/csv": ".csv",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
}

_CACHE_TTL_SECONDS = 600  # 10 minutes — data doesn't change intraday
_cache: dict[str, tuple[float, dict]] = {}


def _folder_id() -> str:
    return os.environ.get("GDRIVE_FOLDER_ID", _DEFAULT_FOLDER_ID)


def _load_credentials() -> service_account.Credentials:
    json_env = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if json_env:
        info = json.loads(json_env)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    key_path = Path(os.environ.get("GDRIVE_SERVICE_ACCOUNT_FILE", str(_DEFAULT_KEY_PATH)))
    if not key_path.exists():
        raise ValueError(
            f"Google Drive credentials not found at {key_path}. Set GDRIVE_SERVICE_ACCOUNT_JSON "
            "(key content) or GDRIVE_SERVICE_ACCOUNT_FILE (key path) as an environment variable."
        )
    return service_account.Credentials.from_service_account_file(str(key_path), scopes=SCOPES)


def is_configured() -> bool:
    try:
        _load_credentials()
        return True
    except Exception:
        return False


def _get_service():
    creds = _load_credentials()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _list_folder_files(service) -> list[dict]:
    results = service.files().list(
        q=f"'{_folder_id()}' in parents and trashed = false",
        fields="files(id, name, mimeType, modifiedTime)",
    ).execute()
    return results.get("files", [])


def _download_bytes(service, file_id: str) -> bytes:
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def fetch_files(force: bool = False) -> dict[str, tuple[bytes, str]]:
    """Return {alias: (bytes, filename)} for whichever of the 3 expected files
    are present in the Drive folder. alias is one of FILE_ALIASES' keys.
    Cached for _CACHE_TTL_SECONDS unless force=True."""
    now = time.time()
    cached = _cache.get("_all")
    if not force and cached and (now - cached[0]) < _CACHE_TTL_SECONDS:
        return cached[1]

    service = _get_service()
    drive_files = _list_folder_files(service)

    result: dict[str, tuple[bytes, str]] = {}
    for alias, match_name in FILE_ALIASES.items():
        found = next((f for f in drive_files if f["name"].strip().lower() == match_name), None)
        if not found:
            continue
        content = _download_bytes(service, found["id"])
        ext = _EXT_BY_MIME.get(found["mimeType"], "")
        name = found["name"]
        filename = name if name.lower().endswith(ext) or not ext else name + ext
        result[alias] = (content, filename)

    _cache["_all"] = (now, result)
    return result


def read_json_file(filename: str) -> dict:
    """Read a JSON file by exact name from the shared folder. Returns {} if
    the file doesn't exist yet (e.g. before the first save)."""
    service = _get_service()
    drive_files = _list_folder_files(service)
    found = next((f for f in drive_files if f["name"].strip().lower() == filename.strip().lower()), None)
    if not found:
        return {}
    content = _download_bytes(service, found["id"])
    return json.loads(content.decode("utf-8")) if content.strip() else {}


def write_json_file(filename: str, data: dict) -> None:
    """Create or update a JSON file by exact name in the shared folder.
    Requires the folder to be shared with the service account as Editor
    (not just Viewer) — raises a clear error otherwise."""
    service = _get_service()
    drive_files = _list_folder_files(service)
    found = next((f for f in drive_files if f["name"].strip().lower() == filename.strip().lower()), None)

    payload = json.dumps(data, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype="application/json")

    try:
        if found:
            service.files().update(fileId=found["id"], media_body=media).execute()
        else:
            service.files().create(
                body={"name": filename, "parents": [_folder_id()]},
                media_body=media,
                fields="id",
            ).execute()
    except Exception as e:
        if "insufficientParentPermissions" in str(e) or "insufficient permissions" in str(e).lower():
            raise ValueError(
                "The Drive folder is shared with the service account as Viewer, which allows reading "
                "the trade files but not writing this file. Share the folder with the service account "
                "as Editor (or Content Manager) to enable saving."
            ) from e
        if "storageQuotaExceeded" in str(e) or "storage quota" in str(e).lower():
            raise ValueError(
                f"'{filename}' doesn't exist in the Drive folder yet, and the service account has no "
                "storage quota of its own to create a brand-new file (a Google Drive limitation). "
                f"Fix: in the Drive folder, manually create an empty file named exactly '{filename}' "
                "containing just {} — once it exists, this same save will update its contents instead "
                "of creating it, which works fine under the service account's Editor access."
            ) from e
        raise
