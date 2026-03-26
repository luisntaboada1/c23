import mimetypes
import re
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from config import (
    CREDS_PATH,
    PDFS_TO_VALIDATE_AS_PUBLIC,
    REQUIRED_FILES,
    SCOPES,
    TOKEN_PATH,
)

DRIVE_FOLDER_MIME_TYPE = "application/vnd.google-apps.folder"


def get_drive_service():
    creds = None

    print("credentials.json path:", CREDS_PATH)
    print("credentials.json exists:", CREDS_PATH.exists())
    print("token.json path:", TOKEN_PATH)
    print("token.json exists:", TOKEN_PATH.exists())

    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def extract_drive_folder_id(folder_link: str) -> str:
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", folder_link)
    if match:
        return match.group(1)

    parsed = urlparse(folder_link)
    query = parse_qs(parsed.query)
    if "id" in query and query["id"]:
        return query["id"][0]

    raise ValueError("Could not extract a folder ID from the provided Google Drive link.")


def build_drive_folder_link(folder_id: str) -> str:
    return f"https://drive.google.com/drive/folders/{folder_id}"


def _list_permissions(service, file_id: str) -> list[dict]:
    permissions = []
    page_token = None

    while True:
        response = service.permissions().list(
            fileId=file_id,
            fields="nextPageToken, permissions(id,type,role,allowFileDiscovery)",
            pageToken=page_token,
            supportsAllDrives=True,
        ).execute()

        permissions.extend(response.get("permissions", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return permissions


def is_public_for_anyone(service, file_id: str) -> bool:
    permissions = _list_permissions(service, file_id)
    return any(
        permission.get("type") == "anyone"
        and permission.get("role") in {"reader", "commenter", "writer"}
        for permission in permissions
    )


def make_file_public(service, file_id: str):
    if is_public_for_anyone(service, file_id):
        return

    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
        supportsAllDrives=True,
    ).execute()


def get_file_links(service, file_id: str) -> dict:
    file_obj = service.files().get(
        fileId=file_id,
        fields="id,name,webViewLink,webContentLink",
        supportsAllDrives=True,
    ).execute()

    return {
        "id": file_obj["id"],
        "name": file_obj["name"],
        "webViewLink": file_obj.get("webViewLink"),
        "webContentLink": file_obj.get("webContentLink"),
        "link": file_obj.get("webViewLink") or file_obj.get("webContentLink"),
    }


def _find_file_in_folder_by_name(service, folder_id: str, filename: str) -> list[dict]:
    safe_name = filename.replace("'", "\\'")
    query = f"'{folder_id}' in parents and trashed = false and name = '{safe_name}'"

    response = service.files().list(
        q=query,
        fields="files(id,name,parents)",
        pageSize=50,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    return response.get("files", [])


def upload_file_to_folder(
    service,
    folder_id: str,
    local_file_path: str | Path,
    *,
    make_public: bool = False,
) -> dict:
    local_file_path = Path(local_file_path)
    if not local_file_path.exists():
        raise FileNotFoundError(f"Local file not found: {local_file_path}")

    mime_type = mimetypes.guess_type(local_file_path.name)[0] or "application/octet-stream"
    media = MediaFileUpload(str(local_file_path), mimetype=mime_type, resumable=False)

    existing_files = _find_file_in_folder_by_name(service, folder_id, local_file_path.name)

    if len(existing_files) > 1:
        raise ValueError(
            "More than one file with the same generated name already exists in the folder: "
            f"{local_file_path.name}. Clean them up before running again."
        )

    if existing_files:
        uploaded = service.files().update(
            fileId=existing_files[0]["id"],
            media_body=media,
            fields="id,name",
            supportsAllDrives=True,
        ).execute()
    else:
        uploaded = service.files().create(
            body={"name": local_file_path.name, "parents": [folder_id]},
            media_body=media,
            fields="id,name",
            supportsAllDrives=True,
        ).execute()

    if make_public:
        make_file_public(service, uploaded["id"])

    return get_file_links(service, uploaded["id"])


def download_file_from_drive(service, file_id: str, destination_path: str | Path) -> str:
    destination_path = Path(destination_path)
    destination_path.parent.mkdir(parents=True, exist_ok=True)

    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    file_buffer = BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    with open(destination_path, "wb") as output_file:
        output_file.write(file_buffer.getvalue())

    return str(destination_path)


def validate_required_folder_files(service, folder_link: str) -> dict | None:
    folder_id = extract_drive_folder_id(folder_link)
    folder_name = get_drive_folder_name(service, folder_id)

    names_query = " or ".join([f"name = '{filename}'" for filename in REQUIRED_FILES])
    query = f"'{folder_id}' in parents and trashed = false and ({names_query})"

    found_files = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            fields="nextPageToken, files(id,name,mimeType,webViewLink,webContentLink)",
            pageSize=100,
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        found_files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    files_by_name = {}
    for file_obj in found_files:
        files_by_name.setdefault(file_obj["name"], []).append(file_obj)

    duplicates = [name for name, items in files_by_name.items() if len(items) > 1]
    if duplicates:
        print("Warning: duplicate filenames found in the folder:")
        for name in duplicates:
            print(f" - {name}")
        print("Stopping because the result would be ambiguous.")
        return None

    missing = [name for name in REQUIRED_FILES if name not in files_by_name]
    if missing:
        print("Missing required files:")
        for name in missing:
            print(f" - {name}")
        return None

    not_public = []
    for pdf_name in PDFS_TO_VALIDATE_AS_PUBLIC:
        file_obj = files_by_name[pdf_name][0]
        if not is_public_for_anyone(service, file_obj["id"]):
            not_public.append(pdf_name)

    if not_public:
        print("The following files are not public for anyone:")
        for name in not_public:
            print(f" - {name}")
        print("Please make them public and try again.")
        return None

    links = {}
    ids = {}

    for drive_name, alias in REQUIRED_FILES.items():
        file_obj = files_by_name[drive_name][0]
        links[alias] = file_obj.get("webViewLink") or file_obj.get("webContentLink")
        ids[alias] = file_obj["id"]

    return {
        "folder_id": folder_id,
        "folder_name": folder_name,
        "links": links,
        "ids": ids,
    }


def get_drive_folder_name(service, folder_id: str) -> str:
    folder = service.files().get(
        fileId=folder_id,
        fields="id,name,mimeType",
        supportsAllDrives=True,
    ).execute()

    return folder["name"]


def list_child_folders(service, parent_folder_link: str) -> list[dict]:
    parent_folder_id = extract_drive_folder_id(parent_folder_link)
    query = (
        f"'{parent_folder_id}' in parents and trashed = false "
        f"and mimeType = '{DRIVE_FOLDER_MIME_TYPE}'"
    )

    child_folders = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            fields="nextPageToken, files(id,name)",
            pageSize=100,
            pageToken=page_token,
            orderBy="name_natural",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        child_folders.extend(response.get("files", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return [
        {
            "id": folder["id"],
            "name": folder["name"],
            "link": build_drive_folder_link(folder["id"]),
        }
        for folder in child_folders
    ]
