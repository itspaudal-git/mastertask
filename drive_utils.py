from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

def get_google_drive_service():
    creds = Credentials.from_service_account_file('cred.json')
    service = build('drive', 'v3', credentials=creds)
    return service

def list_drive_contents(service, folder_id, drive_id):
    results = service.files().list(
        q=f"'{folder_id}' in parents",
        pageSize=1000,
        fields="nextPageToken, files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        driveId=drive_id,
        corpora='drive'
    ).execute()
    items = results.get('files', [])
    
    while 'nextPageToken' in results:
        page_token = results['nextPageToken']
        results = service.files().list(
            q=f"'{folder_id}' in parents",
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            driveId=drive_id,
            corpora='drive'
        ).execute()
        items.extend(results.get('files', []))
    
    return items

def find_sheet_id(term, parent_folder_id, drive_id):
    drive_service = get_google_drive_service()
    items = list_drive_contents(drive_service, parent_folder_id, drive_id)
    
    term_folder_id = None
    term_name = f"Term {term}".strip().lower()
    for item in items:
        if item['mimeType'] == 'application/vnd.google-apps.folder' and term_name in item['name'].strip().lower():
            term_folder_id = item['id']
            break
    
    if not term_folder_id:
        return None
    
    items = list_drive_contents(drive_service, term_folder_id, drive_id)
    
    sheet_id = None
    for item in items:
        if item['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            sheet_id = item['id']
            break
    
    return sheet_id
