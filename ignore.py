from googleapiclient.discovery import build
from google.oauth2 import service_account

# This function fetches data from a Google Sheet and returns it as a list.
def get_sheet_data_as_list(sheet_id, range_name, service_account_file='cred.json'):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    try:
        # Fetch the data from the sheet
        result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get('values', [])

        # Flatten the list if it's a list of lists
        if values and isinstance(values[0], list):
            values = [item for sublist in values for item in sublist]

        return values

    except Exception as e:
        print(f"An error occurred: {e}")
        return []
