# Python script to interface with Google Sheets

import json
import google.oauth2.credentials # pip3 install google-auth
from googleapiclient.discovery import build # pip3 install google-api-python-client

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

sheets_api = None

'''
Function to initialise the Google Sheets API.
'''
def init_sheets_api():
    global sheets_api

    if sheets_api is None:
        debug_msg('Initialising Google Sheets API')

        # Load credentials
        with open('google_tokens.json', 'r') as tokenfile:
            google_tokens = json.load(tokenfile)
        creds = google.oauth2.credentials.Credentials(
            None,
            refresh_token=google_tokens['refresh_token'],
            token_uri=google_tokens['token_uri'],
            client_id=google_tokens['client_id'],
            client_secret=google_tokens['client_secret']
        )

        # Create the Google Sheets API
        sheets_service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheets_api = sheets_service.spreadsheets()


'''Function to append a row to a Google Sheet.
document_id: ID of the spreadsheet document as in the URL
sheet_name: Name of the sheet as in the tabs
rows: List of rows, with each row being a list of values
'''
def append_rows(document_id, sheet_name, rows):
    init_sheets_api()
    request_body = {
        "range": sheet_name,
        "majorDimension": 'ROWS',
        'values': rows
    }
    request = sheets_api.values().append(
        spreadsheetId=document_id,
        range=sheet_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=request_body)
    return request.execute()

'''
Function to read a column from a Google Sheet.
document_id: ID of the spreadsheet document as in the URL
sheet_name: Name of the sheet as in the tabs
column_code: Which column to read, i.e. 'A'
offset: Which row to start from
'''
def read_column(document_id, sheet_name, column_code, offset):
    init_sheets_api()
    sheet_range = f'{sheet_name}!{column_code}{offset}:{column_code}'
    result = sheets_api.values().get(
        spreadsheetId=document_id,
        range=sheet_range).execute()
    return result.get('values', [])

'''
Function to read cells from a Google Sheet.
document_id: ID of the spreadsheet document as in the URL
sheet_name: Name of the sheet as in the tabs
start_column_code: Start column to read (like 'A')
start_column_offset: Which row to start from (like 3)
end_column_code: End column to read (like 'B')
end_column_offset: Which row to end (like 5)
'''
def read_cells(document_id, sheet_name, start_column_code, start_column_offset, end_column_code, end_column_offset=None):
    init_sheets_api()
    sheet_range = f'{sheet_name}!{start_column_code}{start_column_offset}:{end_column_code}'
    if end_column_offset:
        sheet_range += f'{end_column_offset}'
    result = sheets_api.values().get(
        spreadsheetId=document_id,
        range=sheet_range).execute()
    return result.get('values', [])

'''
Function to update cell value to a Google Sheet.
document_id: ID of the spreadsheet document as in the URL
sheet_name: Name of the sheet as in the tabs
column_code: column to update (like 'A')
row_offset: offset to update (like 5)
value: Value to be updated (like 'Completed)
'''
def update_cell(document_id, sheet_name, column_code, row_offset, value):
    init_sheets_api()

    values = [[value]]
    request_body = {
        "majorDimension": 'ROWS',
        'values': values
    }

    sheet_range = f'{sheet_name}!{column_code}{row_offset}:{column_code}{row_offset}'
    request = sheets_api.values().update(
        spreadsheetId=document_id,
        range=sheet_range,
        valueInputOption='USER_ENTERED',
        body=request_body)

    return request.execute()

'''
Prints verbose debug message.
'''
def debug_msg(msg):
    if g_EnableDebugMsg:
        print(msg)
