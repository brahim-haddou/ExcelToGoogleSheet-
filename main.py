import os
from functions import Create_Service


CLIENT_SECRET_FILE = 'code_secret_client.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

# Spreadsheet file
print("-------> creation of spreadsheet")

########################
def create_spreadsheet():
    spreadsheet = {
        'properties': {
            'title': 'First google sheet file',
            'locale': 'en_US',
            'timeZone': 'Casablanca',
            'autoRecalc': 'HOUR'
        },
        'sheets':[
            {
            'properties': {
                    'title': 'sheet 1'
                }, 
            },
            {
            'properties': {
                    'title': 'sheet 2'
                }, 
            },
            {
            'properties': {
                    'title': 'sheet 3'
                }, 
            },
            {
            'properties': {
                    'title': 'sheet 4'
                }, 
            }
        ]
    }

    spreadsheet = service.spreadsheets().create(body=spreadsheet).execute()
    print(spreadsheet)


########################
def update_spreadsheet():
    spreadsheet_id = '1UYhC84kcSwmS2qtFqdWrXF3MnJa9DtSMzw8Cirstavo'
    my_spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    worksheet_name = 'sheet!'
    cell_range_insert = 'B2'
    values = (
        ('col A','col B','col C','col D','col E'),
        ('A1','B1','C1','D1','E1')
    )
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': values
    }

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name+cell_range_insert,
        body=value_range_body
    ).execute()

########################
def clear_spreadsheet():
    spreadsheet_id = '1UYhC84kcSwmS2qtFqdWrXF3MnJa9DtSMzw8Cirstavo'
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range='sheet'
    ).execute()
########################
def append_to_spreadsheet():
    spreadsheet_id = '1UYhC84kcSwmS2qtFqdWrXF3MnJa9DtSMzw8Cirstavo'
    my_spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    worksheet_name = 'sheet!'
    cell_range_insert = 'B2'
    values = (
        ('col A','col B','col C','col D','col E'),
        ('A1','B1','C1','D1','E1')
    )
    value_range_body = {
        'majorDimension': 'COLUMNS',
        'values': values
    }

    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name+cell_range_insert,
        body=value_range_body
    ).execute()

if __name__ == "__main__":
    append_to_spreadsheet()
    