from functions import create_service
from excelFunc import excel_file, list_files

CLIENT_SECRET_FILE = 'code_secret_client.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file']

service = create_service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
print(service.permissions())


def create_spreadsheet(title):
    import datetime
    tz_string = datetime.datetime.now(datetime.timezone.utc).astimezone().tzname()
    spreadsheet = {
        'properties': {
            'title': title,
            'locale': 'en_US',
            'timeZone': tz_string.split()[0],
            'autoRecalc': 'HOUR'
        },
        'sheets': {
            'properties': {
                "sheetId": 0,
                "title": 'sheet'
            }
        }
    }
    response_c = service.spreadsheets().create(body=spreadsheet).execute()
    return response_c


def add_sheets(gsheet_id, sheet_name):
    try:
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                    }
                }
            }]
        }
        
        response_s = service.spreadsheets().batchUpdate(
            spreadsheetId=gsheet_id,
            body=request_body
        ).execute()
        
        return response_s
    except Exception as e:
        print(e)


def update_spreadsheet(gsheet_id, sheet_name, data):
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': data
    }
    
    response_u = service.spreadsheets().values().update(
        spreadsheetId=gsheet_id,
        valueInputOption='USER_ENTERED',
        range=sheet_name,
        body=value_range_body
    ).execute()
    return response_u


def delete_sheet(gsheet_id):
    try:
        request_body = {
            'requests': [{
                'deleteSheet': {
                    'sheetId': 0
                }
            }]
        }
        
        response_s = service.spreadsheets().batchUpdate(
            spreadsheetId=gsheet_id,
            body=request_body
        ).execute()
        
        return response_s
    except Exception as e:
        print(e)


"""
if __name__ == "__main__":
    path = input()
    f_names, f_dirs = list_files(path)
    for j in range(len(f_names)):
        response = create_spreadsheet(f_names[j])
        spreadsheetId = response['spreadsheetId']
        Sheets, values = excel_file(f_dirs[j])
        for i in range(len(Sheets)):
            add_sheets(spreadsheetId, Sheets[i])
            update_spreadsheet(spreadsheetId, Sheets[i], values[i])
        delete_sheet(spreadsheetId)
"""
