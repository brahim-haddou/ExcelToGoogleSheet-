from functions import create_service
from excelFunc import excel_file, list_files

CLIENT_SECRET_FILE = 'code_secret_client.json.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = create_service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
print(service)


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


def add_sheets(g_sheet_id, sheet_name):
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
            spreadsheetId=g_sheet_id,
            body=request_body
        ).execute()
        
        return response_s
    except Exception as e:
        print(e)


def update_spreadsheet(g_sheet_id, sheet_name, data):
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': data
    }
    
    response_u = service.spreadsheets().values().update(
        spreadsheetId=g_sheet_id,
        valueInputOption='USER_ENTERED',
        range=sheet_name,
        body=value_range_body
    ).execute()
    return response_u


def delete_sheet(g_sheet_id):
    try:
        request_body = {
            'requests': [{
                'deleteSheet': {
                    'sheetId': 0
                }
            }]
        }
        
        response_s = service.spreadsheets().batchUpdate(
            spreadsheetId=g_sheet_id,
            body=request_body
        ).execute()
        
        return response_s
    except Exception as e:
        print(e)


def update_cell(g_sheet_id, sheet_id, value, sheet_BG_color, sheet_F_name, sheet_F_size, sheet_F_color):
    try:
        request_body = {
            'requests': []
        }
        for l in range(len(sheet_F_name)):
            for k in range(len(sheet_F_name[0])):
                request = {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "stringValue": str(value[l][k]),
                                        },
                                        "userEnteredFormat": {
                                            "backgroundColor": {
                                                "red": sheet_BG_color[l][k][0] / 256,
                                                "green": sheet_BG_color[l][k][1] / 256,
                                                "blue": sheet_BG_color[l][k][2] / 256
                                            },
                                            "textFormat": {
                                                "foregroundColor": {
                                                    "red": sheet_F_color[l][k][0] / 256,
                                                    "green": sheet_F_color[l][k][1] / 256,
                                                    "blue": sheet_F_color[l][k][2] / 256
                                                },
                                                "fontFamily": sheet_F_name[l][k],
                                                "fontSize": sheet_F_size[l][k]
                                            }
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "*",
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": l,
                            "endRowIndex": l + 1,
                            "startColumnIndex": k,
                            "endColumnIndex": k + 1
                        }
                    }
                }
                request_body["requests"].append(request)
        response_s = service.spreadsheets().batchUpdate(
            spreadsheetId=g_sheet_id,
            body=request_body
        ).execute()
        return response_s
    except Exception as e:
        print(e)


if __name__ == "__main__":
    path = input("enter the path : ")
    fo_name = path.split('\\')[-1:][0]
    f_names, f_dirs = list_files(path)
    file_metadata = {
        'name': 'Invoices',
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=fo_name,
                                  fields='id').execute()
    # for j in range(len(f_names)):
    #     response = create_spreadsheet(f_names[j])
    #     spreadsheetId = response['spreadsheetId']
    #     Sheets, values, sheets_BG_color, sheets_F_name, sheets_F_size, sheets_F_color = excel_file(f_dirs[j])
    #     for i in range(len(Sheets)):
    #         response_sheet = add_sheets(spreadsheetId, Sheets[i])
    #         w_sheet_id = response_sheet['replies'][0]['addSheet']['properties']['sheetId']
    #         if values[i]:
    #             update_cell(spreadsheetId, w_sheet_id, values[i], sheets_BG_color[i], sheets_F_name[i],
    #                         sheets_F_size[i], sheets_F_color[i])
    #     delete_sheet(spreadsheetId)
