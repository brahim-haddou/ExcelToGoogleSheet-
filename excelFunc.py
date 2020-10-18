import xlrd
import win32com.client as win32


def excel_file(path):
    xls = xlrd.open_workbook(path, on_demand=True)
    sheets = xls.sheet_names()
    del xls
    
    xlApp = win32.Dispatch('Excel.Application')
    wb = xlApp.Workbooks.Open(path)
    data = []
    for sheet in sheets:
        ws = wb.Worksheets(sheet)
        rngData = ws.Range('A1').CurrentRegion()
        data.append(rngData)
    print(sheets, data)
    wb.Close(True)
    return sheets, data
