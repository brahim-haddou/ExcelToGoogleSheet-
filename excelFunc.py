import xlrd
import win32com.client as win32
import os


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
    wb.Close(True)
    return sheets, data


def list_files(path):
    d_path = os.path.normpath(os.path.expanduser(path))
    f_names = []
    f_dirs = []
    for f_name in os.listdir(d_path):
        if f_name.endswith(".xlsx"):
            f_dir = os.path.join(d_path, f_name)
            f_names.append(f_name.split(".")[0])
            f_dirs.append(f_dir)
    return f_names, f_dirs
