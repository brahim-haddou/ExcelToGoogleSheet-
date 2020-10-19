import xlrd
import os


def excel_file(path):
    xls = xlrd.open_workbook(path, on_demand=True)
    sheets = xls.sheet_names()
    data = []
    for s in xls.sheets():
        values = []
        for row in range(s.nrows):
            col_value = []
            for col in range(s.ncols):
                value = s.cell(row, col).value
                try:
                    value = str(int(value))
                except:
                    pass
                col_value.append(value)
            values.append(col_value)
        data.append(values)
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
