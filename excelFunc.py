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


if __name__ == "__main__":
    excel_file = r"C:\Users\Flex5\Desktop\Test\test al.xlsx"
    import win32com.client as win32
    from pprint import pprint
    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    wb = excel.Workbooks.Open(excel_file)
    
    ws = wb.Worksheets("Sheet1")
    colors = []
    for i in range(1, 10):
        clr = []
        for j in range(1, 10):
            ws.Cells(i, j).Value = i
            clr.append(ws.Cells(i, j).Interior.ColorIndex)
            print(ws.Cells(i, j).NumberFormat)
            print(ws.Cells(i, j).Borders.Value)
            print(ws.Cells(i, j).Borders.Color)
            print(ws.Cells(i, j).Font.Name)
            print(ws.Cells(i, j).Font.Size)
            print(ws.Cells(i, j).Font.Color)
            print(ws.Cells(i, j).Borders.LineStyle)
            print()
        colors.append(clr)
    pprint(colors)
