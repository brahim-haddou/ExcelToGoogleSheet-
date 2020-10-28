import xlrd
import os
import win32com.client as win32


def excel_file(path):
    xls = xlrd.open_workbook(path, on_demand=True)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    wb = excel.Workbooks.Open(path)
    sheets = xls.sheet_names()
    sheets_data = []
    sheets_BG_color = []
    sheets_F_name = []
    sheets_F_size = []
    sheets_F_color = []
    
    for s in xls.sheets():
        ws = wb.Worksheets(s.name)
        values = []
        colors = []
        Font_Name = []
        Font_Size = []
        Font_Color = []
        for row in range(s.nrows):
            col_value = []
            clr = []
            Font_N = []
            Font_S = []
            Font_C = []
            for col in range(s.ncols):
                value = s.cell(row, col).value
                C = ws.Cells(row + 1, col + 1).Interior.Color
                R = C % 256
                G = C // 256 % 256
                B = C // 65536 % 256
                c = [R, G, B]
                clr.append(c)
                Font_N.append(ws.Cells(row + 1, col + 1).Font.Name)
                Font_S.append(ws.Cells(row + 1, col + 1).Font.Size)
                C = ws.Cells(row + 1, col + 1).Font.Color
                R = C % 256
                G = C // 256 % 256
                B = C // 65536 % 256
                c = [R, G, B]
                Font_C.append(c)
                try:
                    value = str(int(value))
                except:
                    pass
                col_value.append(value)
            
            values.append(col_value)
            colors.append(clr)
            Font_Name.append(Font_N)
            Font_Size.append(Font_S)
            Font_Color.append(Font_C)
        
        sheets_data.append(values)
        sheets_BG_color.append(colors)
        sheets_F_name.append(Font_Name)
        sheets_F_size.append(Font_Size)
        sheets_F_color.append(Font_Color)
    wb.Close(True)
    excel.Quit()
    del xls
    return sheets, sheets_data, sheets_BG_color, sheets_F_name, sheets_F_size, sheets_F_color


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
    from pprint import pprint
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(r'C:\Users\Flex5\Desktop\Test\test1.xlsx')
    ws = wb.Worksheets("Sheet1")
    pprint(ws.Cells(1, 1).Borders(1).LineStyle)
    pprint(ws.Cells(1, 1).Borders(1).Weight)
    pprint(ws.Cells(1, 1).Borders(1).Color)
    wb.Close(True)
    excel.Quit()
