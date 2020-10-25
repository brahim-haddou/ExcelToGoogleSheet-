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
    excel_file = r"C:\Users\Flex5\Desktop\Test\test alll.xlsx"
    import win32com.client as win32
    from pprint import pprint
    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    wb = excel.Workbooks.Open(excel_file)
    
    ws = wb.Worksheets("Sheet1")
    colors = []
    Number_Format = []
    Borders_V = []
    Borders_C = []
    Borders_LineStyle = []
    Font_Name = []
    Font_Size = []
    Font_Color = []
    for i in range(1, 10):
        clr = []
        Num_Format = []
        Bor_V = []
        Bor_C = []
        Bor_LineStyle = []
        Font_N = []
        Font_S = []
        Font_C = []
        
        for j in range(1, 10):
            clr.append(ws.Cells(i, j).Interior.ColorIndex)
            Num_Format.append(ws.Cells(i, j).NumberFormat)
            Bor_V.append(ws.Cells(i, j).Borders.Value)
            Bor_C.append(ws.Cells(i, j).Borders.Color)
            Bor_LineStyle.append(ws.Cells(i, j).Borders.LineStyle)
            Font_N.append(ws.Cells(i, j).Font.Name)
            Font_S.append(ws.Cells(i, j).Font.Size)
            Font_C.append(ws.Cells(i, j).Font.Color)
            
        colors.append(clr)
        Font_Color.append(Font_C)
        Number_Format.append(Num_Format)
        Borders_V.append(Bor_V)
        Borders_C.append(Bor_C)
        Borders_LineStyle.append(Bor_LineStyle)
        Font_Name.append(Font_N)
        Font_Size.append(Font_S)
        
    pprint(colors)
    pprint('')
    pprint(Font_Color)
    pprint('')
    pprint(Number_Format)
    pprint('')
    pprint(Borders_V)
    pprint('')
    pprint(Borders_C)
    pprint('')
    pprint(Borders_LineStyle)
    pprint('')
    pprint(Font_Name)
    pprint('')
    pprint(Font_Size)
    pprint('')
    
    del ws
    del excel
