import win32com.client as wc
def create_table():
    xl = wc.GetObject(Class="Excel.Application")
    wb = xl.Workbooks("demo_01.xlsx")
    ws = wb.Sheets(1)
    ws.Range("A1").Value="No"
