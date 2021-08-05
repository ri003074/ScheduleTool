import win32com.client
import sys



def sample():
    xl = win32com.client.GetObject(Class="Excel.Application")
    wb = xl.Workbooks("demo_01.xlsx")
    ws = wb.Sheets(1)
    ws.Range(" A1").Value = 123


if __name__ == "__main__":
    sys.dont_write_bytecode = 1
    sample()
