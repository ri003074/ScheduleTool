import win32com.client


def sample():
    xl = win32com.client.GetObject(Class="Excel.Application")
    wb = xl.Workbooks("demo_01.xlsx")
    ws = wb.Sheets(1)
    ws.Range("A1").Value = 123


if __name__ == "__main__":
    sample()
