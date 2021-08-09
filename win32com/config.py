from collections import defaultdict

import win32com.client

from variables import xlToLeft
from variables import xlUp


class Info:
    def __init__(self, excel_file_name):
        self.excel_file_name = excel_file_name

        xl = win32com.client.GetObject(Class="Excel.Application")
        wb = xl.Workbooks(self.excel_file_name)
        self.ws = wb.Sheets("config")

        last_row_index = (
            self.ws.Cells(self.ws.Cells.Rows.Count, 1).End(xlUp).Row
        )

        self.config_information = defaultdict(list)
        for row in range(1, last_row_index + 1):
            last_column_index = (
                self.ws.Cells(row, self.ws.Columns.Count).End(xlToLeft).Column
            )
            for column in range(2, last_column_index + 1):
                self.config_information[self.ws.Cells(row, 1).Value].append(
                    self.ws.Cells(row, column).Value
                )
