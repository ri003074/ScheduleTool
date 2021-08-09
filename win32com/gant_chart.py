import calendar
from collections import defaultdict

import win32com.client
from icecream import ic

from variables import CALENDAR_CELL_WIDTH
from variables import xlHAlignCenter
from variables import xlToLeft
from variables import xlToRight
from variables import xlUp

""" information for naming
cells(row_index, column_index)
"""


def specify_cell_width_and_height(cells, cell_width, cell_height):
    cells.ColumnWidth = cell_width
    cells.RowHeight = cell_height


def calculate_consecutive_month_and_year(
    start_year, start_month, end_year, end_month
):
    arr = []
    current_year = start_year
    current_month = start_month
    arr.append([current_year, current_month])

    while True:
        if current_month == end_month and current_year == end_year:
            break
        elif current_month >= 12:
            current_month = 1
            current_year += 1
        else:
            current_month += 1

        arr.append([current_year, current_month])

    return arr


def convert_year_month_to_calendar(year, month):
    cal = calendar.Calendar()
    day_and_day_of_week = cal.monthdays2calendar(year=year, month=month)
    arr = []
    for data_per_week in day_and_day_of_week:
        for day, day_of_week in data_per_week:
            if day == 0:
                continue
            else:
                arr.append([day, get_day_of_week_from_num(day_of_week)])

    return arr


def get_day_of_week_from_num(num):
    data = [
        "月",
        "火",
        "水",
        "木",
        "金",
        "土",
        "日",
    ]
    return data[num]


class Gantt:
    def __init__(
        self,
        excel_file_name,
        start_year,
        start_month,
        end_year,
        end_month,
        check_items,
    ):
        self.excel_file_name = excel_file_name
        self.start_year = start_year
        self.start_month = start_month
        self.end_year = end_year
        self.end_month = end_month
        self.working_year_month_days = []
        self.check_items = check_items
        self.item_count = len(self.check_items)
        self.calendar_year_row_index = 0
        self.task_info_row_index = 6
        self.day_of_week_row_index = 5

        self.default_ws_name = "schedule"
        self.xl = win32com.client.GetObject(Class="Excel.Application")
        self.wb = self.xl.Workbooks(self.excel_file_name)
        self.ws = self.wb.Sheets(self.default_ws_name)

    def assign_task(self, task_info):
        for index, check_item in enumerate(task_info.values(), start=1):
            self.ws.Cells(self.task_info_row_index, index).Value = check_item
            self.ws.Cells(
                self.task_info_row_index, index
            ).HorizontalAlignment = xlHAlignCenter

        task_start_day_cell_column_index = (
            self.check_items.index("Start Day") + 1
        )
        start = f"{self.ws.Cells(self.task_info_row_index, task_start_day_cell_column_index).Value:%Y-%m-%d}"  # start = 2021-08-02
        start_year, start_month, start_day = map(int, start.split("-"))

        # clear calendar
        self.ws.Range(
            self.ws.Cells(self.task_info_row_index, self.item_count + 1),
            self.ws.Cells(self.task_info_row_index, self.ws.Columns.Count),
        ).Clear()

        cell_to_fill_column_index = self.find_cell_column_from_year_month_day(
            start_year,
            start_month,
            start_day,
            self.calendar_year_row_index,
            self.item_count + 1,
        )

        # task_info[2] == man-hour
        for _ in range(int(task_info.get("Man-hours"))):
            if (
                self.ws.Cells(
                    self.day_of_week_row_index, cell_to_fill_column_index
                ).Value
                == "土"
            ):
                cell_to_fill_column_index += 1
            if (
                self.ws.Cells(
                    self.day_of_week_row_index, cell_to_fill_column_index
                ).Value
                == "日"
            ):
                cell_to_fill_column_index += 1

            self.paint_cell(
                cell_row_index=self.task_info_row_index,
                cell_column_index=cell_to_fill_column_index,
                color=37,
            )
            cell_to_fill_column_index += 1

        else:
            cell_to_fill_column_index -= 1

        end_day_cell_column_index = self.check_items.index("End Day") + 1
        self.ws.Cells(
            self.task_info_row_index, end_day_cell_column_index
        ).Value = "-".join(
            map(
                str,
                self.working_year_month_days[
                    cell_to_fill_column_index - self.item_count - 1
                ],
            )
        )
        self.task_info_row_index += 1

    def add_items(self, start_row_index, start_column_index):
        self.ws.Cells.Clear()

        self.calendar_year_row_index = start_row_index - 3

        check_item_cells = self.ws.Range(
            self.ws.Cells(start_row_index, start_column_index),
            self.ws.Cells(
                start_row_index, start_column_index + len(self.check_items)
            ),
        )
        check_item_cells.Value = self.check_items
        check_item_cells.ColumnWidth = 11
        check_item_cells.HorizontalAlignment = xlHAlignCenter

        for index, item in enumerate(self.check_items):
            self.paint_cell(
                start_row_index, index + start_column_index, color=37
            )

    def add_calendar(self, start_row_index, start_column_index, year, month):
        # calender
        # 0-> Monday, 6-> Sunday
        days_and_weeks = convert_year_month_to_calendar(year, month)
        start_row_index = start_row_index
        start_column_index = start_column_index
        # year
        self.ws.Cells(start_row_index, start_column_index).Value = year
        self.ws.Cells(
            start_row_index, start_column_index
        ).HorizontalAlignment = xlHAlignCenter

        # month
        self.ws.Cells(start_row_index + 1, start_column_index).Value = month
        self.ws.Cells(
            start_row_index + 1, start_column_index
        ).HorizontalAlignment = xlHAlignCenter

        # day and day_of_week
        day_and_day_and_week_cells = self.ws.Range(
            self.ws.Cells(
                start_row_index + self.calendar_year_row_index,
                start_column_index,
            ),
            self.ws.Cells(
                start_row_index + self.calendar_year_row_index + 1,
                start_column_index + len(days_and_weeks) - 1,
            ),
        )
        days = [
            day_and_day_of_week[0] for day_and_day_of_week in days_and_weeks
        ]
        day_of_weeks = [
            day_and_day_of_week[1] for day_and_day_of_week in days_and_weeks
        ]

        day_and_day_and_week_cells.Value = [days, day_of_weeks]
        day_and_day_and_week_cells.HorizontalAlignment = xlHAlignCenter
        specify_cell_width_and_height(
            day_and_day_and_week_cells,
            cell_width=CALENDAR_CELL_WIDTH,
            cell_height=21,
        )

        for index, day_and_week in enumerate(days_and_weeks):
            self.working_year_month_days.append([year, month, day_and_week[0]])

    def add_calendars(self):
        working_month_and_days = calculate_consecutive_month_and_year(
            start_year=self.start_year,
            start_month=self.start_month,
            end_year=self.end_year,
            end_month=self.end_month,
        )
        print(working_month_and_days)
        start_column_index_offset = 1
        for working_month_and_day in working_month_and_days:
            year = working_month_and_day[0]
            month = working_month_and_day[1]
            self.add_calendar(
                start_row_index=self.calendar_year_row_index,
                start_column_index=self.item_count + start_column_index_offset,
                year=year,
                month=month,
            )
            start_column_index_offset += len(
                convert_year_month_to_calendar(year, month)
            )

    def find_cell_column_from_year_month_day(
        self,
        year,
        month,
        day,
        calendar_start_row_index,
        calendar_start_column_index,
    ):
        end_column = (
            self.ws.Cells(calendar_start_row_index, self.ws.Columns.Count)
            .End(xlToRight)
            .Column
        )

        cell_column = 0
        for cell_column_index in range(
            calendar_start_column_index, end_column
        ):
            if (
                self.ws.Cells(
                    calendar_start_row_index, cell_column_index
                ).Value
                == year
                and self.ws.Cells(
                    calendar_start_row_index + 1, cell_column_index
                ).Value
                == month
            ):
                for cell_column_index_day in range(
                    cell_column_index, end_column + 31
                ):
                    if (
                        self.ws.Cells(
                            calendar_start_row_index
                            + self.calendar_year_row_index,
                            cell_column_index_day,
                        ).Value
                        == day
                    ):
                        cell_column = cell_column_index_day
                        break
                else:
                    continue

                break

        return cell_column

    def specify_cell_width_and_height(
        self, row_index, column_index, cell_width, cell_height
    ):
        self.ws.Cells(row_index, column_index).ColumnWidth = cell_width
        self.ws.Cells(row_index, column_index).RowHeight = cell_height

    def paint_cell(self, cell_row_index, cell_column_index, color):
        self.ws.Cells(
            cell_row_index, cell_column_index
        ).Interior.ColorIndex = color

    def get_last_row(self, column_index):
        last_row_index = (
            self.ws.Cells(self.ws.Cells.Rows.Count, column_index).End(xlUp).Row
        )
        return last_row_index

    def get_last_column(self, row_index):
        last_column_index = (
            self.ws.Cells(row_index, self.ws.Cells.Columns.Count)
            .End(xlToLeft)
            .Column
        )
        return last_column_index

    def update_calendars(self):
        self.task_info_row_index = 6  # back to default
        last_row_index = self.get_last_row(1)

        for row_index in range(6, last_row_index + 1):
            check_items = defaultdict()
            for column_index in range(1, len(self.check_items) + 1):
                check_items[
                    self.check_items[column_index - 1]
                ] = self.ws.Cells(row_index, column_index).Value
            self.assign_task(check_items)

    def store_calendars(self):
        self.working_year_month_days = []
        start_row_index = 2
        start_column_index = self.item_count + 1
        self.calendar_year_row_index = 2

        last_data_column_index = self.get_last_column(start_row_index + 2)
        year = []
        month = []
        year_month_counter = 0
        for row in range(start_row_index, start_row_index + 3):
            for column in range(
                start_column_index, last_data_column_index + 1
            ):
                cell_value = self.ws.Cells(row, column).Value
                if cell_value is not None:
                    if row == start_row_index:
                        year.append(cell_value)
                    elif row == start_row_index + 1:
                        month.append(cell_value)
                    elif row == start_row_index + 2:

                        if (
                            column > start_column_index
                            and cell_value
                            < self.ws.Cells(row, column - 1).Value
                        ):
                            year_month_counter += 1
                        self.working_year_month_days.append(
                            [
                                int(year[year_month_counter]),
                                int(month[year_month_counter]),
                                int(cell_value),
                            ]
                        )

    def save_schedule(self):
        print(f"{1:02}")
        sheets_count = self.wb.Sheets.Count
        for sheet_number in range(sheets_count):
            ws_name = self.wb.Sheets[sheet_number].Name
            self.ws = self.wb.Sheets(ws_name)
            output_file_name = f"{sheet_number:02}" + "_" + ws_name + ".csv"
            ic(ws_name)
            ic(output_file_name)

            with open(output_file_name, "w") as f:
                last_row_index = self.get_last_row(1)
                for row in range(1, last_row_index + 1):
                    last_column_index = self.get_last_column(row)
                    for column in range(1, last_column_index + 1):
                        output_content = self.ws.Cells(row, column).Value
                        if output_content is not None:
                            f.write(str(output_content))
                        else:
                            f.write("")
                        f.write(",")
                    f.write("\n")

        self.ws = self.wb.Sheets(self.default_ws_name)
