import calendar

import win32com.client
from icecream import ic

from variables import CALENDAR_CELL_WIDTH
from variables import END_MONTH
from variables import END_YEAR
from variables import EXCEL_FILE_NAME
from variables import ITEMS
from variables import START_MONTH
from variables import START_YEAR
from variables import xlHAlignCenter
from variables import xlToRight


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
        self, excel_file_name, start_year, start_month, end_year, end_month
    ):
        self.excel_file_name = excel_file_name
        self.start_year = start_year
        self.start_month = start_month
        self.end_year = end_year
        self.end_month = end_month
        self.working_year_month_days = []
        self.item_count = 0
        self.calendar_reference_row_index = 0

        xl = win32com.client.GetObject(Class="Excel.Application")
        wb = xl.Workbooks(self.excel_file_name)
        self.ws = wb.Sheets(1)
        self.ws.Cells.Clear()

    def assign_task(self, row_index, name, man_hour, start_date):
        self.ws.Cells(row_index, 5).Value = name
        self.ws.Cells(row_index, 5).HorizontalAlignment = xlHAlignCenter
        self.ws.Cells(row_index, 6).Value = man_hour
        self.ws.Cells(row_index, 6).HorizontalAlignment = xlHAlignCenter
        self.ws.Cells(row_index, 7).Value = start_date
        self.ws.Cells(row_index, 7).HorizontalAlignment = xlHAlignCenter

        start = f"{self.ws.Cells(row_index, 7).Value:%Y-%m-%d}"
        start_year, start_month, start_day = map(int, start.split("-"))

        cell_column_index_start = self.find_cell_column_from_year_month_day(
            start_year,
            start_month,
            start_day,
            self.calendar_reference_row_index,
            self.item_count + 1,
        )

        for _ in range(man_hour):
            if self.ws.Cells(5, cell_column_index_start).Value == "土":
                cell_column_index_start += 1
            if self.ws.Cells(5, cell_column_index_start).Value == "日":
                cell_column_index_start += 1

            self.paint_cell(
                cell_row_index=row_index,
                cell_column_index=cell_column_index_start,
                color=37,
            )
            cell_column_index_start += 1

        else:
            cell_column_index_start -= 1

        self.ws.Cells(row_index, 8).Value = "-".join(
            map(
                str,
                self.working_year_month_days[
                    cell_column_index_start - self.item_count - 1
                ],
            )
        )
        self.ws.Cells(row_index, 8).HorizontalAlignment = xlHAlignCenter

    def add_items(self, start_row_index, start_column_index, items):
        self.item_count = len(items)
        self.calendar_reference_row_index = start_row_index - 3

        for index, item in enumerate(items):
            self.ws.Cells(
                start_row_index, index + start_column_index
            ).Value = item
            self.ws.Cells(
                start_row_index, index + start_column_index
            ).ColumnWidth = 10
            self.ws.Cells(
                start_row_index, index + start_column_index
            ).HorizontalAlignment = xlHAlignCenter

            self.paint_cell(start_row_index, index + start_column_index, 37)

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
        for index, day_and_week in enumerate(days_and_weeks):
            self.working_year_month_days.append([year, month, day_and_week[0]])
            self.ws.Cells(
                start_row_index + self.calendar_reference_row_index,
                index + start_column_index,
            ).Value = day_and_week[0]

            self.specify_cell_width_and_height(
                row_index=start_row_index + self.calendar_reference_row_index,
                column_index=index + start_column_index,
                cell_width=CALENDAR_CELL_WIDTH,
                cell_height=21,
            )
            self.ws.Cells(
                start_row_index + self.calendar_reference_row_index,
                index + start_column_index,
            ).HorizontalAlignment = xlHAlignCenter
            self.ws.Cells(
                start_row_index + 3, index + start_column_index
            ).Value = day_and_week[1]
            self.ws.Cells(
                start_row_index + 3, index + start_column_index
            ).HorizontalAlignment = xlHAlignCenter

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
                start_row_index=self.calendar_reference_row_index,
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
                ic()
                for cell_column_index_day in range(
                    cell_column_index, end_column + 31
                ):
                    if (
                        self.ws.Cells(
                            calendar_start_row_index
                            + self.calendar_reference_row_index,
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


if __name__ == "__main__":
    gantt_chart = Gantt(
        excel_file_name=EXCEL_FILE_NAME,
        start_year=START_YEAR,
        start_month=START_MONTH,
        end_year=END_YEAR,
        end_month=END_MONTH,
    )
    gantt_chart.add_items(start_row_index=5, start_column_index=1, items=ITEMS)
    gantt_chart.add_calendars()

    tasks = [
        [6, "Kenta", 3, "2021-08-02"],
        [7, "Kenta", 8, "2021-08-06"],
        [8, "Kenta", 6, "2021-08-18"],
    ]
    for task in tasks:
        gantt_chart.assign_task(*task)
