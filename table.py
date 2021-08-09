import calendar

import win32com.client as wc
from icecream import ic

from variables import xlHAlignCenter
from variables import xlToRight


def get_worksheet_object():
    xl = wc.GetObject(Class="Excel.Application")
    wb = xl.Workbooks("demo_01.xlsx")
    ws = wb.Sheets(1)
    return ws


def search_cell_from_year_month_day(
    year, month, day, calendar_start_row_index, calendar_start_column_index
):
    ws = get_worksheet_object()
    end_column = (
        ws.Cells(calendar_start_row_index, ws.Columns.Count)
        .End(xlToRight)
        .Column
    )
    ic(end_column)

    for cell_column_index in range(calendar_start_column_index, end_column):
        if (
            ws.Cells(calendar_start_row_index, cell_column_index).Value == year
            and ws.Cells(calendar_start_row_index + 1, cell_column_index).Value
            == month
        ):
            ic()
            for cell_column_index_day in range(
                cell_column_index, end_column + 31
            ):
                if (
                    ws.Cells(
                        calendar_start_row_index + 2, cell_column_index_day
                    ).Value
                    == day
                ):
                    paint_cell(
                        calendar_start_row_index + 3, cell_column_index_day, 37
                    )
                    break


def assign_task():
    ws = get_worksheet_object()
    ws.Cells(6, 5).Value = "Kenta"

    start_day = ws.Cells(6, 7).Value
    print(f"{start_day:%Y-%m-%d}")
    ic(type(start_day))

    # paint_cell(10, 10, 37)
    search_cell_from_year_month_day(2021, 12, 2, 2, 9)


# done
def paint_cell(cell_row_index, cell_column_index, color):
    worksheet = get_worksheet_object()
    worksheet.Cells(
        cell_row_index, cell_column_index
    ).Interior.ColorIndex = color


def create_table():
    xl = wc.GetObject(Class="Excel.Application")
    wb = xl.Workbooks("demo_01.xlsx")
    ws = wb.Sheets(1)

    # clear
    ws.Cells.Clear()

    items = [
        "No",
        "Group",
        "Category",
        "Status",
        "Assign",
        "Man-hours",
        "Start Day",
        "End Day",
    ]
    start_year = 2021
    start_month = 11
    end_year = 2022
    end_month = 2

    add_items(
        worksheet=ws, start_row_index=5, start_column_index=1, items=items,
    )

    working_month_and_days = calculate_consecutive_month_and_year(
        start_year=start_year,
        start_month=start_month,
        end_year=end_year,
        end_month=end_month,
    )

    calendar_start_column_index = len(items) + 1
    add_all_calendars(
        worksheet=ws,
        working_month_and_days=working_month_and_days,
        start_column_index=calendar_start_column_index,
    )


def add_all_calendars(worksheet, working_month_and_days, start_column_index):
    for working_month_and_day in working_month_and_days:
        year = working_month_and_day[0]
        month = working_month_and_day[1]
        add_calendar(
            worksheet=worksheet,
            start_row_index=2,
            start_column_index=start_column_index,
            year=year,
            month=month,
        )
        start_column_index += len(covert_year_month_to_calendar(year, month))


# done
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


# done
def add_calendar(worksheet, start_row_index, start_column_index, year, month):
    # calender
    # 0-> Monday, 6-> Sunday
    days_and_weeks = covert_year_month_to_calendar(year, month)
    start_row_index = start_row_index
    start_column_index = start_column_index
    # year
    specify_cell_width_and_height(
        worksheet=worksheet,
        row_index=start_row_index,
        column_index=start_column_index,
        cell_width=4,
        cell_height=21,
    )
    worksheet.Cells(start_row_index, start_column_index).Value = year
    worksheet.Cells(
        start_row_index, start_column_index
    ).HorizontalAlignment = xlHAlignCenter

    # month
    specify_cell_width_and_height(
        worksheet=worksheet,
        row_index=start_row_index + 1,
        column_index=start_column_index,
        cell_width=4,
        cell_height=21,
    )
    worksheet.Cells(start_row_index + 1, start_column_index).Value = month
    worksheet.Cells(
        start_row_index + 1, start_column_index
    ).HorizontalAlignment = xlHAlignCenter

    # day and day_of_week
    for index, day_and_week in enumerate(days_and_weeks):
        worksheet.Cells(
            start_row_index + 2, index + start_column_index
        ).Value = day_and_week[0]

        specify_cell_width_and_height(
            worksheet=worksheet,
            row_index=start_row_index + 2,
            column_index=index + start_column_index,
            cell_width=4,
            cell_height=21,
        )
        worksheet.Cells(
            start_row_index + 2, index + start_column_index
        ).HorizontalAlignment = xlHAlignCenter
        worksheet.Cells(
            start_row_index + 3, index + start_column_index
        ).Value = day_and_week[1]
        worksheet.Cells(
            start_row_index + 3, index + start_column_index
        ).HorizontalAlignment = xlHAlignCenter


def specify_cell_width_and_height(
    worksheet, row_index, column_index, cell_width, cell_height
):
    worksheet.Cells(row_index, column_index).ColumnWidth = cell_width
    worksheet.Cells(row_index, column_index).RowHeight = cell_height


def add_items(worksheet, start_row_index, start_column_index, items):
    # input string
    for index, item in enumerate(items):
        worksheet.Cells(
            start_row_index, index + start_column_index
        ).Value = item
        worksheet.Cells(
            start_row_index, index + start_column_index
        ).ColumnWidth = 10
        worksheet.Cells(
            start_row_index, index + start_column_index
        ).HorizontalAlignment = xlHAlignCenter
        worksheet.Cells(
            start_row_index, index + start_column_index
        ).Interior.ColorIndex = 37


# done
def covert_year_month_to_calendar(year, month):
    cal = calendar.Calendar()
    day_day_of_week = cal.monthdays2calendar(year=year, month=month)
    arr = []
    for data_per_week in day_day_of_week:
        for data in data_per_week:
            if data[0] == 0:
                continue
            else:
                new_data = list(data)
                new_data[1] = get_day_of_week_from_num(data[1])

                arr.append(new_data)
    return arr


# done
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
