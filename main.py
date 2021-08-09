from enum import Enum

from config import Info
from gant_chart import Gantt
from variables import END_MONTH
from variables import END_YEAR
from variables import EXCEL_FILE_NAME
from variables import START_MONTH
from variables import START_YEAR


class Action(Enum):
    ADD = 1
    ASSIGN = 2
    UPDATE = 3
    SAVE = 4


if __name__ == "__main__":
    action = Action.ASSIGN

    info = Info(excel_file_name=EXCEL_FILE_NAME)
    gantt_chart = Gantt(
        excel_file_name=EXCEL_FILE_NAME,
        start_year=START_YEAR,
        start_month=START_MONTH,
        end_year=END_YEAR,
        end_month=END_MONTH,
        check_items=info.config_information.get("check items"),
    )
    if action == Action.ADD:
        gantt_chart.add_items(start_row_index=5, start_column_index=1)
        gantt_chart.add_calendars()

    elif action == Action.ASSIGN:
        tasks = [
            {
                "No": 1,
                "Assign": "Kenta",
                "Urgency": "Urgent",
                "Man-hours": 3,
                "Start Day": "2021-12-01",
                "End Day": "",
                "Status": "On-Going",
            },
            {
                "No": 2,
                "Assign": "Kenta",
                "Urgency": "Urgent",
                "Man-hours": 5,
                "Start Day": "2021-12-31",
                "End Day": "",
                "Status": "On-Going",
            },
        ]
        gantt_chart.store_calendars()
        for task in tasks:
            gantt_chart.assign_task(task)

    elif action == Action.UPDATE:
        gantt_chart.store_calendars()
        gantt_chart.update_calendars()

    elif action == Action.SAVE:
        gantt_chart.save_schedule()
