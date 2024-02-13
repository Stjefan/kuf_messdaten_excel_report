from datetime import datetime, date

def get_start_end_week(day_in_week: datetime):
    week_number = day_in_week.isocalendar().week
    week_start_date = date.fromisocalendar(day_in_week.year, week_number, 1)
    week_end_date = date.fromisocalendar(day_in_week.year, week_number, 7)
    week_start = datetime(week_start_date.year, week_start_date.month, week_start_date.day)
    week_end = datetime(week_end_date.year, week_end_date.month, week_end_date.day)
    return week_start_date, week_end_date


import pytz
DATE_FORMAT = "%Y-%m-%d %H:%M:%S%Z"

est = pytz.timezone("Europe/Berlin")