from io import BytesIO, StringIO
import os
from calendar import monthrange
from uuid import UUID
import locale
from datetime import datetime, date, timedelta
from .utils import get_start_end_week
import pandas as pd

from .database_connection import ExcelReportDbService
locale.setlocale(locale.LC_ALL, 'de')
import pytz
DATE_FORMAT = "%Y-%m-%d %H:%M:%S%Z"

est = pytz.timezone("Europe/Berlin")

import numpy as np
destination = "./tables"

def make_pretty(styler):
            
            styler.format(lambda x: locale.format_string("%.1f", x))

            styler.format_index(lambda v: v.strftime("%A, den %d.%m.%Y"))
            styler.set_table_styles(
    [
    #      {"selector": "", "props": [("border", "1px solid")]},
    #   {"selector": "tbody td", "props": [("border", "1px solid"), ("text-align", "center")]},
    #  {"selector": "th", "props": [("border", "1px solid")]},)
    {"selector": "table", "props": [("width", "100%")]},
    {"selector": "tr:nth-child(even)", "props": [("background-color","#f2f2f2")]},
    {"selector": "tbody td", "props": [("text-align", "center")]},
     {'selector': 'caption',
    'props': [
        ('text-align', 'left'),
        ('font-size', '24px')
    ]}])
            return styler


def fun_with_styling():
     # Create a DateTime index
    date_index = pd.date_range('2023-01-01', periods=5, freq='D')

    # Create random data
    data = np.random.rand(5, 3)

    # Create the DataFrame
    df = pd.DataFrame(data, index=date_index, columns=['A', 'B', 'C'])

    styler = df.style.set_caption("Fun With Styling")
    df = styler.pipe(make_pretty)
    
    df.to_html(
        os.path.join(destination, f'Fun_With_Styling.html'),
        escape=False)

def get_tabelle_baulaerm(cursor, start: datetime, end: datetime, mp_name, mp_id):

    beurteilungszeitraum = "tag"
    # parsed_date = datetime(2023, 5, 1)
    max_y_axis: int = 70
    interval_y_axis: int = 70

    # start = parsed_date
    # end = parsed_date

    start_localized = est.localize(
        datetime(start.year, start.month, start.day, 0, 0, 0)
    )
    end_localized = est.localize(
        datetime(end.year, end.month, end.day, 0, 0, 0)
    )

    

    results = {}
    # with m.db_connection.connection.cursor() as cursor:
        
    for tbl in ["baustellenbeurteilungspegel", "baustellenbeurteilungspegelumgebung"]:
        for el in [
                (f"Mittelungspegel Fremdgeräusch  {mp_name}", mp_id),
            ]:
            name, messpunkt_id = el
            columns = ["time", "pegel"]
            q = f"""
            SELECT time_bucket('24 hours', time) AS time_group, last(pegel, time) AS lr FROM (select {','.join(columns)}, extract('hour' from time) AS HOUR_IN_DAY, time::date AS PARSED_DATE
                from dauerauswertung_{tbl}
                WHERE time > '{start_localized.strftime(DATE_FORMAT)}' AND time < '{end_localized.strftime(DATE_FORMAT)}'
                AND messpunkt_id = '{messpunkt_id}'::uuid) T1 WHERE T1.HOUR_IN_DAY >= 7 AND T1.HOUR_IN_DAY < 20 GROUP BY time_group ORDER BY time_group;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
            cursor.execute(q)

            
            print(q)
            result_dict = cursor.fetchall()
            df = pd.DataFrame(result_dict, columns=["time_group", f"lr_{tbl}"])
            df["time_group"] = df["time_group"].dt.tz_convert("Europe/Berlin")
            df = df.rename(columns={"time_group": "Datum"})
            df = df.set_index("Datum")
            results[tbl] = df

    return pd.merge(left=results["baustellenbeurteilungspegel"], right=results["baustellenbeurteilungspegelumgebung"], left_index=True, right_index=True)


def create_html_table(day_in_week: datetime, mp_name_id) -> StringIO:

    m = ExcelReportDbService()
    c = m.db_connection.connection.cursor()
        
    
    from_datetime, to_datetime = get_start_end_week(day_in_week)

    name, id = mp_name_id
    bytesio_obj = StringIO()
    df = get_tabelle_baulaerm(c, from_datetime, to_datetime + timedelta(days=1), name, id)
    print(df)
    multi_index = pd.MultiIndex.from_product([["Tagzeitraum", "Nachtzeitraum"], ["Lr<sub>Baustelle</sub>(A)", "HG"]])
    
    df["a"] = None
    df["b"] = None
    
    print(df)
    r = pd.DataFrame(df.values, columns=multi_index, index=df.index)
    # df = df.rename(columns={
    #     "lr_baustellenbeurteilungspegel": "Beurteilungspegel Baustelle L<sub>r, tag</sub>(A)",
    #     "lr_baustellenbeurteilungspegelumgebung": "Beurteilungspegel Fremdgeräusch, L<sub>r, HG</sub>(A)"
    # })
    s = r.style.format(precision=1, thousands=".", decimal=",", na_rep="-")
    s.format_index(lambda v: v.strftime("%A, den %d.%m.%Y"))
    s1 = s.set_table_styles([
        {'selector': 'th.col_heading', 'props': 'text-align: center;'},
        {'selector': 'th.col_heading.level0', 'props': 'font-size: 1.5em;'},
        {'selector': 'td', 'props': 'text-align: center;'},
    ], overwrite=False)
    s1.set_caption(name)
    # styler = df.style.set_caption(name)
    # df = styler.pipe(make_pretty)
    
    s1.to_html(
            bytesio_obj,
        # os.path.join(destination, f'{from_datetime.strftime("%Y_Lr_Woche_%V")}_{name}.html'),
        escape=False)
    
    bytesio_obj.seek(0)
    return bytesio_obj
