from datetime import datetime, date, timedelta
from io import BytesIO
from .utils import get_start_end_week, est, DATE_FORMAT
from uuid import UUID
from .database_connection import ExcelReportDbService
import plotly.graph_objects as go
import os, logging
from plotly.subplots import make_subplots
import pandas as pd

def get_plotly_baulaerm_auswertung_an_messpunkt(parsed_date, mp_name, mp_id, folder_name = "."):
    m = ExcelReportDbService()
    beurteilungszeitraum = "tag"
    # parsed_date = datetime(2023, 5, 1)
    max_y_axis: int = 80
    interval_y_axis: int = 60
    start = parsed_date
    end = parsed_date
    if beurteilungszeitraum == "tag":
        start = est.localize(
            datetime(start.year, start.month, start.day, 7, 0, 0)
        )
        end = est.localize(
            datetime(start.year, start.month, start.day, 20, 0, 0)
        )

    elif beurteilungszeitraum == "nacht":
        start = est.localize(
            datetime(start.year, start.month, start.day, 20, 0, 0)
        )
        end = est.localize(
            datetime(start.year, start.month, start.day, 7, 0, 0)
            + timedelta(hours=24)
        )
    else:
        raise KeyError("Unpassend")

    traces = []
    with m.db_connection.connection.cursor() as cursor:
        for el in [
            (mp_name, mp_id)
        ]:
            name, messpunkt_id = el
            columns = ["time", "pegel"]
            q = f"""select {','.join(columns)} 
                from dauerauswertung_baustellenbeurteilungspegel
                WHERE time > '{start.strftime(DATE_FORMAT)}' AND time < '{end.strftime(DATE_FORMAT)}'
                AND messpunkt_id = '{messpunkt_id}'::uuid;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
            cursor.execute(q)

            
            print(q)
            result_dict = cursor.fetchall()
            df = pd.DataFrame(result_dict, columns=columns)
            df["time"] = df["time"].dt.tz_convert("Europe/Berlin")
            print("Baustellenbeurteilungspegel", df)
            traces.append((
                go.Scatter(
                    x=df["time"], y=df["pegel"], mode="lines", name=f"L<sub>r, Baustelle</sub>"
                ), False)
            )
        
    with m.db_connection.connection.cursor() as cursor:
        for el in [
            (f"Lr Umgebung {mp_name}", mp_id),
            
        ]:
            name, messpunkt_id = el
            columns = ["time", "pegel"]
            q = f"""select {','.join(columns)} 
                from dauerauswertung_baustellenbeurteilungspegelumgebung
                WHERE time > '{start.strftime(DATE_FORMAT)}' AND time < '{end.strftime(DATE_FORMAT)}'
                AND messpunkt_id = '{messpunkt_id}'::uuid;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
            cursor.execute(q)

            
            print(q)
            result_dict = cursor.fetchall()
            df = pd.DataFrame(result_dict, columns=columns)
            df["time"] = df["time"].dt.tz_convert("Europe/Berlin")
            print("dauerauswertung_baustellenbeurteilungspegelumgebung", df)
            traces.append((
                go.Scatter(
                    x=df["time"], y=df["pegel"], mode="lines", name="L<sub>r, Umgebung</sub>"
                ),  False)
            )
        # for el in [
        #     (f"Mittelungspegel Fremdgeräusch  {mp_name}", mp_id),
        # ]:
        #     name, messpunkt_id = el
        #     q2 = f"""select time, pegel, messpunkt_id, berechnet_von_id from dauerauswertung_baustellenumgebungspegel WHERE time >= '{start.strftime(DATE_FORMAT)}' AND time < '{end.strftime(DATE_FORMAT)}' AND messpunkt_id = '{messpunkt_id}'::uuid;"""
        #     cursor.execute(q2)
        #     q2_dict = cursor.fetchall()
        #     q2_cols = ["time", "pegel", "messpunkt_id", "berechnet_von_id"]
        #     df = pd.DataFrame(q2_dict, columns=q2_cols)
        #     print("lm", df)
        #     df["time"] = df["time"].dt.tz_convert("Europe/Berlin")
        #     traces.append((
        #         go.Scatter(
        #                 x=df["time"], y=df["pegel"], mode="lines", name="L<sub>m, 300s, Umgebung</sub>"
        #             ), False)
        #     )    
       


        layout = go.Layout(
            title=f"{mp_name} - {start.strftime('%A, den %d.%m.%Y')} (7:00-20:00)",
            xaxis={"title": "Zeit", "dtick": 60 * 60 * 1000, "range": [
                start, end
            ]},
            yaxis={
                "title": "L<sub>r</sub> [dB(A)]",
                "range": [max_y_axis - interval_y_axis, max_y_axis],
            },
            # yaxis2={
            #     "title": "Lm_300s [dB(A)]",
            #     "range": [max_y_axis - interval_y_axis, max_y_axis],
            # },
            legend={
                "orientation":"h",
                "yanchor": "bottom",
    "y":1.02,
    "xanchor":"right",
    "x": 1
            }
        )
        fig = make_subplots(specs=[[{"secondary_y": False}]])
        for t in traces:
            fig.add_trace(t[0], secondary_y=False)

        fig.update_layout(layout)
        # fig.update_yaxes(title_text="Lr [dB(A)]", secondary_y=False)
        # fig.update_yaxes(title_text="Lm_5min [dB(A)]", secondary_y=True)

        
        # figure = go.Figure(data=traces, layout=layout)
        # fig.show()
        fig.write_image(os.path.join(folder_name, f"lr_{parsed_date.strftime('%Y_%m_%d')}_{mp_name}.png"))



def get_plotly_baulaerm_weekly_charts(cursor, parsed_date, mp_name, mp_id, folder_name = "."):
    beurteilungszeitraum = "tag"
    # parsed_date = datetime(2023, 5, 1)
    max_y_axis: int = 80
    interval_y_axis: int = 60
    start = parsed_date
    end = parsed_date
    if beurteilungszeitraum == "tag":
        start = est.localize(
            datetime(start.year, start.month, start.day, 7, 0, 0)
        )
        end = est.localize(
            datetime(start.year, start.month, start.day, 20, 0, 0)
        )

    elif beurteilungszeitraum == "nacht":
        start = est.localize(
            datetime(start.year, start.month, start.day, 20, 0, 0)
        )
        end = est.localize(
            datetime(start.year, start.month, start.day, 7, 0, 0)
            + timedelta(hours=24)
        )
    else:
        raise KeyError("Unpassend")

    traces = []

    for el in [
        (mp_name, mp_id)
    ]:
        name, messpunkt_id = el
        columns = ["time", "pegel"]
        q = f"""select {','.join(columns)} 
            from dauerauswertung_baustellenbeurteilungspegel
            WHERE time > '{start.strftime(DATE_FORMAT)}' AND time < '{end.strftime(DATE_FORMAT)}'
            AND messpunkt_id = '{messpunkt_id}'::uuid;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
        cursor.execute(q)

        
        print(q)
        result_dict = cursor.fetchall()
        df = pd.DataFrame(result_dict, columns=columns)
        df["time"] = df["time"].dt.tz_convert("Europe/Berlin")
        print("Baustellenbeurteilungspegel", df)
        traces.append((
            go.Scatter(
                x=df["time"], y=df["pegel"], mode="lines", name=f"L<sub>r, Baustelle</sub>"
            ), False)
        )
    

    for el in [
        (f"Lr Umgebung {mp_name}", mp_id),
        
    ]:
        name, messpunkt_id = el
        columns = ["time", "pegel"]
        q = f"""select {','.join(columns)} 
            from dauerauswertung_baustellenbeurteilungspegelumgebung
            WHERE time > '{start.strftime(DATE_FORMAT)}' AND time < '{end.strftime(DATE_FORMAT)}'
            AND messpunkt_id = '{messpunkt_id}'::uuid;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
        cursor.execute(q)

            
        print(q)
        result_dict = cursor.fetchall()
        df = pd.DataFrame(result_dict, columns=columns)
        df["time"] = df["time"].dt.tz_convert("Europe/Berlin")
        print("dauerauswertung_baustellenbeurteilungspegelumgebung", df)
        traces.append((
            go.Scatter(
                x=df["time"], y=df["pegel"], mode="lines", name="L<sub>r, Umgebung</sub>"
            ),  False)
        )

        layout = go.Layout(
            title=f"{mp_name} - {start.strftime('%A, den %d.%m.%Y')} (7:00-20:00)",
            xaxis={"title": "Zeit", "dtick": 60 * 60 * 1000, "range": [
                start, end
            ]},
            yaxis={
                "title": "L<sub>r</sub> [dB(A)]",
                "range": [max_y_axis - interval_y_axis, max_y_axis],
            },
            # yaxis2={
            #     "title": "Lm_300s [dB(A)]",
            #     "range": [max_y_axis - interval_y_axis, max_y_axis],
            # },
            legend={
                "orientation":"h",
                "yanchor": "bottom",
    "y":1.02,
    "xanchor":"right",
    "x": 1
            }
        )
        fig = make_subplots(specs=[[{"secondary_y": False}]])
        for t in traces:
            fig.add_trace(t[0], secondary_y=False)

        fig.update_layout(layout)
        # fig.update_yaxes(title_text="Lr [dB(A)]", secondary_y=False)
        # fig.update_yaxes(title_text="Lm_5min [dB(A)]", secondary_y=True)

        
        # figure = go.Figure(data=traces, layout=layout)
        # fig.show()
        bytes_io = BytesIO()
        fig.write_image(bytes_io
            #os.path.join(folder_name, f"lr_{parsed_date.strftime('%Y_%m_%d')}_{mp_name}.png")
            )
        bytes_io.seek(0)
        return bytes_io

def create_png_charts(day_in_week: datetime, i):
    # day_in_week =  datetime(2023, 6, 5)
    from_datetime, to_datetime = get_start_end_week(day_in_week)

    m = ExcelReportDbService()
    result_list = []
    with m.db_connection.connection.cursor() as c:
        name, id = i
        for d in range(0, 6+1):
            try:
                bytes_io = get_plotly_baulaerm_weekly_charts(c, from_datetime + timedelta(days=d), name, id, f"./images/{name}/")
                result_list.append((from_datetime + timedelta(days=d), name, bytes_io))
            except Exception as ex:
                logging.exception(ex)
    return result_list