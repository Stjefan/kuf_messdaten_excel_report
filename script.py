from io import BytesIO
from src.kuf_messdaten_excel_report.wochenbericht import (
    erstelle_xslx_baulaerm_wochenbericht,
)
from src.kuf_messdaten_excel_report.monatsbericht import (
    UebersichtImmissionsort,
    erstelle_xslx_monatsbericht,
    UebersichtMonat,
)

from datetime import datetime, timedelta, date
from calendar import monthrange
from uuid import UUID

from dotenv import load_dotenv
import os
import pandas as pd
import numpy as np

from src.kuf_messdaten_excel_report.database_connection import (
    ExcelReportDbService,
    ImmissionsortHelper,
)

import os, logging
from src.kuf_messdaten_excel_report.html_tables import create_html_table, fun_with_styling
from src.kuf_messdaten_excel_report.png_charts import create_png_charts

from src.kuf_messdaten_excel_report import create_monatsbericht_mannheim, create_monatsbericht_immendingen



if __name__ == "__main__":
    print("Hello from script")
    if False:
        mp_1 = (UUID("16b2a784-8b6b-4b7e-9abf-fd2d5a8a0091"), "mp1")
        mp_3 = (UUID("d0aa76cf-36e8-43d1-bb62-ff9cc2c275c0"), "mp3")
        mp_4 = (UUID("ab4e7e2d-8c39-48c2-b80c-b80f6b619657"), "mp4")
        mp_2 = (UUID("965157eb-ab17-496f-879a-55ce924f6252"), "mp2")

        for mp in [
            mp_1, mp_2, mp_3, 
            mp_4]:
            m = ExcelReportDbService(CS)
            # m.get_fremdgeraeuschpegel(datetime(2023, 4, 1, 0, 0, 0), datetime(2023, 4, 30, 23, 59,59), mp[0])
            m.get_wochenuebersicht_vorhandene_messdaten(datetime(2023, 5, 1, 0, 0, 0), datetime(2023, 5, 1, 20, 0,0), mp[0])
    day_in_week = datetime(2023, 5, 7)
    if False:
        
        

        week_number = day_in_week.isocalendar().week
        week_start_date = date.fromisocalendar(2023, week_number, 1)
        week_end_date = date.fromisocalendar(2023, week_number, 7)
        week_start = datetime(week_start_date.year, week_start_date.month, week_start_date.day)
        week_end = datetime(week_end_date.year, week_end_date.month, week_end_date.day)
        m = ExcelReportDbService(CS)
        id_project =  UUID("8d7e0d22-620c-45b4-ac38-25b63ddf79e0")
        io_2 = (UUID("c4862493-478b-49ec-ba03-a779551bf575"), "io2", "mp2")
        io_1 = (UUID("f4311d0b-cd3a-4cf1-a0df-d4f1a5edbef7"), "io1", "mp1")
        io_3= (UUID("c27fe3cd-af55-43ec-9a52-0b2aec78df8b"), "io3", "mp3")
        io_4= (UUID("89b09198-44ee-43b9-bb03-a0a138c6d26a"), "io4", "mp4")

        mp_1 = (UUID("16b2a784-8b6b-4b7e-9abf-fd2d5a8a0091"), "mp1")
        mp_3 = (UUID("d0aa76cf-36e8-43d1-bb62-ff9cc2c275c0"), "mp3")
        mp_4 = (UUID("ab4e7e2d-8c39-48c2-b80c-b80f6b619657"), "mp4")
        mp_2 = (UUID("965157eb-ab17-496f-879a-55ce924f6252"), "mp2")

        ios_dict = {}
        mps_dict = {}
        bytesio_obj = BytesIO()
        for d in range(0, 7):
            from_date = week_start + timedelta(days=d) + timedelta(hours=7)
            to_date = from_date + timedelta(hours=13)
            for io in [io_1, io_2, io_3, io_4]:
                id, name, for_mp = io
                r = m.get_wochenbericht_1(id_project, from_date, to_date, id)
                r = r.set_index("time")

                ios_dict[for_mp] = r
            for mp in [mp_1, mp_2, mp_3, mp_4]:
                id, name = mp
                r = m.get_maxpegel_1(None, from_date , to_date, id)
                r = r.set_index("time")

                u1 = m.get_umgebungslaerm_1(None, from_date, to_date, id)
                lr_r = ios_dict[name]
                dti = pd.date_range(from_date, end=to_date, freq='5s')
                dti.name = "time"
                result = pd.DataFrame(index=dti, columns=["maxpegel", "lr"])
                print(result, u1)

                
                result.loc[r.index, "maxpegel"] = r["maxpegel"]
                result.loc[lr_r.index, "lr"] = lr_r["pegel"]
                
                # r.loc[lr_r.index, "lr"] = lr_r["pegel"]
                result = result.reset_index()
                mps_dict[f"{name}_{from_date.strftime('%A')}"] = result
        if True:
            erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, mps_dict, day_in_week)
            target_dir = "."
            with open(os.path.join(target_dir, "wochenbericht_1.xlsx"), "wb") as f:
                f.write(bytesio_obj.getbuffer())
                print("Writing succes")
    if False:
        create_monatsbericht_mannheim(2023, 6)
        # create_monatsbericht_immendingen(2023, 4)
    if False:
        bytesio_obj = BytesIO()
        df_1 = pd.DataFrame(
            np.random.randint(0, 100, size=(int(11 * 3600 / 5), 2)),
            columns=["maxpegel", "lr"],
        )
        df_2 = pd.DataFrame(
            np.random.randint(0, 100, size=(int(13 * 3600 / 5), 2)),
            columns=["maxpegel", "lr"],
        )
        ios_dict = {"io1": df_1, "io3": df_1}
        print(df_1)
        print()
        ts_1 = pd.date_range(datetime.now(), periods=11 * 3600 / 5, freq="s")
        ts_2 = pd.date_range(datetime.now(), periods=13 * 3600 / 5, freq="s")
        df_1["time"] = ts_1
        df_2["time"] = ts_2
        erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, ios_dict)
        # erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, ios_dict)
        target_dir = "."
        with open(os.path.join(target_dir, "wochenbericht_1.xlsx"), "wb") as f:
            f.write(bytesio_obj.getbuffer())
            print("Writing succes")

    # m = MessdatenServiceV3(CS)
    # print(m.get_beurteilungspegel("c4862493-478b-49ec-ba03-a779551bf575", datetime(2023, 4, 13, 20, 0, 0), datetime(2023, 4, 14, 7, 0, 0)))
    if False:
        df_1 = pd.DataFrame(np.random.randint(0, 100, size=(7, 1)), columns=["pegel"])
        df_2 = pd.DataFrame(np.random.randint(0, 100, size=(7, 1)), columns=["pegel"])

        u = UebersichtMonat()

        bytesio_obj = BytesIO()
        ios_dict = {"io1": df_1, "io3": df_2}
        erstelle_xslx_monatsbericht(bytesio_obj, datetime(2023, 4, 10), u)
        # erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, ios_dict)
        target_dir = "."
        with open(os.path.join(target_dir, "monatsbericht_1.xlsx"), "wb") as f:
            f.write(bytesio_obj.getbuffer())
            print("Writing succes")
    if False:
        first = datetime(2023, 4, 1)
        _, no_days = monthrange(first.year, first.month)
        last = first + timedelta(days=no_days - 1)
        df_1 = pd.DataFrame(
            np.random.randint(0, 100, size=(no_days, 6)),
            columns=[
                "lr_tag",
                "max_tag",
                "max_lr_nacht",
                "max_lr_arg",
                "max_nacht",
                "max_nacht_arg",
            ],
            index=pd.date_range(first, last, freq="d"),
        )
        print(df_1)

    mp_name_id_list = [("MP1", UUID("16b2a784-8b6b-4b7e-9abf-fd2d5a8a0091")),
                    ("MP2", UUID("965157eb-ab17-496f-879a-55ce924f6252")),
                    ("MP3", UUID("d0aa76cf-36e8-43d1-bb62-ff9cc2c275c0")),
                    ("MP4", UUID("ab4e7e2d-8c39-48c2-b80c-b80f6b619657"))
                    ]
    destination = "./tables/"
    if False:
        m = ExcelReportDbService()
        c = m.db_connection.connection.cursor()
        
        for mp in mp_name_id_list:
            from_datetime = datetime(2023, 6, 27)
            string_io = create_html_table(from_datetime, mp,c)
            name = mp[0]
            with open(os.path.join(destination, f'{from_datetime.strftime("%Y_Lr_Woche_%V")}_{name}.html'), "w") as f:
                f.write(string_io.getvalue())
                print("Writing succes")
    


    if False:
        from_datetime = datetime(2023, 6, 27)
        destination = "./images/"
        m = ExcelReportDbService()
        c = m.db_connection.connection.cursor()
        
        for mp in mp_name_id_list:
            name = mp[0]
            results = create_png_charts(datetime(2023, 6, 27), mp, c)
            for b in results:
                curr_io: BytesIO = b[2]
                with open(os.path.join(destination, f"lr_{b[0].strftime('%Y_%m_%d')}_{b[1]}.png"), "wb") as f:
                    f.write(curr_io.getvalue())
                    print("Writing succes")
        

