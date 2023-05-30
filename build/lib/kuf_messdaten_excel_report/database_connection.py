from typing import ClassVar
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
import logging

from datetime import datetime
import pandas as pd

from uuid import UUID

from dataclasses import dataclass

@dataclass
class ImmissionsortHelper:
    id_in_db: UUID
    id: int = 0
    





@dataclass
class ExcelReportDbService:
    dbname: ClassVar[str] = "dauerauswertung"
    conn_string: str
    alchemyEngine: Engine = None
    
    # q_filtered = f"""SELECT count(*) FROM tsdb_rejected rej JOIN tsdb_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{after_time}' AND time <= '{before_time}' AND m.projekt_id = {projekt_id}"""


    # q_wetter_filter = f"""SELECT count(*) FROM tsdb_rejected rej JOIN tsdb_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{after_time}' AND time <= '{before_time}' AND m.projekt_id = {projekt_id} AND (rej.filter_id = 5 or rej.filter_id = 6)"""
    # q_sonstige_filter = f"""SELECT count(*) FROM tsdb_rejected rej JOIN tsdb_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{after_time}' AND time <= '{before_time}' AND m.projekt_id ={projekt_id} AND rej.filter_id != 5 and rej.filter_id != 6"""

    # q_schalllesitungspegel = f"""SELECT count(*) FROM tsdb_rejected rej JOIN tsdb_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{after_time}' AND time <= '{before_time}' AND m.projekt_id = 1"""

    
    # q_schallleistungpegel = f"""SELECT * FROM tsdb_schallleistungpegel WHERE time >= '{after_time}' AND time <= '{before_time}' AND messpunkt_id = 4;"""



    
    def __post_init__(self):
        # Connect to PostgreSQL server
        self.alchemyEngine = create_engine(
        self.conn_string
        # 'postgresql://postgres:password@127.0.0.1:5432/tsdb'
        )
        self.db_connection = self.alchemyEngine.connect()

    def get_wochenbericht_1(self, projekt_id: UUID, from_date: datetime, to_date: datetime, immissionsort_id = UUID("c4862493478b49ecba03a779551bf575")):
        columns = ["time", "pegel", "laermursache_id", "immissionsort_id", "u1.name"]
        
        q_1_a = f"""select {','.join(columns)} 
            from dauerauswertung_beurteilungspegelbaulaerm b1 JOIN dauerauswertung_laermursacheanimmissionsorten u1 
            ON b1.laermursache_id = u1.id AND u1.name = 'Gesamt'
            WHERE time > '{from_date}' AND time < '{to_date}'
            AND immissionsort_id = '{immissionsort_id}'::uuid ORDER BY time;"""
        # rs = self.db_connection.execute(text(q_1_a))
        print(q_1_a)
        df = pd.read_sql(q_1_a, self.db_connection)
        df['time'] = df['time'].dt.tz_convert('Europe/Berlin')
        df['time'] = df['time'].dt.tz_localize(None)
        print(df)

        return df
    

    def get_fremdgeraeuschpegel(self, from_date: datetime, to_date: datetime, messpunkt_id: UUID):
        q_1_a = f"""select time_bucket('24 hours', time) AS time_group, last(pegel, time), max(time) AS time
            from dauerauswertung_baustellenumgebungspegel WHERE messpunkt_id = '{messpunkt_id}'::uuid AND time > '{from_date}' AND time < '{to_date}' GROUP BY time_group;"""
        # rs = self.db_connection.execute(text(q_1_a))
        print(q_1_a)
        df = pd.read_sql(q_1_a, self.db_connection)
        df['time'] = df['time'].dt.tz_convert('Europe/Berlin')
        df['time'] = df['time'].dt.tz_localize(None)

        print(df)

        return df

    def get_maxpegel_1(self, projekt_id: UUID, from_date: datetime, to_date: datetime, messpunkt_id: UUID):
        columns = ["time", "pegel AS maxpegel", "messpunkt_id"]
        
        q_1_a = f"""select {','.join(columns)} 
            from dauerauswertung_richtungungsgewertetertaktmaximalpegel r JOIN dauerauswertung_LaermursacheAnMesspunkt l ON l.gemessen_an_id = '{messpunkt_id}'::uuid AND l.id = r.messpunkt_id           
            WHERE time > '{from_date}' AND time < '{to_date}'
            ORDER BY time;"""
        # rs = self.db_connection.execute(text(q_1_a))
        print(q_1_a)
        df = pd.read_sql(q_1_a, self.db_connection)
        print(df)
        df['time'] = df['time'].dt.tz_convert('Europe/Berlin')
        df['time'] = df['time'].dt.tz_localize(None)
        df = df.dropna()
        print(df)

        return df
    
    def get_number_verwertbare_sekunden(self, projekt_id: UUID, from_date: datetime, to_date: datetime):
        query_verfuegbare_sekunden = f"""SELECT sum(verhandene_messwerte) FROM {self.dbname}_auswertungslauf WHERE zeitpunkt_im_beurteilungszeitraum >= '{from_date}' AND zeitpunkt_im_beurteilungszeitraum <= '{to_date}' AND zuordnung_id = '{projekt_id}'::uuid;"""

        rs = self.db_connection.execute(text(query_verfuegbare_sekunden))

        return rs.fetchone()[0]
        for row in rs:
            print(row, " vs ", int((to_date - from_date).total_seconds()))
            


    def get_number_aussortiert_wetter(self, projekt_id: UUID, from_date: datetime, to_date: datetime):
        ""
        q_wetter_filter = f"""SELECT count(*) FROM {self.dbname}_rejected rej JOIN {self.dbname}_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{from_date}' AND time <= '{to_date}' AND m.projekt_id = '{projekt_id}'::uuid AND (rej.filter_id = '266d7818-e618-46d2-874e-ce02345116f4'::uuid or rej.filter_id = '163b93e1-5a24-44f2-8aa4-cdc5454bd39b'::uuid)"""
        rs = self.db_connection.execute(text(q_wetter_filter))

        return rs.fetchone()[0]
        for row in rs:
            print(row, " vs ", int((to_date - from_date).total_seconds()))

    def get_number_aussortiert_sonstiges(self, projekt_id: UUID, from_date: datetime, to_date: datetime):
        q_sonstige_filter = f"""SELECT count(*) FROM {self.dbname}_rejected rej JOIN {self.dbname}_messpunkt m ON rej.messpunkt_id = m.id WHERE time >= '{from_date}' AND time <= '{to_date}' AND m.projekt_id = '{projekt_id}'::uuid AND (rej.filter_id != '266d7818-e618-46d2-874e-ce02345116f4'::uuid AND rej.filter_id != '163b93e1-5a24-44f2-8aa4-cdc5454bd39b'::uuid)"""
        rs = self.db_connection.execute(text(q_sonstige_filter))

        return rs.fetchone()[0]
        for row in rs:
            print(row, " vs ", int((to_date - from_date).total_seconds()))

    def get_estimated_single(self, immissionsort: ImmissionsortHelper, from_date: datetime, to_date: datetime) -> pd.DataFrame:
        cols = ["time", "pegel", ]
        q = f"select {','.join(cols)}, time from \"dauerauswertung_lrpegel\" where immissionsort_id = '{immissionsort.id_in_db}'::uuid and time >= '{from_date.astimezone()}' and time < '{to_date.astimezone()}' ORDER BY TIME"
        direction_df = pd.read_sql(q, self.db_connection)

        print(direction_df)

        data_dict = {
            "time": "Timestamp",
            "estimated_1": f"E{immissionsort.id}_estimated_1",
            "estimated_2": f"E{immissionsort.id}_estimated_2",
        }
        direction_df.rename(columns=data_dict, inplace=True)
        # print(resu_df["Timestamp"].iloc[0].tzinfo)
        direction_df['Timestamp'] = direction_df['Timestamp'].dt.tz_convert('Europe/Berlin')
        # cet = pytz.timezone('CET').utcoffset()
        # resu_df['Timestamp'] = resu_df['Timestamp'] + cet
        direction_df['Timestamp'] = direction_df['Timestamp'].dt.tz_localize(None)
        direction_df.set_index("Timestamp", inplace=True)
        print(direction_df)
        logging.debug(direction_df)
        return direction_df
    
    def get_kennzahlen_immissionsort(self, immissionsort_id: UUID, after_time: datetime, before_time: datetime):
        q_max_pegel_night = f"""
    SELECT extract('day' from time), pegel, extract('hour' from time) AS Stunde FROM (
        SELECT time::date AS Date, time, pegel, ROW_NUMBER() OVER (
            PARTITION BY time::date
            ORDER BY pegel DESC, time
        ) as rank FROM {self.dbname}_maxpegel 
        WHERE time >= '{after_time}' AND time <= '{before_time}' AND (time::time <= '06:00' OR time::time >= '22:00') AND immissionsort_id = '{immissionsort_id}'::uuid) T1
    WHERE T1.rank = 1 ORDER BY Date;
    """
        max_pegel_night_df = pd.read_sql(q_max_pegel_night, self.db_connection)
        max_pegel_night_df["date_part"] = max_pegel_night_df["date_part"].astype(int)
        print("max_pegel_night_df", max_pegel_night_df)
        q_max_pegel_day = f"""
            SELECT extract('day' from time), pegel, extract('hour' from time) AS Stunde FROM (
                SELECT time::date AS Date, time, pegel, ROW_NUMBER() OVER (
                PARTITION BY time::date
                ORDER BY pegel DESC, time
            ) as rank FROM {self.dbname}_maxpegel
            WHERE time >= '{after_time}' AND time <= '{before_time}' AND (time::time >= '06:00' OR time::time <= '22:00') AND immissionsort_id = '{immissionsort_id}'::uuid) T1
            WHERE T1.rank = 1;
        """
        max_pegel_day_df = pd.read_sql(q_max_pegel_day, self.db_connection)
        max_pegel_day_df["date_part"] = max_pegel_day_df["date_part"].astype(int)
        print("max_pegel_day_df", max_pegel_day_df)

        q_day = f"""
    SELECT extract('day' from time::date), max(pegel) as lr_pegel FROM {self.dbname}_lrpegel lr WHERE time >= '{after_time}' AND time <= '{before_time}' AND immissionsort_id = '{immissionsort_id}'::uuid GROUP BY time::date;
    """
        day_df = pd.read_sql(q_day, self.db_connection)
        day_df["date_part"] = day_df["date_part"].astype(int)
        print("day_df", day_df)

        q_night = f"""
        SELECT T2.*, T2.time, date_part, extract('day' from T1.date_part) AS DAY_IN_MONTH, T1.pegel, extract('hour' from T2.time) AS HOUR_IN_DAY FROM 
            (SELECT time::date AS date_part, max(pegel) AS pegel FROM {self.dbname}_lrpegel lr WHERE time >= '{after_time}' AND time <= '{before_time}' AND (time::time <= '06:00' OR time::time >= '22:00') AND immissionsort_id = '{immissionsort_id}'::uuid GROUP BY time::date) T1
        JOIN {self.dbname}_lrpegel T2 On T1.date_part = T2.time::date and T1.pegel = T2.pegel AND T2.immissionsort_id = '{immissionsort_id}'::uuid ORDER BY T2.time;
        """
        q_night = f"""
    SELECT extract('day' from time), pegel, extract('hour' from time) AS Stunde FROM (
        SELECT time::date AS Date, time, pegel, ROW_NUMBER() OVER (
            PARTITION BY time::date
            ORDER BY pegel DESC, time
        ) as rank FROM {self.dbname}_lrpegel 
        WHERE time >= '{after_time}' AND time <= '{before_time}' AND (time::time <= '06:00' OR time::time >= '22:00') AND immissionsort_id = '{immissionsort_id}'::uuid) T1
    WHERE T1.rank = 1 ORDER BY Date;
    """
        night_df = pd.read_sql(q_night, self.db_connection)
        print("DEBUG")
        print(night_df)

        return day_df, max_pegel_day_df, night_df, max_pegel_night_df