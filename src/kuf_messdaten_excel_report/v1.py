


from sqlalchemy import create_engine
from sqlalchemy.engine import Engine
import logging
import pandas as pd
from dataclasses import dataclass

from datetime import datetime




from dotenv import load_dotenv
import os




@dataclass
class Messpunkt:
    name: str
    id_in_db: str


@dataclass
class DbService:
    conn_string: str
    alchemyEngine: Engine = None
    


    
    def __post_init__(self):
        # Connect to PostgreSQL server
        self.alchemyEngine = create_engine(
        self.conn_string
        # 'postgresql://postgres:password@127.0.0.1:5432/tsdb'
        )
        self.dbConnection = self.alchemyEngine.connect()

    def get_beurteilungspegel(self, immissionsort_id: str, from_date: datetime, to_date: datetime) -> pd.DataFrame:
        cols = ["estimated_1", "estimated_2"]
        q = f"""select *
        from dauerauswertung_beurteilungspegelbaulaerm b1 JOIN dauerauswertung_laermursacheanimmissionsorten u1 
        ON b1.laermursache_id = u1.id 
        WHERE time > '{from_date.astimezone()}' AND time < '{to_date.astimezone()}'
        AND immissionsort_id = '{immissionsort_id}'::uuid AND name = 'Gesamt' ORDER BY time;"""  # WHERE b1.laermursache_id = '31b9dc20-0f4d-4e15-a530-17b810cada01'::uuid;
        


        df = pd.read_sql(q, self.dbConnection)

        
        # print(resu_df["Timestamp"].iloc[0].tzinfo)
        df['time'] = df['time'].dt.tz_convert('Europe/Berlin')
        # cet = pytz.timezone('CET').utcoffset()
        # resu_df['Timestamp'] = resu_df['Timestamp'] + cet
        df['time'] = df['time'].dt.tz_localize(None)
        df.set_index("time", inplace=True)
        print(df)
        logging.debug(df)
        return df


