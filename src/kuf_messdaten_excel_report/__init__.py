
import os
from datetime import date, datetime, timedelta
from calendar import monthrange
import locale
import typing
import xlsxwriter
from uuid import UUID
import logging
from io import BytesIO

from .database_connection import (
    ExcelReportDbService,
    ImmissionsortHelper,
)
from dotenv import load_dotenv
import os
load_dotenv()

from .monatsbericht import UebersichtImmissionsort, UebersichtMonat, erstelle_xslx_monatsbericht

CS = os.getenv("CS_POSTGRES")


locale.setlocale(locale.LC_TIME, 'de_DE')




def create_monatsbericht_immendingen(year: int, month: int, as_buffer=False):
    first_in_month = datetime(year, month, 1)
    _, days_in_month = monthrange(year, month)
    first_next_month = first_in_month + timedelta(days=days_in_month)

    m = ExcelReportDbService(CS)
    id_project =  UUID("4c7be8b7-5515-4ab1-9b49-c5208ff87c08")
    bezeichnung_monatsbericht_file = f"Immendingen_{first_in_month.strftime('%Y_%m')}.xlsx"
    n1 = m.get_number_verwertbare_sekunden(id_project, first_in_month, first_next_month)
    n2 = m.get_number_aussortiert_wetter(id_project, first_in_month, first_next_month)
    n3 = m.get_number_aussortiert_sonstiges(
        id_project, first_in_month, first_next_month
    )

    
    # m.get_estimated_single()
    my_set = {}
    for el in [
        (
            "Immendingen - Ziegelhütte 4",
            UUID("8ab43da3-74e9-4fc4-a09c-36d5906a631a"),
            40,
            37,
        ),
        (
            "Zimmern - Kreutzerweg. 4",
            UUID("dc190b4d-7a15-4484-a94d-55101ec96778"),
            38,
            32,
        ),
        ("Am Hewenegg 1", UUID("0cea8915-b061-4a97-8866-0d2d5748db83"), 52, 46),
        ("Am Hewenegg 8", UUID("e9984e2a-1d1f-441d-9693-3463c4c4f9ea"), 52, 42),
        (
            "Immendingen - Bachzimmererstr. 32",
            UUID("105533ee-93dd-43c2-8df8-aa77f731c5cd"),
            36,
            30,
        ),
    ]:
        id = el[1]
        name = el[0]
        gw_tag = el[2]
        gw_nacht = el[3]
        (
            lr_tag,
            max_pegel_tag,
            lr_nacht,
            max_pegel_nacht,
        ) = m.get_kennzahlen_immissionsort(id, first_in_month, first_next_month)

        io = UebersichtImmissionsort(
            name,
            id=id,
            grenzwert_tag=gw_tag,
            grenzwert_nacht=gw_nacht,
            lr_tag=lr_tag,
            max_pegel_tag=max_pegel_tag,
            lr_nacht=lr_nacht,
            max_pegel_nacht=max_pegel_nacht,
        )
        my_set[io.name] = io
    
    u = UebersichtMonat("Immendingen", n2, n3, n1, my_set)


    bytesio_obj = BytesIO()
    # ios_dict = {"io1": df_1, "io3": df_2}
    erstelle_xslx_monatsbericht(bytesio_obj, first_in_month, u)
    # erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, ios_dict)
    target_dir = "."
    if not as_buffer:
        with open(os.path.join(target_dir, bezeichnung_monatsbericht_file), "wb") as f:
            f.write(bytesio_obj.getbuffer())
            print("Writing succes")
    else:
        return bytesio_obj.getbuffer()


def create_monatsbericht_mannheim(year: int, month: int, as_buffer=False):
    first_in_month = datetime(year, month, 1)
    _, days_in_month = monthrange(year, month)
    first_next_month = first_in_month + timedelta(days=days_in_month)
    m = ExcelReportDbService(CS)
    id_project = UUID("ab0f9a1a-f0b2-4a0d-9868-9b0f6c8d17c4")
    bezeichnung_monatsbericht_file = f"Mannheim_{first_in_month.strftime('%Y_%m')}.xlsx"
    n1 = m.get_number_verwertbare_sekunden(id_project, first_in_month, first_next_month)
    n2 = m.get_number_aussortiert_wetter(id_project, first_in_month, first_next_month)
    n3 = m.get_number_aussortiert_sonstiges(
        id_project, first_in_month, first_next_month
    )
    # m.get_estimated_single()
    my_set = {}
    for el in [
        ("Fichtenweg 2", UUID("7a452a86-453c-4dca-96ed-aca8f756199e"), 55, 45),
        ("Speckweg 18", UUID("4b38584b-9ff4-45cd-981a-54372deda64a"), 55, 45),
        ("Spiegelfabrik 16", UUID("c7974f63-ee9c-475d-8af1-8fa575a560ed"), 55, 45),
    ]:
        id = el[1]
        name = el[0]
        gw_tag = el[2]
        gw_nacht = el[3]
        (
            lr_tag,
            max_pegel_tag,
            lr_nacht,
            max_pegel_nacht,
        ) = m.get_kennzahlen_immissionsort(id, first_in_month, first_next_month)

        io = UebersichtImmissionsort(
            name,
            id=id,
            grenzwert_tag=gw_tag,
            grenzwert_nacht=gw_nacht,
            lr_tag=lr_tag,
            max_pegel_tag=max_pegel_tag,
            lr_nacht=lr_nacht,
            max_pegel_nacht=max_pegel_nacht,
        )
        my_set[io.name] = io

    u = UebersichtMonat("Mannheim", n2, n3, n1, my_set)

    bytesio_obj = BytesIO()
    # ios_dict = {"io1": df_1, "io3": df_2}
    erstelle_xslx_monatsbericht(bytesio_obj, first_in_month, u)
    # erstelle_xslx_baulaerm_wochenbericht(bytesio_obj, ios_dict)
    target_dir = "."
    if not as_buffer:
        with open(os.path.join(target_dir, bezeichnung_monatsbericht_file), "wb") as f:
            f.write(bytesio_obj.getbuffer())
            print("Writing succes")
    else:
        return bytesio_obj.getbuffer()



def add_chart(wb, worksheet):
    bezeichnung_worksheet = worksheet.name
    number_days_in_month = 30
    ausgewerteter_monat = "blub"
    line_chart = wb.add_chart({'type': 'line'})
    line_chart.set_title({'name': f"Beurteilungspegel {bezeichnung_worksheet}\n{ausgewerteter_monat}"})
    line_chart.add_series({"name": "Beurteilungspegel Tag", 'values': f"'{bezeichnung_worksheet}'!$D$4:$D${number_days_in_month+3}"})
    line_chart.add_series({"name": "Grenzwert Tag", 'values': f"='{bezeichnung_worksheet}'!$M$4:$M${number_days_in_month+3}"})
    line_chart.add_series({"name": "Beurteilungspegel Nacht", 'values': f"'{bezeichnung_worksheet}'!$F$4:$F${number_days_in_month+3}"})
    line_chart.add_series({"name": "Grenzwert Nacht", 'values': f"'{bezeichnung_worksheet}'!$N$4:$N${number_days_in_month+3}"})

    line_chart.set_size({'width': 625, 'height': 250})
    worksheet.insert_chart('B36', line_chart)


if __name__ == "__main__":
    if True:
        bytesio_obj=BytesIO()
        erstelle_xslx_monatsbericht(bytesio_obj)
        target_dir = "."
        with open(os.path.join(target_dir, "wochenbericht_1.xlsx"),"wb") as f:
            f.write(bytesio_obj.getbuffer())
    if False:

        print()
        week_number = date(2023, 4, 28).isocalendar().week
        week_start = date.fromisocalendar(2023, week_number, 1)
        week_end = date.fromisocalendar(2023, week_number, 7)
        print(f"Woche vom {week_start} zum {week_end}")
