
import os
from datetime import date, datetime, timedelta, time
from calendar import monthrange
import locale
import typing
import xlsxwriter
import pandas as pd
import numpy as np

import logging
from io import BytesIO
import importlib.resources
from .excel_helper import ExcelWorksheet, ExcelCell

rel_kufi_fußzeile_thumbnail = "resources/kufi_fußzeile_thumbnail.jpg"
rel_kufi_logo_thumbnail = "resources/kf_logo_thumbnail.jpg"
kufi_fußzeile_thumbnail = importlib.resources.files(__package__).joinpath(rel_kufi_fußzeile_thumbnail)
kufi_logo_thumbnail = importlib.resources.files(__package__).joinpath(rel_kufi_logo_thumbnail)

gray = "#808080"
dark_gray = "#A9A9A9"
silver = "#C0C0C0"
light_gray = "#D3D3D3"
white_smoke = "#F5F5F5"
green = "#008000"


def erstelle_xslx_baulaerm_wochenbericht(output: BytesIO, ios: dict[str, pd.DataFrame], day_in_week: date):

    workbook = xlsxwriter.Workbook(output)
    my_sheetname = "Übersicht"
    worksheet = workbook.add_worksheet(my_sheetname)
    
    week_number = day_in_week.isocalendar().week
    week_start = date.fromisocalendar(2023, week_number, 1)
    week_end = date.fromisocalendar(2023, week_number, 7)

    first_of_month = week_start

    ausgewerteter_monat = first_of_month.strftime("%B %Y")
    curr_excel_ws = ExcelWorksheet(50, 12, workbook, worksheet)

    curr_excel_ws.ws_cells[(10, 0)].content = f"Woche vom {week_start} zum {week_end}"
    curr_excel_ws.write_to_workbook()
    worksheet.set_header('&C&G&R\n\n\nSeite &P von &N', {"image_center": kufi_logo_thumbnail, 'margin': 0.5})
    worksheet.set_footer('&L&F&C&G&RDatum: &D', {"image_center": kufi_fußzeile_thumbnail, 'margin': 0.6 / 2.54})

    # create_content_worksheet_uebersicht(workbook, worksheet, data, number_days_in_month, data.monat)

    worksheet.set_margins(left=1.5/2.54, right=1/2.54, top=4/2.54, bottom=2.3/2.54)
    for i in ios.keys():
        ws = add_worksheet_wochenbericht_immissionsort(workbook, i, ios[i])
        add_chart(workbook, ws, "B12", len(ios[i]), "Tagzeitraum")
        # add_chart(workbook, ws, "B48", len(ios[i]), "Nachtzeitraum")    

    workbook.close()
    logging.debug("Finished successfully")


def add_worksheet_wochenbericht_immissionsort(wb, name: str, data: pd.DataFrame):
    worksheet = wb.add_worksheet(name)
    number_days_in_month = len(data)
    curr_excel_ws = ExcelWorksheet(number_days_in_month, 12, wb, worksheet)
    bezeichnung_worksheet = worksheet.get_name()

   
    

    
    # number_days_in_month = 30

    curr_excel_ws.make_box_around_cells(0, 3, 2, 8, 2, bg_color=white_smoke)
    # curr_excel_ws.make_box_around_cells(3, 5, number_days_in_month+2, 8, 2, bg_color=silver)
    # curr_excel_ws.make_box_around_cells(0, 1, 2, 2)
    # curr_excel_ws.make_box_around_cells(3, 1, number_days_in_month+2, 2)

    # worksheet.set_column(0, 0, 1)
    # worksheet.set_column(1, 10, 9)
    # worksheet.set_column(3, 3, 13)
    # worksheet.set_column(4, 4, 11)
    # worksheet.set_column(5, 5, 13)
    # worksheet.set_column(6, 6, 9)
    # worksheet.set_column(7, 7, 12)
    # worksheet.set_column(8, 8, 9)

    # for i in range(1, 35):
    #     worksheet.set_row(i, 12)
    # worksheet.set_row(2, 30)

    curr_excel_ws.make_box_around_cells(0, 0, 3, 10, boxtype=1, bg_color=silver)
    # curr_excel_ws.make_box_around_cells(3, 4, number_days_in_month + 2, 4, boxtype=1)
    # curr_excel_ws.make_box_around_cells(3, 5, number_days_in_month + 2, 5, boxtype=1)
    # curr_excel_ws.make_box_around_cells(3, 6, number_days_in_month + 2, 6, boxtype=1, bg_color=silver)
    # curr_excel_ws.make_box_around_cells(3, 7, number_days_in_month + 2, 7, boxtype=1, bg_color=silver)
    # curr_excel_ws.make_box_around_cells(3, 8, number_days_in_month + 2, 8, boxtype=1, bg_color=silver)

    # curr_excel_ws.make_box_around_cells(3, 11, number_days_in_month+2, 11)
    # curr_excel_ws.make_box_around_cells(3, 12, number_days_in_month+2, 13, bg_color=silver)
    curr_excel_ws.make_box_around_cells(0, 0, 30, 10, bg_color=white_smoke)

    curr_excel_ws.merge_cells(0, 0, 3, 10)
    curr_excel_ws.ws_cells[(0, 0)].content = "Pegel-Zeit-Verlauf"
    curr_excel_ws.ws_cells[(0, 0)].format.set_bold(True)
    curr_excel_ws.ws_cells[(0, 0)].format.set_align("center")

    

    date_format = wb.add_format({"num_format": "dd/mm/yyyy"})
    datetime_format= wb.add_format({"num_format": 'hh:mm:ss',})
    # print(data["time"].values.tolist())
    datetime_list = []

    datetime_list = data["time"].dt.to_pydatetime().tolist()

    worksheet.write_column("P1", datetime_list, datetime_format)
    worksheet.write_column("O1", map(mapper_func, data["lr"].values.tolist()))
    worksheet.write_column("N1", map(mapper_func, data["maxpegel"].values.tolist()))

    

    curr_excel_ws.write_to_workbook()
    worksheet.print_area(0, 0, 50, 8)
    return worksheet


def mapper_func(i):
    if i is np.NaN:
        return None
    else:
        return i

def add_chart(wb, worksheet, position: str, number_days_in_month, bezeichnung, column_data_1 = "N", column_data_2 = "O", column_categories = "P"):
    bezeichnung_worksheet = worksheet.name
    ausgewerteter_monat = bezeichnung

    line_chart = wb.add_chart({'type': 'line'})
    line_chart.set_title({'name': f"Baulärm - {bezeichnung_worksheet}\n{ausgewerteter_monat}",
                          })
    line_chart.add_series({"name": "5-sec-Maximalpegel LAFmax", 'values': f"'{bezeichnung_worksheet}'!${column_data_1}$4:${column_data_1}${number_days_in_month+3}", 
                           'categories': f"'{bezeichnung_worksheet}'!${column_categories}$4:${column_categories}${number_days_in_month+3}"
                           })
    line_chart.add_series({"name": "Beurteilungspegel", 'values': f"='{bezeichnung_worksheet}'!${column_data_2}$4:${column_data_2}${number_days_in_month+3}",
                           'categories': f"'{bezeichnung_worksheet}'!${column_categories}$4:${column_categories}${number_days_in_month+3}"})
   
    line_chart.set_x_axis(
    {
        # "min": date(2023, 1 , 1),
        # "max": date(2023, 1, 2),
        "date_axis": True,
    }
    )
    line_chart.set_size({'width': 625, 'height': 250})
    line_chart.show_blanks_as('span')
    worksheet.insert_chart(position, line_chart)







if __name__ == "__main__":
    if True:
        bytesio_obj=BytesIO()
        erstelle_xslx_baulaerm_wochenbericht(bytesio_obj)
        target_dir = "."
        with open(os.path.join(target_dir, "wochenbericht_1.xlsx"),"wb") as f:
            f.write(bytesio_obj.getbuffer())
    if False:

        print()
        week_number = date(2023, 4, 28).isocalendar().week
        week_start = date.fromisocalendar(2023, week_number, 1)
        week_end = date.fromisocalendar(2023, week_number, 7)
        print(f"Woche vom {week_start} zum {week_end}")