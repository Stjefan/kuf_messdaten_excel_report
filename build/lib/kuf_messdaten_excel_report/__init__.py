
import os
from datetime import date, datetime
from calendar import monthrange
import locale
import typing
import xlsxwriter

import logging
from io import BytesIO

kufi_fußzeile_thumbnail = "resources/kufi_fußzeile_thumbnail.jpg"
kufi_logo_thumbnail = "resources/kf_logo_thumbnail.jpg"
kufi_fußzeile_thumbnail = "resources/kufi_fußzeile_thumbnail.jpg"
kufi_logo_thumbnail = "resources/kf_logo_thumbnail.jpg"

gray = "#808080"
dark_gray = "#A9A9A9"
silver = "#C0C0C0"
light_gray = "#D3D3D3"
white_smoke = "#F5F5F5"
green = "#008000"


def erstelle_xslx_monatsbericht(output: BytesIO):

    workbook = xlsxwriter.Workbook(output)
    my_sheetname = "Übersicht"
    worksheet = workbook.add_worksheet(my_sheetname)


    worksheet.set_header('&C&G&R\n\n\nSeite &P von &N', {"image_center": kufi_logo_thumbnail, 'margin': 0.5})
    worksheet.set_footer('&L&F&C&G&RDatum: &D', {"image_center": kufi_fußzeile_thumbnail, 'margin': 0.6 / 2.54})

    # create_content_worksheet_uebersicht(workbook, worksheet, data, number_days_in_month, data.monat)

    worksheet.set_margins(left=1.5/2.54, right=1/2.54, top=4/2.54, bottom=2.3/2.54)
    for i in ["MP1", "MP2", "MP3", "MP4"]:
        ws = workbook.add_worksheet(i)

        add_chart(workbook, ws)    

    # if schallleistungspegel_sheet:
    #     erstelle_sheet_schallleistungspegel(workbook,  data.monat.year, data.monat.month, data.schallleistungspegel)

    # for key in data.details_io.keys():
    #     logging.info(f"Writing data for {key}")
    #     monatsuebersicht_an_io = data.details_io[key]
    #     io_details_worksheet = workbook.add_worksheet(monatsuebersicht_an_io.immissionsort.name_4_excel)
    #     io_details_worksheet.set_margins(left=1.5/2.54, right=1/2.54, top=4/2.54, bottom=2.3/2.54)
    #     io_details_worksheet.set_header('&C&G&R\n\n\nSeite &P von &N', {"image_center": os.path.join(os.path.dirname(__file__), kufi_logo_thumbnail), 'margin': 0.5})
    #     # io_details_worksheet.set_footer('&L&F&C&G&RDatum: &D', {"image_center": os.path.join(os.path.dirname(__file__), kufi_fußzeile_thumbnail), 'margin': 0.6 / 2.54})
    #     io_details_worksheet.set_footer('&C&G&R&6Datum: &D', {"image_center": os.path.join(os.path.dirname(__file__), kufi_fußzeile_thumbnail), 'margin': 0.6 / 2.54})
    #     create_content_worksheet_io_details(workbook, io_details_worksheet, monatsuebersicht_an_io, data.monat, number_days_in_month)
            
    

    workbook.close()
    logging.debug("Finished successfully")


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
