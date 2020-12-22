from openpyxl import load_workbook
from icalendar import Calendar, Event
import datetime
import re


def event_datetime(day, hours_list):
    if not isinstance(day, datetime.datetime):
        print ('{} {}'.format(day, hours_list ))
        return
    start_txt = hours_list[0][0:4]
    end_txt = hours_list[-1][5:9]
    day_str = datetime.datetime.strftime(day, '%Y%m%d')
    start_ical_fmt = "{}T{}00".format(day_str, start_txt)
    end_ical_fmt = "{}T{}00".format(day_str, end_txt)
    return start_ical_fmt, end_ical_fmt


def event_from_row(xsl_sheet, row_nr, hours_list):
    day = xsl_sheet.cell(row_nr, column=1).value
#r.d.
    event_name = xsl_sheet.cell(row=row_nr, column=6).value
#
    event_SL1 = xsl_sheet.cell(row=row_nr, column=4).value
    event_SL2 = xsl_sheet.cell(row=row_nr, column=5).value
    event_SL1_str = str(xsl_sheet.cell(row=row_nr, column=4).value)
    event_SL2_str = str(xsl_sheet.cell(row=row_nr, column=5).value)
    if (( event_SL1_str == "None")  and (event_SL2_str == "None")  ) :
        event_desc = event_name
    else:
        event_desc = "{}   (SL: {} / {} )".format(event_name, event_SL1, event_SL2)
#r.d.
    evt_from, evt_to = event_datetime(day, hours_list)
    evt_txt  = "Von:{} Bis:{} Veranstaltung:{}".format(evt_from, evt_to, event_desc)
    cal_evt = Event()
    cal_evt['DTSTART'] = evt_from
    cal_evt['DTEND'] = evt_to
    cal_evt['SUMMARY'] = event_desc
    return cal_evt


def xl_to_calendar(xslfile):
    wb = load_workbook(filename = xslfile)
    xsl_cal = wb['Hauptplan2021']
    shtime = xsl_cal['C']
    cal = Calendar()
    for t in shtime:
        match = re.findall(r'\d{4}-\d{4}',t.value)
        if match:
            evt = event_from_row(xsl_cal, t.row, match)
            cal.add_component(evt)
    cal_file_name = xslfile.rsplit('.', 1)[0] + '.ics'
    cal_file = open(cal_file_name, 'wb')
    cal_file.write(cal.to_ical())


def main():
    xl_to_calendar('201221_PSUE_Schiesskalender2021.xlsx')


if __name__ == "__main__":
    main()
