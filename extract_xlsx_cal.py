from openpyxl import load_workbook
from icalendar import Calendar, Event
import datetime
import urllib.request as urq
import sys
import re

local_name = './data/PSUE Schiesskalender 2026_20260301_1.xlsx'
worksheet = 'Tabelle1'


# psue_url = 'http://www.psue.ch/calendar/PSUE_Schiesskalender2021.xlsx'


def xsl_get(url):
    wb = load_workbook(filename=local_name)
    xsl_cal = wb[worksheet]
    return xsl_cal


def event_datetime(day, hours_list):
    if not isinstance(day, datetime.datetime):
        print('{} {}'.format(day, hours_list))
        return
    start_txt = hours_list[0][0:4]
    end_txt = hours_list[-1][5:9]
    day_str = datetime.datetime.strftime(day, '%Y%m%d')
    start_ical_fmt = "{}T{}00".format(day_str, start_txt)
    end_ical_fmt = "{}T{}00".format(day_str, end_txt)
    return start_ical_fmt, end_ical_fmt


def event_from_row(xsl_sheet, row_nr, hours_list):
    day = xsl_sheet.cell(row_nr, column=1).value
    event_name = xsl_sheet.cell(row=row_nr, column=6).value
    event_SL1 = xsl_sheet.cell(row=row_nr, column=4).value
    event_SL2 = xsl_sheet.cell(row=row_nr, column=5).value
    if event_SL1 and event_SL2:
        event_desc = f"{event_name}   (SL: {event_SL1}/{event_SL2} )"
    elif event_SL1 and not event_SL2:
        event_desc = f"{event_name}   (SL: {event_SL1} )"
    else:
        event_desc = event_name

    evt_from, evt_to = event_datetime(day, hours_list)
    evt_txt = "Von:{} Bis:{} Veranstaltung:{}".format(evt_from, evt_to, event_desc)
    cal_evt = Event()
    cal_evt['DTSTART'] = evt_from
    cal_evt['DTEND'] = evt_to
    cal_evt['SUMMARY'] = event_desc
    return cal_evt


def xl_to_calendar(xl_filename):
    xsl_cal = xsl_get(xl_filename)
    evt_time = xsl_cal['C']
    category_training = xsl_cal['G']
    cal_other = Calendar()
    cal_training = Calendar()
    cal_training.add('METHOD', 'REQUEST')
    for row, t in enumerate(evt_time):
        if t.value is None:
            continue
        match = re.findall(r'\d{4}-\d{4}', t.value)
        if match:
            evt = event_from_row(xsl_cal, t.row, match)
            # iff
            if category_training[row].value == "Ja":
                cal_training.add_component(evt)
            else:
                cal_other.add_component(evt)
    events_cal_file_name = xl_filename.rsplit('.', 1)[0] + '_sonst.ics'
    cal_file = open(events_cal_file_name, 'wb')
    cal_file.write(cal_other.to_ical())

    training_file_name = xl_filename.rsplit('.', 1)[0] + '_training.ics'
    training_file = open(training_file_name, 'wb')
    training_file.write(cal_training.to_ical())

    num_events = len(cal_other.subcomponents) + len(cal_training.subcomponents)

    print("{} Events converted from {}".format(num_events, xl_filename))


def main():
    global local_name
    if len(sys.argv) >= 2:  # cmd line argument
        local_name = sys.argv[1]
    # else:    # download
    #     # urq.urlretrieve(psue_url, local_name)

    xl_to_calendar(local_name)


if __name__ == "__main__":
    main()
