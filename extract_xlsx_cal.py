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

def xl_to_calendar(xslfile):
    wb = load_workbook(filename = xslfile)
    xsl_cal = wb['Hauptplan2021']
    shtime = xsl_cal['C']
    for t in shtime:
        match = re.findall(r'\d{4}-\d{4}',t.value)
        if match:
            day = xsl_cal.cell(row=t.row, column=1).value
            event_datetime(day, match)
            print()


def main():
    xl_to_calendar('201214_PSUE_Schiesskalender2021.xlsx')


if __name__ == "__main__":
    main()
