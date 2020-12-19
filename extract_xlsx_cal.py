from openpyxl import load_workbook
import datetime
import re

wb = load_workbook(filename = '201214_PSUE_Schiesskalender2021.xlsx')
events = wb['Hauptplan2021']
shtime = events['C']
for t in shtime:
    match = re.findall(r'\d{4}-\d{4}',t.value)
    if match:
        day = events.cell(row=t.row, column=1).value
        print(day.strftime("%b %d %Y "), end='')
        [print (' {} '.format(x), end='') for x in match]
        start_time = match[0][0:4]
        datetime.datetime.strptime(start_time,'%H%M')
        end_time = match[-1][5:9]
        datetime.datetime.strptime(end_time,'%H%M')
        print()

