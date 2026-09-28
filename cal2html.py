from datetime import datetime
from itertools import groupby
from jinja2 import Template
import extract_xlsx_events as extr
import urllib.request as urq
import sys

if len(sys.argv) >= 2:  # cmd line argument
    local_name = sys.argv[1]
else:  # download
    print('Try this: ' + __file__ + ' [Excel file name]')
    exit(0)

evts = extr.xl_to_events(local_name)

template = ""
with open('psue_cal.html.jinja', 'r') as file:
    template = file.read()

# Excel rows are in date order, so consecutive events share a month
months = [{'name': month_evts[0].day.strftime('%B %Y'), 'events': month_evts}
          for month_evts in (list(g) for _, g in groupby(evts, key=lambda e: (e.day.year, e.day.month)))]

event_data = {'months': months,
              'date': datetime.now()}

j2_template = Template(template)

html = j2_template.render(event_data).encode('ascii', 'xmlcharrefreplace')

with open('psue_cal.html', 'w') as file:
    template = file.write(html.decode())
