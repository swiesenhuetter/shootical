from datetime import datetime
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
print(evts)

template = ""
with open('psue_cal.html.jinja', 'r') as file:
    template = file.read()

event_data = {'events': extr.xl_to_events(local_name),
              'date': datetime.now()}

j2_template = Template(template)
j2_template.globals['odd_month'] = extr.odd_month

html = j2_template.render(event_data).encode('ascii', 'xmlcharrefreplace')

with open('psue_cal.html', 'w') as file:
    template = file.write(html.decode())
