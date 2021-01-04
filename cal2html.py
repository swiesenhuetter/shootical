from jinja2 import Template 
import extract_xlsx_events as extr
import urllib.request as urq
import sys


local_name = 'psue_cal.xlsx'
psue_url = 'http://www.psue.ch/calendar/201221_PSUE_Schiesskalender2021.xlsx'


if len(sys.argv) > 2: # cmd line argument
    local_name = sys.argv[1]
else:    # download
    urq.urlretrieve(psue_url, local_name)
evts = extr.xl_to_events(local_name)
print(evts)


template = ""
with open('psue_cal.html.jinja','r') as file: 
    template = file.read()


event_data = {'events' : extr.xl_to_events(local_name)}

j2_template = Template(template)
j2_template.globals['odd_month'] = extr.odd_month

html = j2_template.render(event_data)


with open('psue_cal.html','w') as file: 
    template = file.write(html)


