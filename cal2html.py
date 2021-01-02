from jinja2 import Template 
import extract_xlsx_events as extr

template = ""

with open('psue_cal.html.jinja','r') as file: 
    template = file.read()


data = {'events' : extr.xl_to_events('calendar.xlsx')}


j2_template = Template(template)

html = j2_template.render(data)


with open('psue_cal.html','w') as file: 
    template = file.write(html)
