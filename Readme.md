# PSUE Excel -> Kalender  
'extract_xlsx_cal.py' ist ein Python-Script dass den PSUE Schiesskalender in Excelform liest und daraus ein geichnamiges *.ics file erzeugt. 
Dieses kann in den Kalender importiert werden (z.B. Google Calendar). 
Das Programm orientiert sich am bestehenden Excel-Format. Das heisst, eine Trainingszeit wird in der Form "1600-1800" erwartet, 
ohne Leerzeichen vor oder nach dem Bindestrich. In der gleichen Zeile befindet sich ein Datum in Spalte eins und eine Beschreibung des
Anlasses in Spalte sechs.
Das Uhrzeit Format wurde im Excel Kalender nicht immer konsequent eingehalten. Beispiel: Am 19.6. steht "0800 -1200" mit Leerzeichen davor und am 9.10. steht "0900 - 1200" 
Natürlich kann man das ganz einfach behandeln, indem mann alle 'whitespace entfernt', dies macht aber das Programm länger und damit unübersichtlicher.


![regextester](doc\RegexTesterCapture.png)
