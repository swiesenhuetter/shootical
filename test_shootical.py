import pytest
import extract_xlsx_cal
from datetime import datetime


def test_event_datetime():
    dtpast = datetime.strptime("19780311T121314","%Y%m%dT%H%M%S")
    assert dtpast.year == 1978
    evt_start, evt_end = extract_xlsx_cal.event_datetime(dtpast, ["1200-1400"])
    assert len(evt_start) == 15
    assert len(evt_end) == 15

import extract_xlsx_events as xle

def test_oddmonth():
    d1 = '01.01.2001'
    assert xle.odd_month(d1)
    assert not xle.odd_month('01.02.2020')

