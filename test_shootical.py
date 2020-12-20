import pytest
import extract_xlsx_cal
from datetime import datetime


def test_event_datetime():
    dtpast = datetime.strptime("19780311T121314","%Y%m%dT%H%M%S")
    assert dtpast.year == 1978
    evt_start, evt_end = extract_xlsx_cal.event_datetime(dtpast, ["1200-1400"])
    assert len(evt_start) == 15
    assert len(evt_end) == 15