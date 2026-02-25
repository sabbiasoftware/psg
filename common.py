from datetime import datetime as dt
from decimal import Decimal as dec
from enum import Enum, auto

def read_strings(fn, do_strip=False, do_lower=False):
    strings = []
    try:
        with open(fn, "r", encoding="utf-8") as f:
            strings = f.read().splitlines()
        if do_strip:
            strings = [s.strip() for s in strings]
        if do_lower:
            strings = [s.lower() for s in strings]
        return strings
    except Exception as exc:
        print(f"Could not open '{fn}'")
        print(f"Exception: {type(exc)}, Arguments: {exc.args}")
        return []


def read_dates(fn):
    try:
        dates = []
        with open(fn, "r") as f:
            dates = [dt.strptime(s.strip(), "%Y-%m-%d") for s in f]
        return dates
    except Exception as exc:
        print(f"Could not parse '{fn}'")
        print(f"Exception: {type(exc)}, Arguments: {exc.args}")
        return []


# obsolete...
def format_hours(hours):
    if isinstance(hours, int):
        return str(hours)
    elif isinstance(hours, dec):
        if hours.as_integer_ratio()[1] == 1:
            return "{:.0f}".format(hours)
        else:
            return "{:.2f}".format(hours)
    else:
        print(type(hours))
        return "???"

def dec_to_number(d):
    if d.as_integer_ratio()[1] == 1:
        return int(d)
    else:
        return float(d)


def format_date(t):
    return dt.strftime(t, "%Y-%m-%d")


def format_datetime(t):
    return dt.strftime(t, "%Y-%m-%d   %H:%M:%S")

class HourType(Enum):
    WORK = auto()
    VACATION = auto()
    SICK = auto()
    HOLIDAY = auto()
    STANDBY = auto()

class HourFormat(Enum):
    WORK = auto()
    UNDER = auto()
    OVER = auto()
    VACATION = auto()
    SICK = auto()
    HOLIDAY = auto()
    MISS = auto()
    QUESTION = auto()
    EMPTY = auto()
