import datetime
import pandas as pd

CAPIQ_DATE_STR_FORMAT = '%m/%d/%Y'


def freq_and_periods_to_begin_date_str(freq: str, periods: int) -> str:
    ts = _freq_and_periods_to_begin_date(freq, periods)
    return ts.strftime(CAPIQ_DATE_STR_FORMAT)


def _freq_and_periods_to_begin_date(freq: str, periods: int) -> pd.Timestamp:
    # Map legacy CIQ frequency codes to pandas offset aliases.
    # Pandas 3.x removed bare 'Q' and 'Y'; they are now 'QE' and 'YE'.
    _FREQ_MAP = {"Q": "QE", "Y": "YE"}
    pd_freq = _FREQ_MAP.get(freq, freq)
    all_dates = pd.date_range(end=datetime.datetime.today(), periods=periods, freq=pd_freq)
    return all_dates[0]


def today_as_str() -> str:
    return datetime.datetime.today().strftime(CAPIQ_DATE_STR_FORMAT)