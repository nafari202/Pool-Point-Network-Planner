"""Turn real monthly Census apparel sales into a weekly demand index."""
import numpy as np
import pandas as pd
import requests

from . import config


def load_census_sales(refresh: bool = False) -> pd.DataFrame:
    """Monthly U.S. clothing-store sales in $ millions, cached under data/raw."""
    path = config.RAW_DIR / f"{config.FRED_SERIES}.csv"
    if refresh or not path.exists():
        response = requests.get(config.FRED_URL, timeout=30)
        response.raise_for_status()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(response.content)
    sales = pd.read_csv(path)
    sales.columns = ["month", "sales_musd"]
    sales["month"] = pd.to_datetime(sales["month"])
    return sales


def history_weeks() -> pd.DatetimeIndex:
    return pd.date_range(config.HISTORY_START, config.HISTORY_END, freq="W-SAT")


def _as_days(dates) -> np.ndarray:
    return np.asarray(dates, dtype="datetime64[D]").astype(float)


def weekly_demand_index(sales: pd.DataFrame, week_ends: pd.DatetimeIndex) -> pd.Series:
    """Weekly index of store shipment demand, averaging 1.0 over the history.

    Each month's sales become a daily rate pinned to mid-month. The rate is interpolated
    to the middle of each retail week, then shifted so shipments lead sales.
    """
    s = sales.copy()
    days = s["month"].dt.days_in_month
    daily_rate = s["sales_musd"] / days
    mid_month = s["month"] + pd.to_timedelta(days / 2, unit="D")

    lead = pd.Timedelta(weeks=config.SHIP_LEAD_WEEKS)
    week_mid = week_ends - pd.Timedelta(days=3) + lead
    rate = np.interp(_as_days(week_mid), _as_days(mid_month), daily_rate)

    index = pd.Series(rate, index=week_ends, name="demand_index")
    return index / index.mean()
