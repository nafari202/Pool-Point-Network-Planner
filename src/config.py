"""Shared settings for the pool point network planner."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
OUTPUT_DIR = ROOT / "outputs"
IMAGE_DIR = ROOT / "images"

SEED = 42

# Real demand signal: U.S. Census Monthly Retail Trade, clothing & clothing accessories
# stores, not seasonally adjusted, millions of dollars (via FRED).
FRED_SERIES = "MRTSSM448USN"
FRED_URL = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={FRED_SERIES}"

# History is 208 retail weeks ending on Saturdays. The last week is the latest one the
# Census data covers, so it doubles as the planning "as-of" date.
HISTORY_START = "2022-08-06"
HISTORY_END = "2026-07-25"
SHIP_LEAD_WEEKS = 1          # stores receive product about a week before it sells

FORECAST_WEEKS = 26
# Backtest windows start this many weeks before the as-of date and run 26 weeks each:
# 52 = Aug 2025 to Jan 2026 (the same season being forecast), 26 = Feb to Jul 2026.
BACKTEST_ORIGINS = (52, 26)
BACKTEST_WEEKS = 26

PALLETS_PER_TRAILER = 26     # 53' dry van, single-stacked standard pallets
HIGH_SCENARIO_QUANTILE = 0.80
# Contracted lane capacity = 80th percentile of last year's weekly trailers plus growth room
CAPACITY_GROWTH_ALLOWANCE = 0.08

TARGETS = {
    "on_time_pct": 0.95,
    "trailer_fill_pct": 0.75,
}

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
