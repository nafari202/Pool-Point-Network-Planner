"""Simulated outbound network: one DC, eight pool points, four carriers, 136 stores.

Locations are real cities with approximate road miles from a South Jersey DC. Carriers,
stores, rates, and service levels are fictional.
"""
import math

import numpy as np
import pandas as pd

from . import config

CARRIERS = pd.DataFrame([
    # Linehaul cost per trailer = fixed charge + miles x rate per mile
    {"carrier_id": "CARR-A", "carrier_name": "Carrier A", "fixed_charge": 275, "rate_per_mile": 3.05, "base_on_time": 0.965},
    {"carrier_id": "CARR-B", "carrier_name": "Carrier B", "fixed_charge": 250, "rate_per_mile": 2.85, "base_on_time": 0.925},
    {"carrier_id": "CARR-C", "carrier_name": "Carrier C", "fixed_charge": 300, "rate_per_mile": 3.20, "base_on_time": 0.975},
    {"carrier_id": "CARR-D", "carrier_name": "Carrier D", "fixed_charge": 260, "rate_per_mile": 2.95, "base_on_time": 0.950},
])

POOL_POINTS = pd.DataFrame([
    # handling_per_pallet = cross-dock fee; cost_per_stop = local delivery to one store
    {"pool_point_id": "PP01", "city": "Albany", "state": "NY", "miles_from_dc": 205, "carrier_id": "CARR-A", "handling_per_pallet": 9.50, "cost_per_stop": 165, "winter_exposed": True},
    {"pool_point_id": "PP02", "city": "Hartford", "state": "CT", "miles_from_dc": 190, "carrier_id": "CARR-A", "handling_per_pallet": 10.25, "cost_per_stop": 180, "winter_exposed": True},
    {"pool_point_id": "PP03", "city": "Worcester", "state": "MA", "miles_from_dc": 265, "carrier_id": "CARR-B", "handling_per_pallet": 10.75, "cost_per_stop": 195, "winter_exposed": True},
    {"pool_point_id": "PP04", "city": "Harrisburg", "state": "PA", "miles_from_dc": 120, "carrier_id": "CARR-C", "handling_per_pallet": 8.75, "cost_per_stop": 150, "winter_exposed": False},
    {"pool_point_id": "PP05", "city": "Pittsburgh", "state": "PA", "miles_from_dc": 300, "carrier_id": "CARR-C", "handling_per_pallet": 9.25, "cost_per_stop": 170, "winter_exposed": True},
    {"pool_point_id": "PP06", "city": "Baltimore", "state": "MD", "miles_from_dc": 115, "carrier_id": "CARR-D", "handling_per_pallet": 9.75, "cost_per_stop": 175, "winter_exposed": False},
    {"pool_point_id": "PP07", "city": "Richmond", "state": "VA", "miles_from_dc": 255, "carrier_id": "CARR-D", "handling_per_pallet": 8.50, "cost_per_stop": 155, "winter_exposed": False},
    {"pool_point_id": "PP08", "city": "Syracuse", "state": "NY", "miles_from_dc": 250, "carrier_id": "CARR-B", "handling_per_pallet": 9.00, "cost_per_stop": 160, "winter_exposed": True},
])

# Average weekly pallets for a store of each size at a demand index of 1.0
SIZE_BASE_PALLETS = {"Small": 9, "Medium": 14, "Large": 21}
SIZE_PROBS = {"Small": 0.30, "Medium": 0.50, "Large": 0.20}
DAY_PATTERNS = {
    2: [("Mon", "Thu"), ("Tue", "Fri"), ("Wed", "Sat")],
    3: [("Mon", "Wed", "Fri"), ("Tue", "Thu", "Sat")],
}
# Pallet max leaves 30% headroom over a normal week, so holiday peaks create backlog
PALLET_MAX_HEADROOM = 1.30

# Stores already scheduled to open during the forecast horizon (new store additions)
FUTURE_OPENINGS = [("PP04", "2026-09-12"), ("PP07", "2026-10-17")]


def lane_cost(pool_points: pd.DataFrame, carriers: pd.DataFrame) -> pd.Series:
    """Linehaul cost per trailer for each pool point lane, indexed by pool point."""
    lanes = pool_points.merge(carriers, on="carrier_id")
    cost = lanes["fixed_charge"] + lanes["miles_from_dc"] * lanes["rate_per_mile"]
    return pd.Series(cost.round(2).values, index=lanes["pool_point_id"], name="linehaul_cost_per_trailer")


def _delivery_days(size: str, rng: np.random.Generator) -> tuple:
    if size == "Small":
        n_days = 2
    elif size == "Large":
        n_days = 3
    else:
        n_days = 3 if rng.random() < 0.3 else 2
    patterns = DAY_PATTERNS[n_days]
    return patterns[rng.integers(len(patterns))]


def _make_store(store_id, pool_point_id, market, open_date, rng, size=None):
    size = size or rng.choice(list(SIZE_PROBS), p=list(SIZE_PROBS.values()))
    multiplier = float(np.exp(rng.normal(0, 0.15)))
    days = _delivery_days(size, rng)
    base = SIZE_BASE_PALLETS[size] * multiplier
    pallet_max = int(np.clip(math.ceil(base * PALLET_MAX_HEADROOM / len(days)), 4, 18))
    return {
        "store_id": store_id,
        "pool_point_id": pool_point_id,
        "market": market,
        "size_tier": size,
        "base_weekly_pallets": round(base, 2),
        "delivery_days": "/".join(days),
        "deliveries_per_week": len(days),
        "pallet_max": pallet_max,
        "open_date": pd.Timestamp(open_date),
    }


def build_stores(rng: np.random.Generator) -> pd.DataFrame:
    stores = []
    store_num = 1001
    history_weeks = pd.date_range(config.HISTORY_START, config.HISTORY_END, freq="W-SAT")
    for pp in POOL_POINTS.itertuples():
        for _ in range(int(rng.integers(14, 23))):
            # About 5% of stores opened during the history window; the rest are mature
            if rng.random() < 0.05:
                open_date = history_weeks[rng.integers(20, len(history_weeks) - 20)]
            else:
                open_date = "2015-01-01"
            stores.append(_make_store(f"S{store_num}", pp.pool_point_id, pp.city, open_date, rng))
            store_num += 1
    for pool_point_id, open_date in FUTURE_OPENINGS:
        city = POOL_POINTS.set_index("pool_point_id").at[pool_point_id, "city"]
        stores.append(_make_store(f"S{store_num}", pool_point_id, city, open_date, rng, size="Medium"))
        store_num += 1
    return pd.DataFrame(stores)


def build_network(seed: int = config.SEED) -> dict:
    rng = np.random.default_rng(seed)
    stores = build_stores(rng)
    pool_points = POOL_POINTS.copy()
    pool_points["linehaul_cost_per_trailer"] = lane_cost(pool_points, CARRIERS).values
    return {"carriers": CARRIERS.copy(), "pool_points": pool_points, "stores": stores, "rng": rng}


def ramp_factor(open_date: pd.Timestamp, week_ends: pd.DatetimeIndex) -> np.ndarray:
    """0 before opening, then a ramp from 55% to 100% of normal volume over eight weeks."""
    weeks_open = ((week_ends - open_date).days // 7).to_numpy()
    ramp = np.clip(0.55 + 0.45 * weeks_open / 8, 0.55, 1.0)
    return np.where(weeks_open < 0, 0.0, ramp)
