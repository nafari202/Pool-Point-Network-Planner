"""Simulate four years of weekly store demand and pool point operations."""
import numpy as np
import pandas as pd

from . import config
from .network import ramp_factor


def simulate_store_weeks(stores: pd.DataFrame, demand_index: pd.Series,
                         rng: np.random.Generator) -> pd.DataFrame:
    """Weekly demand, delivered pallets, and backlog for every store.

    Demand follows the real Census seasonality. Each store can receive at most
    pallet_max pallets per delivery, so anything above that rolls into backlog.
    """
    weeks = demand_index.index
    existing = stores[stores["open_date"] <= weeks[-1]].reset_index(drop=True)

    ramps = np.vstack([ramp_factor(d, weeks) for d in existing["open_date"]])
    expected = existing["base_weekly_pallets"].to_numpy()[:, None] * demand_index.to_numpy()[None, :] * ramps
    noise = np.exp(rng.normal(0, 0.10, size=expected.shape))
    demand = np.round(expected * noise)

    weekly_capacity = (existing["pallet_max"] * existing["deliveries_per_week"]).to_numpy()
    delivered = np.zeros_like(demand)
    backlog = np.zeros_like(demand)
    carried = np.zeros(len(existing))
    for w in range(len(weeks)):
        due = demand[:, w] + carried
        delivered[:, w] = np.minimum(due, weekly_capacity)
        carried = due - delivered[:, w]
        backlog[:, w] = carried

    frame = pd.DataFrame({
        "week_ending": np.tile(weeks, len(existing)),
        "store_id": np.repeat(existing["store_id"].to_numpy(), len(weeks)),
        "pool_point_id": np.repeat(existing["pool_point_id"].to_numpy(), len(weeks)),
        "is_open": ramps.ravel() > 0,
        "demand_pallets": demand.ravel(),
        "delivered_pallets": delivered.ravel(),
        "backlog_pallets": backlog.ravel(),
    })
    return frame[frame["is_open"]].drop(columns="is_open").reset_index(drop=True)


def _day_shares(stores: pd.DataFrame) -> pd.DataFrame:
    """Share of each pool point's weekly deliveries that fall on each weekday."""
    rows = []
    for store in stores.itertuples():
        for day in store.delivery_days.split("/"):
            rows.append((store.pool_point_id, day, store.base_weekly_pallets / store.deliveries_per_week))
    by_day = pd.DataFrame(rows, columns=["pool_point_id", "day", "pallets"])
    shares = by_day.pivot_table(index="pool_point_id", columns="day", values="pallets", aggfunc="sum").fillna(0)
    shares = shares.reindex(columns=config.DAYS, fill_value=0)
    return shares.div(shares.sum(axis=1), axis=0)


def simulate_pool_point_weeks(store_weeks: pd.DataFrame, stores: pd.DataFrame,
                              pool_points: pd.DataFrame, carriers: pd.DataFrame,
                              demand_index: pd.Series, rng: np.random.Generator) -> pd.DataFrame:
    """Weekly operations for each pool point lane: trailers, on-time stops, and cost."""
    store_days = stores.set_index("store_id")["deliveries_per_week"]
    sw = store_weeks.assign(stops=store_weeks["store_id"].map(store_days))
    weekly = sw.groupby(["pool_point_id", "week_ending"], as_index=False).agg(
        demand_pallets=("demand_pallets", "sum"),
        pallets_shipped=("delivered_pallets", "sum"),
        backlog_pallets=("backlog_pallets", "sum"),
        stores_with_backlog=("backlog_pallets", lambda s: int((s > 0).sum())),
        open_stores=("store_id", "nunique"),
        stops=("stops", "sum"),
    )

    lanes = pool_points.merge(carriers, on="carrier_id").set_index("pool_point_id")
    shares = _day_shares(stores[stores["open_date"] <= demand_index.index[-1]])
    weekly["demand_index"] = weekly["week_ending"].map(demand_index)

    trailers, on_time = [], []
    for row in weekly.itertuples():
        # One linehaul dispatch per delivery day; partial trailers waste capacity
        daily = row.pallets_shipped * shares.loc[row.pool_point_id].to_numpy()
        daily = daily * np.exp(rng.normal(0, 0.05, size=len(daily)))
        trailers.append(int(np.ceil(daily[daily > 0] / config.PALLETS_PER_TRAILER).sum()))

        lane = lanes.loc[row.pool_point_id]
        p = lane["base_on_time"] - 0.08 * max(0.0, row.demand_index - 1.25)
        if lane["winter_exposed"] and row.week_ending.month in (1, 2):
            p -= 0.04
        p = float(np.clip(p + rng.normal(0, 0.01), 0.80, 0.995))
        on_time.append(int(rng.binomial(int(row.stops), p)))

    weekly["trailers"] = trailers
    weekly["on_time_stops"] = on_time
    weekly["carrier_id"] = weekly["pool_point_id"].map(lanes["carrier_id"])
    weekly["linehaul_cost"] = weekly["trailers"] * weekly["pool_point_id"].map(lanes["linehaul_cost_per_trailer"])
    weekly["handling_cost"] = weekly["pallets_shipped"] * weekly["pool_point_id"].map(lanes["handling_per_pallet"])
    weekly["final_mile_cost"] = weekly["stops"] * weekly["pool_point_id"].map(lanes["cost_per_stop"])
    weekly["total_cost"] = weekly[["linehaul_cost", "handling_cost", "final_mile_cost"]].sum(axis=1).round(2)
    weekly["trailer_fill_pct"] = weekly["pallets_shipped"] / (weekly["trailers"] * config.PALLETS_PER_TRAILER)
    weekly["on_time_pct"] = weekly["on_time_stops"] / weekly["stops"]
    weekly["cost_per_pallet"] = weekly["total_cost"] / weekly["pallets_shipped"]
    return weekly.drop(columns="demand_index")


def set_contracted_capacity(pool_points: pd.DataFrame, pool_weeks: pd.DataFrame) -> pd.DataFrame:
    """Contracted weekly trailers per lane: last year's 80th percentile plus growth room."""
    recent = pool_weeks[pool_weeks["week_ending"] > pool_weeks["week_ending"].max() - pd.Timedelta(weeks=52)]
    p80 = recent.groupby("pool_point_id")["trailers"].quantile(0.80)
    p80 = np.ceil(p80 * (1 + config.CAPACITY_GROWTH_ALLOWANCE)).astype(int)
    totals = recent.groupby("pool_point_id")[["pallets_shipped", "trailers"]].sum()
    fill = totals["pallets_shipped"] / (totals["trailers"] * config.PALLETS_PER_TRAILER)
    out = pool_points.copy()
    out["contracted_trailers_per_week"] = out["pool_point_id"].map(p80)
    out["trailing_fill_pct"] = out["pool_point_id"].map(fill).round(4)
    return out
