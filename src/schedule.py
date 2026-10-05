"""Plan one peak week of store deliveries and measure the impact of schedule changes."""
import math

import numpy as np
import pandas as pd

from . import config

CHANGE_COLUMNS = ["change_id", "change_type", "pool_point_id", "store_id", "day", "delivery_days",
                  "old_value", "new_value", "opening_pallets", "reason", "requested_by"]


# ---------------------------------------------------------------- core planning
def build_schedule(stores: pd.DataFrame, demand: pd.Series, cancelled: set = frozenset()):
    """Spread each store's pallets evenly over its delivery days, capped at pallet max.

    Pallets that don't fit roll to the next delivery. Whatever is left at the end of
    the week is backlog. Returns (deliveries, backlog by store).
    """
    rows, backlog = [], {}
    for store in stores.itertuples():
        # Round half up so the result matches Excel's ROUND in the report
        remaining = int(math.floor(float(demand.get(store.store_id, 0.0)) + 0.5))
        days = [d for d in store.delivery_days.split("/") if (store.store_id, d) not in cancelled]
        for i, day in enumerate(days):
            target = math.ceil(remaining / (len(days) - i)) if remaining > 0 else 0
            pallets = min(target, int(store.pallet_max))
            remaining -= pallets
            if pallets > 0:
                rows.append({"store_id": store.store_id, "pool_point_id": store.pool_point_id,
                             "day": day, "pallets": pallets, "pallet_max": int(store.pallet_max)})
        backlog[store.store_id] = remaining
    deliveries = pd.DataFrame(rows, columns=["store_id", "pool_point_id", "day", "pallets", "pallet_max"])
    return deliveries, pd.Series(backlog, name="backlog_pallets")


def lane_costs(pool_points: pd.DataFrame, carriers: pd.DataFrame) -> pd.DataFrame:
    lanes = pool_points.drop(columns=["linehaul_cost_per_trailer"], errors="ignore").merge(carriers, on="carrier_id")
    lanes["linehaul_cost_per_trailer"] = lanes["fixed_charge"] + lanes["miles_from_dc"] * lanes["rate_per_mile"]
    return lanes.set_index("pool_point_id")


def summarize(deliveries: pd.DataFrame, backlog: pd.Series, pool_points: pd.DataFrame,
              carriers: pd.DataFrame) -> dict:
    lanes = lane_costs(pool_points, carriers)
    daily = deliveries.groupby(["pool_point_id", "day"], as_index=False).agg(
        pallets=("pallets", "sum"), stops=("store_id", "count"))
    daily["trailers"] = np.ceil(daily["pallets"] / config.PALLETS_PER_TRAILER).astype(int)
    linehaul = (daily["trailers"] * daily["pool_point_id"].map(lanes["linehaul_cost_per_trailer"])).sum()
    handling = (daily["pallets"] * daily["pool_point_id"].map(lanes["handling_per_pallet"])).sum()
    final_mile = (daily["stops"] * daily["pool_point_id"].map(lanes["cost_per_stop"])).sum()
    pallets = int(daily["pallets"].sum())
    trailers = int(daily["trailers"].sum())
    total = linehaul + handling + final_mile
    return {
        "deliveries": int(daily["stops"].sum()),
        "pallets_delivered": pallets,
        "backlog_pallets": int(backlog.sum()),
        "stores_with_backlog": int((backlog > 0).sum()),
        "trailers": trailers,
        "trailer_fill_pct": pallets / (trailers * config.PALLETS_PER_TRAILER) if trailers else 0.0,
        "linehaul_cost": round(float(linehaul), 2),
        "total_cost": round(float(total), 2),
        "cost_per_pallet": round(float(total / pallets), 2) if pallets else 0.0,
    }


# ---------------------------------------------------------------- change requests
def default_changes(stores: pd.DataFrame, pool_points: pd.DataFrame) -> pd.DataFrame:
    """A realistic week of change requests, picked from the network by rule."""
    def first_store(pool_point_id, size, day=None, by=None):
        pool = stores[(stores["pool_point_id"] == pool_point_id) & (stores["size_tier"] == size)]
        if day:
            pool = pool[pool["delivery_days"].str.contains(day)]
        if by:
            pool = pool.sort_values(by, ascending=False)
        return pool.iloc[0]

    cancel_store = first_store("PP06", "Medium", day="Thu")
    max_store = first_store("PP01", "Large", by="pallet_max")
    new_store_id = f"S{int(stores['store_id'].str[1:].astype(int).max()) + 1}"
    current_carrier = pool_points.set_index("pool_point_id").at["PP08", "carrier_id"]

    changes = [
        {"change_type": "CANCEL_POOL_POINT_DAY", "pool_point_id": "PP03", "day": "Tue",
         "reason": "Winter storm closes the Worcester pool point for Tuesday",
         "requested_by": "Pool point PP03"},
        {"change_type": "CANCEL_STORE_DELIVERY", "pool_point_id": "PP06", "store_id": cancel_store["store_id"],
         "day": "Thu", "reason": "Store running a physical inventory count on Thursday",
         "requested_by": "Store operations"},
        {"change_type": "PALLET_MAX_CHANGE", "pool_point_id": "PP01", "store_id": max_store["store_id"],
         "old_value": str(int(max_store["pallet_max"])), "new_value": str(max(4, int(max_store["pallet_max"]) - 4)),
         "reason": "Backroom at capacity; store asks for smaller drops", "requested_by": "Store operations"},
        {"change_type": "NEW_STORE", "pool_point_id": "PP04", "store_id": new_store_id, "delivery_days": "Tue/Fri",
         "new_value": "10", "opening_pallets": 26,
         "reason": "New store added to the network with grand-opening fill", "requested_by": "Store planning"},
        {"change_type": "CARRIER_CHANGE", "pool_point_id": "PP08", "old_value": current_carrier, "new_value": "CARR-D",
         "reason": "Syracuse lane awarded to a new carrier", "requested_by": "Transportation procurement"},
    ]
    frame = pd.DataFrame(changes).reindex(columns=CHANGE_COLUMNS)
    frame["change_id"] = [f"CHG-{i:03d}" for i in range(1, len(frame) + 1)]
    return frame


def apply_changes(stores, pool_points, demand, changes, cancelled=frozenset()):
    """Return copies of the inputs with every change request applied."""
    stores = stores.copy()
    pool_points = pool_points.copy()
    demand = demand.copy()
    cancelled = set(cancelled)
    for ch in changes.itertuples():
        if ch.change_type == "CANCEL_POOL_POINT_DAY":
            affected = stores[(stores["pool_point_id"] == ch.pool_point_id)
                              & stores["delivery_days"].str.contains(ch.day)]
            cancelled |= {(s, ch.day) for s in affected["store_id"]}
        elif ch.change_type == "CANCEL_STORE_DELIVERY":
            cancelled.add((ch.store_id, ch.day))
        elif ch.change_type == "PALLET_MAX_CHANGE":
            stores.loc[stores["store_id"] == ch.store_id, "pallet_max"] = int(ch.new_value)
        elif ch.change_type == "NEW_STORE":
            days = ch.delivery_days.split("/")
            city = pool_points.set_index("pool_point_id").at[ch.pool_point_id, "city"]
            stores = pd.concat([stores, pd.DataFrame([{
                "store_id": ch.store_id, "pool_point_id": ch.pool_point_id, "market": city,
                "size_tier": "Medium", "delivery_days": ch.delivery_days, "deliveries_per_week": len(days),
                "pallet_max": int(ch.new_value)}])], ignore_index=True)
            demand[ch.store_id] = float(ch.opening_pallets)
        elif ch.change_type == "CARRIER_CHANGE":
            pool_points.loc[pool_points["pool_point_id"] == ch.pool_point_id, "carrier_id"] = ch.new_value
    return stores, pool_points, demand, cancelled


def mitigation(stores, pool_points, backlog, cancelled, changes):
    """Peak response: add a delivery day for every store with backlog, and raise
    pallet max by 25% unless the store itself asked for smaller drops."""
    stores = stores.copy()
    asked_smaller = set(changes.loc[changes["change_type"] == "PALLET_MAX_CHANGE", "store_id"])
    closed = {(c.pool_point_id, c.day) for c in changes.itertuples() if c.change_type == "CANCEL_POOL_POINT_DAY"}
    actions = []
    for store_id in backlog[backlog > 0].index:
        i = stores.index[stores["store_id"] == store_id][0]
        pool_point_id = stores.at[i, "pool_point_id"]
        days = stores.at[i, "delivery_days"].split("/")
        open_days = [d for d in config.DAYS if d not in days
                     and (store_id, d) not in cancelled and (pool_point_id, d) not in closed]
        added = open_days[-1] if open_days else None   # latest free day keeps early-week drops intact
        if added:
            days = sorted(days + [added], key=config.DAYS.index)
            stores.at[i, "delivery_days"] = "/".join(days)
            stores.at[i, "deliveries_per_week"] = len(days)
        old_max = int(stores.at[i, "pallet_max"])
        if store_id not in asked_smaller:
            stores.at[i, "pallet_max"] = min(20, math.ceil(old_max * 1.25))
        actions.append({"store_id": store_id, "pool_point_id": pool_point_id,
                        "backlog_before": int(backlog[store_id]), "added_delivery_day": added or "",
                        "pallet_max_before": old_max, "pallet_max_after": int(stores.at[i, "pallet_max"])})
    return stores, pd.DataFrame(actions)


# ---------------------------------------------------------------- scenario runner
def run_week(stores, pool_points, carriers, store_forecast, plan_week, changes):
    """Baseline, each change on its own, all changes together, and the mitigation plan."""
    week_demand = store_forecast[store_forecast["week_ending"] == plan_week].set_index("store_id")["forecast_pallets"]
    active = stores[stores["open_date"] <= plan_week].copy()

    base_del, base_back = build_schedule(active, week_demand)
    baseline = summarize(base_del, base_back, pool_points, carriers)

    per_change = []
    for _, ch in changes.iterrows():
        s, p, d, c = apply_changes(active, pool_points, week_demand, pd.DataFrame([ch]))
        result = summarize(*build_schedule(s, d, c), p, carriers)
        per_change.append({
            "change_id": ch["change_id"], "change_type": ch["change_type"],
            "backlog_change": result["backlog_pallets"] - baseline["backlog_pallets"],
            "deliveries_change": result["deliveries"] - baseline["deliveries"],
            "trailers_change": result["trailers"] - baseline["trailers"],
            "cost_change": round(result["total_cost"] - baseline["total_cost"], 2),
        })

    s, p, d, c = apply_changes(active, pool_points, week_demand, changes)
    chg_del, chg_back = build_schedule(s, d, c)
    with_changes = summarize(chg_del, chg_back, p, carriers)

    m_stores, actions = mitigation(s, p, chg_back, c, changes)
    mit_del, mit_back = build_schedule(m_stores, d, c)
    mitigated = summarize(mit_del, mit_back, p, carriers)

    scenarios = pd.DataFrame([
        {"scenario": "1. Baseline plan", **baseline},
        {"scenario": "2. With change requests", **with_changes},
        {"scenario": "3. With peak mitigation", **mitigated},
    ])
    final_stores = m_stores.set_index("store_id")
    visibility = (final_stores[["pool_point_id", "market", "size_tier", "delivery_days", "pallet_max"]]
                  .join(d.rename("forecast_pallets"))
                  .join(mit_del.groupby("store_id")["pallets"].sum().rename("planned_pallets"))
                  .join(mit_back.rename("backlog_pallets"))
                  .fillna({"forecast_pallets": 0, "planned_pallets": 0, "backlog_pallets": 0})
                  .reset_index())
    return {
        "plan_week": plan_week,
        "scenarios": scenarios,
        "per_change": pd.DataFrame(per_change),
        "mitigation_actions": actions,
        "deliveries": mit_del,
        "backlog_before_mitigation": chg_back,
        "store_visibility": visibility,
        "stores_final": m_stores,
        "pool_points_final": p,
    }
