"""KPI scorecards that benchmark pool points and carriers against targets."""
import numpy as np
import pandas as pd

from . import config


def _rollup(frame: pd.DataFrame, keys) -> pd.DataFrame:
    out = frame.groupby(keys, as_index=False).agg(
        weeks=("week_ending", "nunique"),
        pallets_shipped=("pallets_shipped", "sum"),
        demand_pallets=("demand_pallets", "sum"),
        trailers=("trailers", "sum"),
        stops=("stops", "sum"),
        on_time_stops=("on_time_stops", "sum"),
        total_cost=("total_cost", "sum"),
        avg_backlog_pallets=("backlog_pallets", "mean"),
    )
    out["avg_weekly_pallets"] = out["pallets_shipped"] / out["weeks"]
    out["on_time_pct"] = out["on_time_stops"] / out["stops"]
    out["trailer_fill_pct"] = out["pallets_shipped"] / (out["trailers"] * config.PALLETS_PER_TRAILER)
    out["cost_per_pallet"] = out["total_cost"] / out["pallets_shipped"]
    return out


def pool_point_scorecard(pool_weeks: pd.DataFrame, pool_points: pd.DataFrame,
                         backtest_by_pp: pd.DataFrame, champion: str, weeks: int = 52) -> pd.DataFrame:
    """Trailing-window KPIs per pool point, with year-over-year growth and forecast accuracy."""
    last = pool_weeks["week_ending"].max()
    window = pool_weeks[pool_weeks["week_ending"] > last - pd.Timedelta(weeks=weeks)]
    prior = pool_weeks[(pool_weeks["week_ending"] <= last - pd.Timedelta(weeks=weeks))
                       & (pool_weeks["week_ending"] > last - pd.Timedelta(weeks=2 * weeks))]

    card = _rollup(window, "pool_point_id")
    growth = window.groupby("pool_point_id")["pallets_shipped"].sum() / prior.groupby("pool_point_id")["pallets_shipped"].sum() - 1
    card["yoy_volume_growth"] = card["pool_point_id"].map(growth)
    accuracy = backtest_by_pp[backtest_by_pp["model"] == champion].set_index("pool_point_id")["wape"]
    card["forecast_wape"] = card["pool_point_id"].map(accuracy)
    card = card.merge(pool_points[["pool_point_id", "city", "state", "carrier_id"]], on="pool_point_id")

    network_cpp = window["total_cost"].sum() / window["pallets_shipped"].sum()
    card["cost_index_vs_network"] = card["cost_per_pallet"] / network_cpp
    card["on_time_status"] = np.where(card["on_time_pct"] >= config.TARGETS["on_time_pct"], "Meets", "Below target")
    card["fill_status"] = np.where(card["trailer_fill_pct"] >= config.TARGETS["trailer_fill_pct"], "Meets", "Below target")
    card["cost_status"] = np.where(card["cost_index_vs_network"] > 1.10, "High cost", "In range")

    cols = ["pool_point_id", "city", "state", "carrier_id", "avg_weekly_pallets", "yoy_volume_growth",
            "on_time_pct", "on_time_status", "trailer_fill_pct", "fill_status", "cost_per_pallet",
            "cost_index_vs_network", "cost_status", "avg_backlog_pallets", "forecast_wape",
            "pallets_shipped", "trailers", "stops", "total_cost"]
    return card[cols].sort_values("pool_point_id").reset_index(drop=True)


def carrier_scorecard(pool_weeks: pd.DataFrame, carriers: pd.DataFrame, weeks: int = 52) -> pd.DataFrame:
    last = pool_weeks["week_ending"].max()
    window = pool_weeks[pool_weeks["week_ending"] > last - pd.Timedelta(weeks=weeks)]
    card = _rollup(window, "carrier_id").merge(carriers[["carrier_id", "carrier_name"]], on="carrier_id")
    lanes = window.groupby("carrier_id")["pool_point_id"].nunique()
    card["lanes"] = card["carrier_id"].map(lanes)
    card["on_time_status"] = np.where(card["on_time_pct"] >= config.TARGETS["on_time_pct"], "Meets", "Below target")
    cols = ["carrier_id", "carrier_name", "lanes", "avg_weekly_pallets", "trailers", "on_time_pct",
            "on_time_status", "trailer_fill_pct", "cost_per_pallet", "total_cost"]
    return card[cols]


def network_weekly(pool_weeks: pd.DataFrame) -> pd.DataFrame:
    """Network-level weekly KPI trend."""
    weekly = pool_weeks.groupby("week_ending", as_index=False)[
        ["demand_pallets", "pallets_shipped", "backlog_pallets", "trailers", "stops", "on_time_stops", "total_cost"]].sum()
    weekly["on_time_pct"] = weekly["on_time_stops"] / weekly["stops"]
    weekly["trailer_fill_pct"] = weekly["pallets_shipped"] / (weekly["trailers"] * config.PALLETS_PER_TRAILER)
    weekly["cost_per_pallet"] = weekly["total_cost"] / weekly["pallets_shipped"]
    return weekly
