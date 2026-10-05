"""Convert the pallet forecast into trailers and compare against carrier capacity."""
import numpy as np
import pandas as pd

from . import config


def lane_outlook(pool_forecast: pd.DataFrame, pool_points: pd.DataFrame) -> pd.DataFrame:
    """Trailers needed per lane and week, at each lane's trailing 52-week fill rate."""
    lanes = pool_points[["pool_point_id", "city", "carrier_id", "contracted_trailers_per_week", "trailing_fill_pct"]]
    out = pool_forecast.merge(lanes, on="pool_point_id")
    pallets_per_trailer = config.PALLETS_PER_TRAILER * out["trailing_fill_pct"]
    out["trailers_forecast"] = np.ceil(out["forecast_pallets"] / pallets_per_trailer).astype(int)
    out["trailers_high"] = np.ceil(out["high_pallets"] / pallets_per_trailer).astype(int)
    out["extra_trailers_needed"] = (out["trailers_high"] - out["contracted_trailers_per_week"]).clip(lower=0)
    out["capacity_status"] = np.select(
        [out["trailers_forecast"] > out["contracted_trailers_per_week"],
         out["trailers_high"] > out["contracted_trailers_per_week"]],
        ["Over capacity", "At risk"], default="Covered")
    return out


def carrier_outlook(lanes: pd.DataFrame, carriers: pd.DataFrame) -> pd.DataFrame:
    """Weekly volume and trailer needs rolled up to each carrier."""
    grouped = lanes.groupby(["carrier_id", "week_ending"], as_index=False).agg(
        lanes=("pool_point_id", "nunique"),
        forecast_pallets=("forecast_pallets", "sum"),
        high_pallets=("high_pallets", "sum"),
        trailers_forecast=("trailers_forecast", "sum"),
        trailers_high=("trailers_high", "sum"),
        contracted_trailers=("contracted_trailers_per_week", "sum"),
        extra_trailers_needed=("extra_trailers_needed", "sum"),
    )
    return grouped.merge(carriers[["carrier_id", "carrier_name"]], on="carrier_id")


def carrier_notes(carrier_weeks: pd.DataFrame, lanes: pd.DataFrame) -> str:
    """Plain-language volume outlook to send each carrier, in Markdown."""
    first, last = carrier_weeks["week_ending"].min(), carrier_weeks["week_ending"].max()
    parts = [f"# Carrier Volume Outlook\n\nForecast window: weeks ending {first:%b %d, %Y} to {last:%b %d, %Y}.\n"
             "Trailer counts use each lane's trailing 52-week fill rate. The high case is the "
             f"{int(config.HIGH_SCENARIO_QUANTILE * 100)}th percentile of backtest error.\n"]
    for carrier_id, weeks in carrier_weeks.groupby("carrier_id"):
        name = weeks["carrier_name"].iloc[0]
        lane_names = lanes.loc[lanes["carrier_id"] == carrier_id, "city"].drop_duplicates().tolist()
        peak = weeks.loc[weeks["trailers_high"].idxmax()]
        surge = weeks[weeks["extra_trailers_needed"] > 0]
        parts.append(f"## {name} ({', '.join(lane_names)})\n")
        parts.append(f"- Contracted capacity: {int(weeks['contracted_trailers'].iloc[0])} trailers per week across these lanes.")
        parts.append(f"- Average forecast: {weeks['trailers_forecast'].mean():.0f} trailers per week, "
                     f"{weeks['forecast_pallets'].mean():,.0f} pallets per week.")
        parts.append(f"- Peak week: {peak['week_ending']:%b %d}, with {int(peak['trailers_forecast'])} trailers "
                     f"forecast and up to {int(peak['trailers_high'])} in the high case.")
        if surge.empty:
            parts.append("- No surge capacity needed in this window.\n")
        else:
            parts.append(f"- Surge capacity needed in {len(surge)} weeks, from {surge['week_ending'].min():%b %d} "
                         f"to {surge['week_ending'].max():%b %d}. Largest request: "
                         f"{int(surge['extra_trailers_needed'].max())} extra trailers in one week.\n")
    return "\n".join(parts)
