"""Weekly pallet volume forecasting by pool point.

Three methods are backtested on the last 26 weeks of history:
  1. Seasonal naive: same week last year, scaled by recent year-over-year growth.
  2. Seasonal index: deseasonalized level and trend times a week-of-year index.
     This is the method a planner can rebuild in Excel.
  3. Holt-Winters exponential smoothing with damped trend and weekly seasonality.
The method with the lowest pooled WAPE is refit on all history to forecast 26 weeks.
"""
import warnings

import numpy as np
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing

from . import config
from .network import ramp_factor

SEASON = 52


def week_of_year(index: pd.DatetimeIndex) -> np.ndarray:
    return np.minimum((np.asarray(index.dayofyear) - 1) // 7 + 1, SEASON)


def future_weeks(last_week: pd.Timestamp, horizon: int) -> pd.DatetimeIndex:
    return pd.date_range(last_week + pd.Timedelta(weeks=1), periods=horizon, freq="W-SAT")


def seasonal_naive(train: pd.Series, horizon: int) -> np.ndarray:
    growth = train.iloc[-13:].sum() / train.iloc[-SEASON - 13:-SEASON].sum()
    return train.to_numpy()[-SEASON:][:horizon] * growth


def seasonal_indices(series: pd.Series) -> pd.Series:
    """Week-of-year index from the ratio of each week to its centered 52-week average."""
    trend = series.rolling(SEASON, center=True).mean()
    ratio = (series / trend).dropna()
    index = ratio.groupby(week_of_year(ratio.index)).mean()
    index = index.reindex(range(1, SEASON + 1)).interpolate(limit_direction="both")
    return index / index.mean()


def seasonal_index_model(train: pd.Series, horizon: int) -> np.ndarray:
    si = seasonal_indices(train)
    deseasonalized = train / si.loc[week_of_year(train.index)].to_numpy()
    slope = np.polyfit(np.arange(SEASON), deseasonalized.iloc[-SEASON:].to_numpy(), 1)[0]
    level = deseasonalized.iloc[-8:].mean() + slope * 3.5   # 8-week mean sits 3.5 weeks back
    future = future_weeks(train.index[-1], horizon)
    steps = np.arange(1, horizon + 1)
    return (level + slope * steps) * si.loc[week_of_year(future)].to_numpy()


def holt_winters(train: pd.Series, horizon: int) -> np.ndarray:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        fit = ExponentialSmoothing(
            train.to_numpy(), trend="add", damped_trend=True, seasonal="mul",
            seasonal_periods=SEASON, initialization_method="estimated",
        ).fit()
    return fit.forecast(horizon)


MODELS = {
    "Seasonal naive": seasonal_naive,
    "Seasonal index": seasonal_index_model,
    "Holt-Winters": holt_winters,
}


def pool_point_series(pool_weeks: pd.DataFrame) -> pd.DataFrame:
    """Weekly pallet demand with one column per pool point."""
    return pool_weeks.pivot(index="week_ending", columns="pool_point_id", values="demand_pallets").sort_index()


def wape(actual, forecast) -> float:
    actual, forecast = np.asarray(actual, float), np.asarray(forecast, float)
    return float(np.abs(actual - forecast).sum() / actual.sum())


def bias(actual, forecast) -> float:
    actual, forecast = np.asarray(actual, float), np.asarray(forecast, float)
    return float((forecast - actual).sum() / actual.sum())


def backtest(series: pd.DataFrame, origins=config.BACKTEST_ORIGINS, weeks: int = config.BACKTEST_WEEKS) -> dict:
    """Pretend it is each origin date, forecast the next `weeks` weeks, and score.

    Each origin is a number of weeks before the end of history. Results from all
    windows are pooled, so the score covers both peak and off-peak weeks.
    """
    windows = []
    for back in origins:
        cut = len(series) - back
        windows.append((series.iloc[:cut], series.iloc[cut:cut + weeks]))
    actual = pd.concat([test for _, test in windows])

    predictions, summary, by_pool_point = {}, [], []
    for name, model in MODELS.items():
        pred = pd.concat([
            pd.DataFrame({pp: model(train[pp], len(test)) for pp in series.columns}, index=test.index)
            for train, test in windows
        ])
        predictions[name] = pred
        summary.append({
            "model": name,
            "pool_point_wape": wape(actual.to_numpy().ravel(), pred.to_numpy().ravel()),
            "network_wape": wape(actual.sum(axis=1), pred.sum(axis=1)),
            "bias": bias(actual.to_numpy().ravel(), pred.to_numpy().ravel()),
        })
        for pp in series.columns:
            by_pool_point.append({"model": name, "pool_point_id": pp,
                                  "wape": wape(actual[pp], pred[pp]), "bias": bias(actual[pp], pred[pp])})
    summary = pd.DataFrame(summary).sort_values("pool_point_wape").reset_index(drop=True)
    champion = summary.loc[0, "model"]
    ratios = (actual / predictions[champion]).to_numpy().ravel()
    return {
        "windows": [(test.index[0], test.index[-1]) for _, test in windows],
        "actual": actual,
        "predictions": predictions,
        "summary": summary,
        "by_pool_point": pd.DataFrame(by_pool_point),
        "champion": champion,
        "ratio_quantiles": {q: float(np.quantile(ratios, q)) for q in (0.10, config.HIGH_SCENARIO_QUANTILE, 0.90)},
    }


def forecast_pool_points(series: pd.DataFrame, stores: pd.DataFrame, champion: str,
                         ratio_quantiles: dict, horizon: int = config.FORECAST_WEEKS) -> pd.DataFrame:
    """Forecast each pool point with the champion method, then add stores opening in the horizon."""
    model = MODELS[champion]
    future = future_weeks(series.index[-1], horizon)
    base = pd.DataFrame({pp: model(series[pp], horizon) for pp in series.columns}, index=future)

    new_store_volume = pd.DataFrame(0.0, index=future, columns=series.columns)
    opening = stores[stores["open_date"] > series.index[-1]]
    for store in opening.itertuples():
        si = seasonal_indices(series[store.pool_point_id]).loc[week_of_year(future)].to_numpy()
        new_store_volume[store.pool_point_id] += store.base_weekly_pallets * si * ramp_factor(store.open_date, future)

    rows = []
    q_low, q_high, q_p90 = (ratio_quantiles[q] for q in (0.10, config.HIGH_SCENARIO_QUANTILE, 0.90))
    for pp in series.columns:
        fc = base[pp] + new_store_volume[pp]
        rows.append(pd.DataFrame({
            "week_ending": future,
            "pool_point_id": pp,
            "base_forecast_pallets": base[pp].round(1).to_numpy(),
            "new_store_pallets": new_store_volume[pp].round(1).to_numpy(),
            "forecast_pallets": fc.round(1).to_numpy(),
            "low_pallets": (fc * q_low).round(1).to_numpy(),
            "high_pallets": (fc * q_high).round(1).to_numpy(),
            "p90_pallets": (fc * q_p90).round(1).to_numpy(),
        }))
    return pd.concat(rows, ignore_index=True)


def forecast_stores(store_weeks: pd.DataFrame, stores: pd.DataFrame,
                    pool_forecast: pd.DataFrame) -> pd.DataFrame:
    """Split each pool point's base forecast across its stores by recent share of volume."""
    last_week = store_weeks["week_ending"].max()
    recent = store_weeks[store_weeks["week_ending"] > last_week - pd.Timedelta(weeks=13)]
    volume = recent.groupby(["pool_point_id", "store_id"])["demand_pallets"].sum()
    share = (volume / volume.groupby(level="pool_point_id").transform("sum")).rename("share").reset_index()

    existing = share.merge(pool_forecast[["week_ending", "pool_point_id", "base_forecast_pallets"]], on="pool_point_id")
    existing["forecast_pallets"] = existing["share"] * existing["base_forecast_pallets"]

    future = pool_forecast["week_ending"].drop_duplicates().sort_values()
    future = pd.DatetimeIndex(future)
    new_rows = []
    for store in stores[stores["open_date"] > last_week].itertuples():
        pp_hist = store_weeks[store_weeks["pool_point_id"] == store.pool_point_id]
        series = pp_hist.groupby("week_ending")["demand_pallets"].sum()
        si = seasonal_indices(series).loc[week_of_year(future)].to_numpy()
        volume = store.base_weekly_pallets * si * ramp_factor(store.open_date, future)
        new_rows.append(pd.DataFrame({"week_ending": future, "pool_point_id": store.pool_point_id,
                                      "store_id": store.store_id, "forecast_pallets": volume}))
    out = pd.concat([existing[["week_ending", "pool_point_id", "store_id", "forecast_pallets"]], *new_rows],
                    ignore_index=True)
    out["forecast_pallets"] = out["forecast_pallets"].round(1)
    return out.sort_values(["week_ending", "store_id"]).reset_index(drop=True)
