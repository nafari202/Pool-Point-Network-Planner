"""End-to-end run: data, simulation, forecast, capacity, scheduling, KPIs, outputs."""
import json

import pandas as pd

from . import capacity, charts, config, demand_signal, forecast, kpis, network, report_excel, schedule, simulate


def _change_notes(changes, pool_points_before, plan, carriers, carrier_card):
    lanes_before = schedule.lane_costs(pool_points_before, carriers)["linehaul_cost_per_trailer"]
    on_time = carrier_card.set_index("carrier_id")["on_time_pct"]
    notes = []
    for ch in changes.itertuples():
        if ch.change_type == "CANCEL_POOL_POINT_DAY":
            text = f"All {ch.day} deliveries at {ch.pool_point_id} cancelled; pallets roll to each store's next delivery."
        elif ch.change_type == "CANCEL_STORE_DELIVERY":
            text = f"{ch.store_id} skips {ch.day}; its pallets roll to the next delivery."
        elif ch.change_type == "PALLET_MAX_CHANGE":
            text = f"{ch.store_id} pallet max {ch.old_value} to {ch.new_value} per delivery."
        elif ch.change_type == "NEW_STORE":
            text = f"{ch.store_id} added to {ch.pool_point_id} on {ch.delivery_days} with {int(ch.opening_pallets)} opening pallets."
        else:
            new_cost = carriers.set_index("carrier_id").loc[ch.new_value]
            miles = pool_points_before.set_index("pool_point_id").at[ch.pool_point_id, "miles_from_dc"]
            new_lane = new_cost["fixed_charge"] + miles * new_cost["rate_per_mile"]
            text = (f"Lane cost ${lanes_before[ch.pool_point_id]:,.0f} to ${new_lane:,.0f} per trailer; "
                    f"carrier on-time {on_time[ch.old_value]:.1%} vs {on_time[ch.new_value]:.1%} over the last 52 weeks.")
        notes.append(text)
    plan["per_change"]["notes"] = notes


def run(make_charts: bool = True, make_workbook: bool = True, refresh_data: bool = False) -> dict:
    for folder in (config.DATA_DIR, config.OUTPUT_DIR, config.IMAGE_DIR):
        folder.mkdir(parents=True, exist_ok=True)

    # 1. Real demand signal
    sales = demand_signal.load_census_sales(refresh=refresh_data)
    weeks = demand_signal.history_weeks()
    index = demand_signal.weekly_demand_index(sales, weeks)

    # 2. Simulated network and four years of operations
    net = network.build_network()
    stores, carriers, rng = net["stores"], net["carriers"], net["rng"]
    store_weeks = simulate.simulate_store_weeks(stores, index, rng)
    pool_weeks = simulate.simulate_pool_point_weeks(store_weeks, stores, net["pool_points"], carriers, index, rng)
    pool_points = simulate.set_contracted_capacity(net["pool_points"], pool_weeks)

    # 3. Forecast: backtest, pick the champion, forecast 26 weeks
    series = forecast.pool_point_series(pool_weeks)
    bt = forecast.backtest(series)
    pool_fc = forecast.forecast_pool_points(series, stores, bt["champion"], bt["ratio_quantiles"])
    store_fc = forecast.forecast_stores(store_weeks, stores, pool_fc)

    # 4. Capacity outlook for carriers
    lanes = capacity.lane_outlook(pool_fc, pool_points)
    carrier_weeks = capacity.carrier_outlook(lanes, carriers)
    notes = capacity.carrier_notes(carrier_weeks, lanes)

    # 5. KPIs
    card = kpis.pool_point_scorecard(pool_weeks, pool_points, bt["by_pool_point"], bt["champion"])
    carrier_card = kpis.carrier_scorecard(pool_weeks, carriers)
    network_kpis = kpis.network_weekly(pool_weeks)

    # 6. Peak week schedule with change requests
    network_fc = pool_fc.groupby("week_ending")[["forecast_pallets", "low_pallets", "high_pallets", "p90_pallets"]].sum()
    plan_week = network_fc["forecast_pallets"].idxmax()
    changes_path = config.DATA_DIR / "schedule_changes.csv"
    if changes_path.exists():
        changes = pd.read_csv(changes_path, dtype=str)
        changes["opening_pallets"] = pd.to_numeric(changes["opening_pallets"])
    else:
        changes = schedule.default_changes(stores, pool_points)
        changes.to_csv(changes_path, index=False)
    plan = schedule.run_week(stores, pool_points, carriers, store_fc, plan_week, changes)
    _change_notes(changes, pool_points, plan, carriers, carrier_card)

    as_of = series.index[-1]
    champion_wape = float(bt["summary"].set_index("model").at[bt["champion"], "pool_point_wape"])

    # 7. Save data and outputs
    stores.to_csv(config.DATA_DIR / "stores.csv", index=False)
    pool_points.to_csv(config.DATA_DIR / "pool_points.csv", index=False)
    carriers.to_csv(config.DATA_DIR / "carriers.csv", index=False)
    pool_weeks.to_csv(config.DATA_DIR / "pool_point_weekly_history.csv", index=False)
    store_weeks.to_csv(config.DATA_DIR / "store_weekly_history.csv", index=False)
    index.rename_axis("week_ending").to_csv(config.DATA_DIR / "weekly_demand_index.csv")

    out = config.OUTPUT_DIR
    bt["summary"].to_csv(out / "backtest_summary.csv", index=False)
    bt["by_pool_point"].to_csv(out / "backtest_by_pool_point.csv", index=False)
    pool_fc.to_csv(out / "pool_point_forecast.csv", index=False)
    store_fc.to_csv(out / "store_forecast.csv", index=False)
    lanes.to_csv(out / "lane_capacity_outlook.csv", index=False)
    carrier_weeks.to_csv(out / "carrier_outlook.csv", index=False)
    card.to_csv(out / "kpi_scorecard.csv", index=False)
    carrier_card.to_csv(out / "carrier_scorecard.csv", index=False)
    plan["scenarios"].to_csv(out / "peak_week_scenarios.csv", index=False)
    plan["per_change"].to_csv(out / "schedule_change_impact.csv", index=False)
    plan["store_visibility"].to_csv(out / "store_visibility_peak_week.csv", index=False)
    (out / "carrier_volume_outlook.md").write_text(notes, encoding="utf-8")

    scen = plan["scenarios"].set_index("scenario")
    with_changes, mitigated = scen.iloc[1], scen.iloc[2]
    summary = {
        "as_of_week": f"{as_of:%Y-%m-%d}",
        "plan_week": f"{plan_week:%Y-%m-%d}",
        "pool_points": int(len(pool_points)),
        "stores_open_at_as_of": int((stores["open_date"] <= as_of).sum()),
        "stores_opening_in_forecast": int((stores["open_date"] > as_of).sum()),
        "history_weeks": int(len(weeks)),
        "store_week_rows": int(len(store_weeks)),
        "backtest": bt["summary"].round(4).to_dict(orient="records"),
        "backtest_windows": [[f"{a:%Y-%m-%d}", f"{b:%Y-%m-%d}"] for a, b in bt["windows"]],
        "champion": bt["champion"],
        "forecast_peak_pallets": round(float(network_fc["forecast_pallets"].max())),
        "forecast_avg_weekly_pallets": round(float(network_fc["forecast_pallets"].mean())),
        "lane_weeks_over_capacity": int((lanes["capacity_status"] == "Over capacity").sum()),
        "lane_weeks_total": int(len(lanes)),
        "max_extra_trailers_one_carrier_week": int(carrier_weeks["extra_trailers_needed"].max()),
        "total_surge_trailers_high_case": int(lanes["extra_trailers_needed"].sum()),
        "peak_week": plan["scenarios"].round(4).to_dict(orient="records"),
        "backlog_cleared_pct": round(1 - mitigated["backlog_pallets"] / with_changes["backlog_pallets"], 4),
        "mitigation_cost_change_pct": round(mitigated["total_cost"] / with_changes["total_cost"] - 1, 4),
        "stores_mitigated": int(len(plan["mitigation_actions"])),
        "pool_points_below_on_time": card.loc[card["on_time_status"] != "Meets", "pool_point_id"].tolist(),
        "pool_points_below_fill": card.loc[card["fill_status"] != "Meets", "pool_point_id"].tolist(),
        "pool_points_high_cost": card.loc[card["cost_status"] != "In range", "pool_point_id"].tolist(),
    }
    (out / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")

    results = {
        "sales": sales, "demand_index": index, "stores": stores, "pool_points": pool_points, "carriers": carriers,
        "store_weeks": store_weeks, "pool_weeks": pool_weeks, "series": series, "backtest": bt,
        "pool_forecast": pool_fc, "store_forecast": store_fc, "network_forecast": network_fc, "lanes": lanes,
        "carrier_weeks": carrier_weeks, "carrier_notes": notes, "scorecard": card, "carrier_scorecard": carrier_card,
        "network_kpis": network_kpis, "changes": changes, "plan": plan, "summary": summary,
    }

    if make_charts:
        charts.demand_signal(sales, index)
        charts.backtest(bt)
        charts.forecast(series.sum(axis=1), network_fc)
        charts.capacity_heatmap(lanes)
        charts.kpi_scorecard(card)
        charts.schedule_scenarios(plan["scenarios"], plan_week)

    if make_workbook:
        report_excel.build_workbook(out / "pool_point_operations_report.xlsx", {
            "as_of": as_of, "plan_week": plan_week, "champion": bt["champion"], "wape": champion_wape,
            "carriers": carriers, "pool_points": pool_points, "stores": stores, "stores_all": stores,
            "pool_forecast": pool_fc, "store_forecast": store_fc, "carrier_weeks": carrier_weeks,
            "carrier_notes": notes, "scorecard": card, "carrier_scorecard": carrier_card, "plan": plan,
            "changes": changes, "pool_weeks": pool_weeks,
        })
    return results
