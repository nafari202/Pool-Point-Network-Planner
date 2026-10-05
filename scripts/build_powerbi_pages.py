"""Build the Power BI report pages (PBIR files) for the pool point model.

    python scripts/build_powerbi_pages.py

Rewrites every page under powerbi/Pool-Point-Network.Report/definition/pages, then validates
the project with Microsoft's powerbi-report-author CLI if it is installed
(npm install -g @microsoft/powerbi-report-authoring-cli). Close Power BI Desktop first, or
reload the report afterwards: Desktop overwrites the pages with its own copy when it saves.
"""
import datetime
import json
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROJECT = ROOT / "powerbi" / "Pool-Point-Network.pbip"
DEFINITION = ROOT / "powerbi" / "Pool-Point-Network.Report" / "definition"
SUMMARY = ROOT / "outputs" / "summary.json"

BASE_URL = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/"
VC_SCHEMA = BASE_URL + "visualContainer/2.12.0/schema.json"
PAGE_SCHEMA = BASE_URL + "page/2.1.0/schema.json"
PAGES_SCHEMA = BASE_URL + "pagesMetadata/1.1.0/schema.json"

KM = "Key Measures"
RED, BLUE, AMBER = "#B42318", "#1F4E79", "#C77700"
W, H, M = 1920, 1080, 32  # page size and margin


# ---------- expression helpers ----------
def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def s(text):  # string literal for formatting properties
    return lit("'" + text.replace("'", "''") + "'")


def col(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def meas(prop, entity=KM):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def agg(entity, prop, fn=0):
    return {"Aggregation": {"Expression": col(entity, prop), "Function": fn}}


def ref(field):
    """queryRef / nativeQueryRef for a field, matching what Desktop writes."""
    if "Aggregation" in field:
        c = field["Aggregation"]["Expression"]["Column"]
        e, p = c["Expression"]["SourceRef"]["Entity"], c["Property"]
        return f"Sum({e}.{p})", f"Sum of {p}"
    kind = "Measure" if "Measure" in field else "Column"
    e, p = field[kind]["Expression"]["SourceRef"]["Entity"], field[kind]["Property"]
    return f"{e}.{p}", p


def proj(field, display=None):
    q, n = ref(field)
    p = {"field": field, "queryRef": q, "nativeQueryRef": n}
    if display:
        p["displayName"] = display
    return p


def color_by_measure(measure):
    return {"solid": {"color": {"expr": meas(measure)}}}


def solid(hex_color):
    return {"solid": {"color": s(hex_color)}}


ALL_ROWS = {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}


def filter_def(entity, condition_builder):
    src = {"Column": {"Expression": {"SourceRef": {"Source": "c"}}, "Property": None}}
    return {"Version": 2, "From": [{"Name": "c", "Entity": entity, "Type": 0}],
            "Where": [{"Condition": condition_builder(src)}]}


def in_values(entity, prop, values, negate=False):
    def build(src):
        src["Column"]["Property"] = prop
        cond = {"In": {"Expressions": [src], "Values": [[{"Literal": {"Value": f"'{v}'"}}] for v in values]}}
        return {"Not": {"Expression": cond}} if negate else cond
    return filter_def(entity, build)


def compare(entity, prop, kind, literal):
    def build(src):
        src["Column"]["Property"] = prop
        return {"Comparison": {"ComparisonKind": kind, "Left": src, "Right": {"Literal": {"Value": literal}}}}
    return filter_def(entity, build)


# ---------- visual builders ----------
class Page:
    def __init__(self, name, display):
        self.name, self.display, self.visuals = name, display, []

    def add(self, vname, x, y, w, h, visual, filters=None, title=None):
        container = {"$schema": VC_SCHEMA, "name": vname,
                     "position": {"x": x, "y": y, "z": len(self.visuals) * 1000, "width": w, "height": h,
                                  "tabOrder": len(self.visuals) * 1000},
                     "visual": visual}
        if title is not None:
            visual.setdefault("visualContainerObjects", {})["title"] = [
                {"properties": {"show": lit("true"), "text": s(title)}}]
        if filters:
            container["filterConfig"] = {"filters": filters}
        self.visuals.append(container)


def textbox(runs):
    """runs: list of paragraphs, each a list of (text, style dict)."""
    paragraphs = [{"textRuns": [{"value": t, "textStyle": st} for t, st in para]} for para in runs]
    return {"visualType": "textbox",
            "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]},
            "visualContainerObjects": {"background": [{"properties": {"show": lit("false")}}]}}


def card(measure, font=28, show_label=True, display=None):
    return {"visualType": "card",
            "query": {"queryState": {"Values": {"projections": [proj(meas(measure), display)]}}},
            "objects": {"labels": [{"properties": {"fontSize": lit(f"{font}D"), "labelDisplayUnits": lit("1D")}}],
                        "categoryLabels": [{"properties": {"show": lit("true" if show_label else "false")}}]},
            "drillFilterOtherVisuals": True}


def period_slicer():
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [proj(col("Calendar", "Period"))]}}},
            "objects": {"data": [{"properties": {"mode": s("Basic")}}],
                        "selection": [{"properties": {"singleSelect": lit("true")}}],
                        "general": [{"properties": {"filter": {"filter": in_values("Calendar", "Period", ["Last 52 weeks"])}}}]},
            "syncGroup": {"groupName": "Period", "fieldChanges": True, "filterChanges": True},
            "drillFilterOtherVisuals": True}


def exclude_forecast(page_name):
    return [{"name": f"{page_name}ExcludeForecast", "field": col("Calendar", "Period"), "type": "Categorical",
             "filter": in_values("Calendar", "Period", ["Forecast"], negate=True)}]


def table(fields, font_colors=None, sort=None):
    """fields: list of field dicts (or (field, displayName)); font_colors: {queryRef: color measure}."""
    projs = [proj(*f) if isinstance(f, tuple) else proj(f) for f in fields]
    v = {"visualType": "tableEx", "query": {"queryState": {"Values": {"projections": projs}}},
         "objects": {"values": [{"properties": {"fontSize": lit("13D")}}],
                     "columnHeaders": [{"properties": {"fontSize": lit("13D"), "bold": lit("true")}}],
                     "total": [{"properties": {"fontSize": lit("13D")}}],
                     "grid": [{"properties": {"textSize": lit("13D")}}]},
         "drillFilterOtherVisuals": True}
    for qref, m in (font_colors or {}).items():
        v["objects"]["values"].append({"properties": {"fontColor": color_by_measure(m)},
                                       "selector": {**ALL_ROWS, "metadata": qref}})
    if sort:
        field, direction = sort
        v["query"]["sortDefinition"] = {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}
    return v


def column_chart(category, measure, color_measure=None, target_measure=None, target_label=None):
    v = {"visualType": "clusteredColumnChart",
         "query": {"queryState": {"Category": {"projections": [proj(category)]},
                                  "Y": {"projections": [proj(meas(measure))]}},
                   "sortDefinition": {"sort": [{"field": category, "direction": "Ascending"}], "isDefaultSort": False}},
         "objects": {"labels": [{"properties": {"show": lit("true")}}]},
         "drillFilterOtherVisuals": True}
    if color_measure:
        v["objects"]["dataPoint"] = [{"properties": {"fill": color_by_measure(color_measure)}, "selector": ALL_ROWS}]
    if target_measure:
        v["objects"]["y1AxisReferenceLine"] = [{
            "properties": {"show": lit("true"), "value": {"expr": meas(target_measure)},
                           "lineColor": solid("#333333"), "style": s("dashed"), "transparency": lit("0D"),
                           "displayName": s(target_label or "Target"), "dataLabelShow": lit("true"),
                           "dataLabelText": s("Name")},
            "selector": {"id": "1"}}]
    return v


def line_chart(measures, colors, series=None, target_measure=None, target_label=None):
    state = {"Category": {"projections": [proj(col("Calendar", "Week Ending"))]},
             "Y": {"projections": [proj(meas(m)) for m in measures]}}
    if series:
        state["Series"] = {"projections": [proj(series)]}
    v = {"visualType": "lineChart", "query": {"queryState": state},
         "objects": {"legend": [{"properties": {"show": lit("true"), "position": s("Top")}}],
                     "lineStyles": [{"properties": {"lineChartType": s("linear")}}]},
         "drillFilterOtherVisuals": True}
    if colors:
        v["objects"]["dataPoint"] = [{"properties": {"fill": solid(c)}, "selector": {"metadata": f"{KM}.{m}"}}
                                     for m, c in zip(measures, colors)]
    if target_measure:
        v["objects"]["y1AxisReferenceLine"] = [{
            "properties": {"show": lit("true"), "value": {"expr": meas(target_measure)},
                           "lineColor": solid(RED), "style": s("dashed"),
                           "displayName": s(target_label or "Target"), "dataLabelShow": lit("true"),
                           "dataLabelText": s("Name")},
            "selector": {"id": "1"}}]
    return v


def header(page, title, subtitle_measure="As-Of Label", subtitle_text=None):
    page.add("pageTitle", M, 16, 1300, 56, textbox([[(title, {"fontSize": "24pt", "fontWeight": "bold", "color": "#1A1A1A"})]]))
    if subtitle_text:
        page.add("pageSubtitle", M, 70, 1300, 40, textbox([[(subtitle_text, {"fontSize": "12pt", "color": "#555555"})]]))
    else:
        page.add("asOfLabel", M, 66, 700, 48, card(subtitle_measure, font=12, show_label=False))


def card_row(page, prefix, measures, y, h=120, x0=M, total_w=W - 2 * M, gap=16, displays=None):
    w = (total_w - gap * (len(measures) - 1)) / len(measures)
    for i, m in enumerate(measures):
        page.add(f"{prefix}{i + 1}", round(x0 + i * (w + gap)), y, round(w), h, card(m, display=(displays or {}).get(m)))


# ---------- pages ----------
def scorecard_page():
    p = Page("networkScorecard", "Network Scorecard")
    header(p, "Pool Point Network Scorecard")
    p.add("periodSlicer", 1560, 12, 328, 130, period_slicer(), filters=exclude_forecast(p.name))
    card_row(p, "kpiCard", ["Avg Weekly Pallets", "YoY Volume Growth", "On-Time %", "Trailer Fill %", "Cost per Pallet"], y=156)
    p.add("scorecardTable", M, 296, W - 2 * M, 430, table(
        [col("Pool Points", "Pool Point"), col("Carriers", "Carrier"), meas("Avg Weekly Pallets"),
         meas("YoY Volume Growth"), meas("On-Time %"), meas("Trailer Fill %"), meas("Cost per Pallet"),
         meas("Cost Index vs Network"), meas("Avg Weekly Backlog"), meas("Forecast WAPE")],
        font_colors={f"{KM}.On-Time %": "On-Time Color", f"{KM}.Trailer Fill %": "Fill Color",
                     f"{KM}.Cost Index vs Network": "Cost Color"},
        sort=(col("Pool Points", "Pool Point"), "Ascending")),
        title="Pool point scorecard - red = misses target (on-time 95%, fill 75%, cost more than 1.10x network)")
    p.add("onTimeTrend", M, 742, 920, 314, line_chart(["On-Time %"], [BLUE], target_measure="On-Time Target",
                                                      target_label="Target 95%"), title="Network on-time % by week")
    p.add("fillTrend", 968, 742, 920, 314, line_chart(["Trailer Fill %"], [BLUE], target_measure="Trailer Fill Target",
                                                      target_label="Target 75%"), title="Network trailer fill % by week")
    return p


def carrier_page():
    p = Page("carrierPerformance", "Carrier Performance")
    header(p, "Carrier Performance")
    p.add("periodSlicer", 1560, 12, 328, 130, period_slicer(), filters=exclude_forecast(p.name))
    carrier = col("Carriers", "Carrier")
    p.add("onTimeByCarrier", M, 156, 920, 340, column_chart(carrier, "On-Time %", "On-Time Color", "On-Time Target", "Target 95%"),
          title="On-time % by carrier")
    p.add("fillByCarrier", 968, 156, 920, 340, column_chart(carrier, "Trailer Fill %", "Fill Color", "Trailer Fill Target", "Target 75%"),
          title="Trailer fill % by carrier")
    p.add("carrierTable", M, 512, W - 2 * M, 240, table(
        [carrier, meas("Avg Weekly Pallets"), meas("Trailers"), meas("On-Time %"), meas("On-Time Status"),
         meas("Trailer Fill %"), meas("Cost per Pallet"), meas("Total Cost")],
        font_colors={f"{KM}.On-Time %": "On-Time Color", f"{KM}.On-Time Status": "On-Time Color",
                     f"{KM}.Trailer Fill %": "Fill Color"},
        sort=(carrier, "Ascending")), title="Carrier scorecard")
    p.add("onTimeTrendByCarrier", M, 768, W - 2 * M, 288, line_chart(["On-Time %"], None, series=carrier),
          title="On-time % by week and carrier")
    return p


def forecast_page():
    p = Page("forecastCapacity", "Forecast & Capacity")
    header(p, "Forecast & Lane Capacity")
    card_row(p, "fcCard", ["Forecast Pallets", "Champion Model", "Forecast WAPE", "Lane-Weeks Over Capacity",
                           "Extra Trailers Needed"], y=126, h=118,
             displays={"Forecast Pallets": "Forecast pallets, next 26 weeks", "Forecast WAPE": "Backtest WAPE (champion)",
                       "Extra Trailers Needed": "Extra trailers needed (high case)"})
    last_year = [{"name": "lastYearAndForecast", "field": col("Calendar", "Week Offset"), "type": "Advanced",
                  "filter": compare("Calendar", "Week Offset", 2, "-52L")}]
    p.add("actualVsForecast", M, 260, 1110, 796,
          line_chart(["Pallets Shipped", "Forecast Pallets", "Forecast High Pallets"], [BLUE, AMBER, "#E8B04B"]),
          filters=last_year, title="Weekly pallets - last 52 weeks actual, then 26-week forecast (high = 80th percentile)")
    matrix = {"visualType": "pivotTable",
              "query": {"queryState": {"Rows": {"projections": [proj(col("Calendar", "Week Ending"))]},
                                       "Columns": {"projections": [proj(col("Pool Points", "Pool Point ID"))]},
                                       "Values": {"projections": [proj(meas("Trailers Forecast"))]}}},
              "objects": {"values": [{"properties": {"backColor": color_by_measure("Capacity Color")},
                                      "selector": {**ALL_ROWS, "metadata": f"{KM}.Trailers Forecast"}}],
                          "subTotals": [{"properties": {"rowSubtotals": lit("false"), "columnSubtotals": lit("false")}}],
                          "grid": [{"properties": {"textSize": lit("12D")}}],
                          "columnHeaders": [{"properties": {"fontSize": lit("12D"), "bold": lit("true")}}],
                          "rowHeaders": [{"properties": {"fontSize": lit("12D")}}]},
              "drillFilterOtherVisuals": True}
    p.add("capacityHeatmap", 1158, 260, 730, 796, matrix,
          title="Trailers needed per lane - red over contract, amber at risk in high case, green covered")
    return p


def peak_page():
    p = Page("peakWeekPlan", "Peak Week Plan")
    week = datetime.date.fromisoformat(json.loads(SUMMARY.read_text())["plan_week"])
    header(p, "Peak Week Plan", subtitle_text=f"Plan week ending {week:%b} {week.day}, {week.year}"
                                             " - three scenarios, change requests, and store-level plan")
    scen = col("Peak Week Scenarios", "Scenario")
    combo = {"visualType": "lineClusteredColumnComboChart",
             "query": {"queryState": {"Category": {"projections": [proj(scen)]},
                                      "Y": {"projections": [proj(agg("Peak Week Scenarios", "Backlog Pallets"), "Backlog pallets")]},
                                      "Y2": {"projections": [proj(agg("Peak Week Scenarios", "Total Cost"), "Total cost")]}},
                       "sortDefinition": {"sort": [{"field": scen, "direction": "Ascending"}], "isDefaultSort": False}},
             "objects": {"labels": [{"properties": {"show": lit("true")}}],
                         "legend": [{"properties": {"show": lit("true"), "position": s("Top")}}],
                         "lineStyles": [{"properties": {"lineChartType": s("linear")}}]},
             "drillFilterOtherVisuals": True}
    p.add("scenarioChart", M, 126, 1000, 380, combo, title="Peak week scenarios - backlog pallets vs total cost")
    gx, gw, gh = 1048, (W - M - 1048 - 16) / 2, 182
    for i, m in enumerate(["Peak Stores with Backlog", "Peak Backlog Pallets", "Change Requests", "Change Cost Impact"]):
        p.add(f"peakCard{i + 1}", round(gx + (i % 2) * (gw + 16)), 126 + (i // 2) * (gh + 16), round(gw), gh, card(m))
    sc = "Schedule Changes"
    p.add("changeLog", M, 522, W - 2 * M, 270, table(
        [col(sc, "Change ID"), col(sc, "Change Type"), col("Pool Points", "Pool Point"), col(sc, "Reason"),
         col(sc, "Backlog Change"), col(sc, "Deliveries Change"), col(sc, "Trailers Change"), col(sc, "Cost Change")],
        sort=(col(sc, "Change ID"), "Ascending")), title="Change requests and their peak-week impact")
    sp = "Peak Week Store Plan"
    p.add("storePlan", M, 808, W - 2 * M, 248, table(
        [col("Stores", "Store ID"), col("Pool Points", "Pool Point"), col(sp, "Peak Delivery Days"),
         col(sp, "Peak Pallet Max"), col(sp, "Peak Forecast Pallets"), col(sp, "Planned Pallets"), col(sp, "Peak Backlog Pallets")],
        sort=(col(sp, "Peak Backlog Pallets"), "Descending")), title="Store plan for the peak week (mitigated) - stores with backlog first")
    return p


# ---------- output ----------
def write_pages(definition_dir):
    pages = [scorecard_page(), carrier_page(), forecast_page(), peak_page()]
    pages_dir = definition_dir / "pages"
    if pages_dir.exists():
        shutil.rmtree(pages_dir)
    for pg in pages:
        page_dir = pages_dir / pg.name
        for v in pg.visuals:
            visual_dir = page_dir / "visuals" / v["name"]
            visual_dir.mkdir(parents=True)
            (visual_dir / "visual.json").write_text(json.dumps(v, indent=2), encoding="utf-8")
        page = {"$schema": PAGE_SCHEMA, "name": pg.name, "displayName": pg.display,
                "displayOption": "FitToPage", "height": H, "width": W}
        (page_dir / "page.json").write_text(json.dumps(page, indent=2), encoding="utf-8")
        print(f"{pg.display}: {len(pg.visuals)} visuals")
    order = {"$schema": PAGES_SCHEMA, "pageOrder": [p.name for p in pages], "activePageName": pages[0].name}
    (pages_dir / "pages.json").write_text(json.dumps(order, indent=2), encoding="utf-8")


def add_missing_schemas(project):
    """Desktop drops $schema from these two files on save; the validator requires it."""
    files = {project: "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
             DEFINITION.parent / "definition.pbir":
                 "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json"}
    for path, url in files.items():
        data = json.loads(path.read_text(encoding="utf-8"))
        if "$schema" not in data:
            path.write_text(json.dumps({"$schema": url, **data}, indent=2), encoding="utf-8")


def validate(project):
    cli = shutil.which("powerbi-report-author")
    if not cli:
        print("Skipping validation (install with: npm install -g @microsoft/powerbi-report-authoring-cli).")
        return
    result = subprocess.run([cli, "validate", str(project)], capture_output=True, text=True)
    try:
        report = json.loads(result.stdout).get("data", {})
    except json.JSONDecodeError:
        print(result.stdout.strip() or result.stderr.strip())
        return
    print(f"Validation {report.get('result')}: {report.get('errorCount')} errors, {report.get('warningCount')} warnings")
    if report.get("errorCount"):
        print(result.stdout)


if __name__ == "__main__":
    write_pages(DEFINITION)
    add_missing_schemas(PROJECT)
    validate(PROJECT)
