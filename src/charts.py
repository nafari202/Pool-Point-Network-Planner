"""Static charts for the README and notebook."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm  # noqa: E402
from matplotlib.ticker import FuncFormatter, PercentFormatter  # noqa: E402
from matplotlib.transforms import blended_transform_factory  # noqa: E402

from . import config  # noqa: E402

# Validated reference palette (categorical slots 1-3 pass all-pairs CVD checks)
BLUE, ORANGE, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
INK, INK_2, MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, AXIS, SURFACE = "#e1e0d9", "#c3c2b7", "#fcfcfb"
GOOD, CRITICAL = "#0ca30c", "#d03b3b"
BLUE_LIGHT = "#cde2fb"
DIVERGING = LinearSegmentedColormap.from_list("blue_gray_red", ["#184f95", "#86b6ef", "#f0efec", "#ef9a99", "#c02f2f"])

plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "font.family": ["Segoe UI", "DejaVu Sans", "sans-serif"], "font.size": 10,
    "axes.edgecolor": AXIS, "axes.labelcolor": INK_2, "axes.titlecolor": INK,
    "axes.titlesize": 12, "axes.titleweight": "bold", "axes.titlelocation": "left",
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.axisbelow": True, "grid.color": GRID, "grid.linewidth": 0.8,
    "xtick.color": MUTED, "ytick.color": MUTED, "xtick.labelcolor": INK_2, "ytick.labelcolor": INK_2,
    "legend.frameon": False, "legend.labelcolor": INK_2, "lines.linewidth": 2,
})

thousands = FuncFormatter(lambda v, _: f"{v:,.0f}")


def _save(fig, name):
    path = config.IMAGE_DIR / name
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _subtitle(ax, text):
    ax.text(0, 1.02, text, transform=ax.transAxes, color=INK_2, fontsize=9, va="bottom")


def demand_signal(sales: pd.DataFrame, index: pd.Series):
    fig, axes = plt.subplots(2, 1, figsize=(10, 6.4), gridspec_kw={"hspace": 0.55})
    recent = sales[sales["month"] >= "2019-01-01"]
    ax = axes[0]
    ax.plot(recent["month"], recent["sales_musd"] / 1000, color=BLUE)
    ax.set_title("U.S. clothing store sales, monthly ($ billions)", pad=22)
    _subtitle(ax, "Census Monthly Retail Trade, not seasonally adjusted (FRED MRTSSM448USN)")
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}B"))

    ax = axes[1]
    ax.plot(index.index, index.values, color=BLUE)
    ax.axhline(1.0, color=AXIS, linewidth=1)
    ax.set_title("Weekly store shipment index built from it", pad=22)
    _subtitle(ax, "1.0 = average week. Shipments lead sales by one week.")
    peak = index.idxmax()
    ax.annotate(f"Holiday peak {index.max():.2f}x", xy=(peak, index.max()), xytext=(10, -4),
                textcoords="offset points", color=INK_2, fontsize=9)
    for a in axes:
        a.xaxis.set_major_locator(mdates.YearLocator())
        a.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    return _save(fig, "01_demand_signal.png")


def backtest(bt: dict):
    actual = bt["actual"].sum(axis=1)
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(actual.index, actual.values, color=INK, linewidth=2, label="Actual")
    colors = {"Holt-Winters": BLUE, "Seasonal naive": ORANGE, "Seasonal index": AQUA}
    scores = bt["summary"].set_index("model")["pool_point_wape"]
    for name, pred in bt["predictions"].items():
        total = pred.sum(axis=1)
        # Draw each backtest window separately so the two windows don't join
        for start, end in bt["windows"]:
            seg = total.loc[start:end]
            ax.plot(seg.index, seg.values, color=colors[name], linewidth=1.6,
                    label=f"{name} (WAPE {scores[name]:.1%})" if start == bt["windows"][0][0] else None)
    ax.axvline(bt["windows"][1][0], color=AXIS, linewidth=1, linestyle=":")
    ax.set_title("Backtest: network pallets per week, forecast vs. actual", pad=22)
    _subtitle(ax, "Each method forecast 26 weeks blind from Jul 2025 and from Jan 2026. WAPE pooled across pool points.")
    ax.yaxis.set_major_formatter(thousands)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.legend(loc="upper right", fontsize=9)
    return _save(fig, "02_backtest.png")


def forecast(history: pd.Series, network_fc: pd.DataFrame):
    hist = history.iloc[-104:]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(hist.index, hist.values, color=BLUE, label="Actual demand")
    ax.fill_between(network_fc.index, network_fc["low_pallets"], network_fc["p90_pallets"],
                    color=ORANGE, alpha=0.18, linewidth=0, label="P10 to P90 range")
    ax.plot(network_fc.index, network_fc["forecast_pallets"], color=ORANGE, label="Forecast")
    peak = network_fc["forecast_pallets"].idxmax()
    ax.annotate(f"Peak week {peak:%b %d}: {network_fc.loc[peak, 'forecast_pallets']:,.0f} pallets",
                xy=(peak, network_fc.loc[peak, "forecast_pallets"]), xytext=(-190, 6),
                textcoords="offset points", color=INK_2, fontsize=9)
    ax.set_title("Network volume forecast, next 26 weeks", pad=22)
    _subtitle(ax, "Pallets per week across 8 pool points, Holt-Winters plus two scheduled store openings")
    ax.yaxis.set_major_formatter(thousands)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.legend(loc="upper left", fontsize=9)
    return _save(fig, "03_forecast.png")


def capacity_heatmap(lanes: pd.DataFrame):
    lanes = lanes.assign(gap=lanes["trailers_forecast"] - lanes["contracted_trailers_per_week"],
                         lane=lanes["pool_point_id"] + " " + lanes["city"])
    grid = lanes.pivot(index="lane", columns="week_ending", values="gap")
    fig, ax = plt.subplots(figsize=(11, 4.2))
    limit = max(abs(grid.min().min()), abs(grid.max().max()))
    im = ax.imshow(grid.values, aspect="auto", cmap=DIVERGING, norm=TwoSlopeNorm(0, -limit, limit))
    ax.set_yticks(range(len(grid.index)), grid.index)
    ax.set_xticks(range(0, len(grid.columns), 2), [d.strftime("%b %d") for d in grid.columns[::2]],
                  rotation=45, ha="right")
    for (r, c), v in np.ndenumerate(grid.values):
        if v > 0:
            ax.text(c, r, f"+{int(v)}", ha="center", va="center", fontsize=7, color="white" if v > limit / 2 else INK)
    ax.grid(False)
    ax.set_title("Trailers needed vs. contracted capacity, by lane and week", pad=22)
    _subtitle(ax, "Red cells need surge trailers (number shown). Blue cells have spare contracted capacity.")
    cbar = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.01)
    cbar.set_label("Forecast trailers minus contracted", color=INK_2)
    cbar.outline.set_visible(False)
    return _save(fig, "04_capacity_heatmap.png")


def kpi_scorecard(card: pd.DataFrame):
    card = card.sort_values("pool_point_id", ascending=False)
    labels = card["pool_point_id"] + " " + card["city"]
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.2), gridspec_kw={"wspace": 0.75})
    for ax, col, target, title in [
        (axes[0], "on_time_pct", config.TARGETS["on_time_pct"], "On-time delivery"),
        (axes[1], "trailer_fill_pct", config.TARGETS["trailer_fill_pct"], "Trailer fill"),
    ]:
        meets = card[col] >= target
        colors = np.where(meets, GOOD, CRITICAL)
        ax.barh(labels, card[col], color=colors, height=0.6)
        ax.axvline(target, color=INK, linewidth=1.2, linestyle="--")
        ax.text(target, len(card) - 0.35, f" target {target:.0%}", color=INK, fontsize=9, va="bottom")
        # Values sit in a column right of the plot so they never collide with bars or the target line
        label_x = blended_transform_factory(ax.transAxes, ax.transData)
        for y, (v, ok) in enumerate(zip(card[col], meets)):
            ax.text(1.02, y, f"{v:.1%}  {'✓ meets' if ok else '✗ below'}", transform=label_x,
                    va="center", fontsize=9, color=INK_2)
        lo = min(card[col].min(), target) - 0.03
        hi = max(card[col].max(), target) + 0.01
        ax.set_xlim(lo, hi)
        ax.xaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
        ax.grid(axis="y", visible=False)
        ax.set_title(title, pad=10)
    fig.suptitle("Pool point scorecard, trailing 52 weeks", x=0.06, ha="left", fontsize=13, fontweight="bold", color=INK, y=1.03)
    return _save(fig, "05_kpi_scorecard.png")


def schedule_scenarios(scenarios: pd.DataFrame, plan_week):
    fig, axes = plt.subplots(1, 2, figsize=(11, 4), gridspec_kw={"wspace": 0.35})
    names = scenarios["scenario"].str.replace(r"^\d\. ", "", regex=True)
    for ax, col, title, fmt in [
        (axes[0], "backlog_pallets", "Backlog pallets left at week end", "{:,.0f}"),
        (axes[1], "total_cost", "Total weekly transportation cost", "${:,.0f}"),
    ]:
        ax.bar(names, scenarios[col], color=BLUE, width=0.55)
        for x, (v, cpp) in enumerate(zip(scenarios[col], scenarios["cost_per_pallet"])):
            label = fmt.format(v) if col == "backlog_pallets" else f"{fmt.format(v)}\n${cpp:.2f} per pallet"
            ax.text(x, v, label, ha="center", va="bottom", fontsize=9, color=INK_2)
        ax.set_title(title, pad=10)
        ax.grid(axis="x", visible=False)
        ax.margins(y=0.15)
        ax.yaxis.set_major_formatter(thousands if col == "backlog_pallets"
                                     else FuncFormatter(lambda v, _: f"${v / 1000:,.0f}K" if v else "$0"))
        ax.tick_params(axis="x", labelsize=9)
    fig.suptitle(f"Peak week plan, week ending {plan_week:%b %d, %Y}", x=0.06, ha="left",
                 fontsize=13, fontweight="bold", color=INK, y=1.04)
    return _save(fig, "06_schedule_scenarios.png")
