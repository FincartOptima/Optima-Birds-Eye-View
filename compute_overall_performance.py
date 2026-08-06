"""Overall Portfolio Performance — the entire book (all clients combined),
its full fund/asset-class allocation, and its TWRR-style performance against
every benchmark index carried in the equity Google Sheet.

Methodology note: this tab answers "if I held today's exact fund mix, how did
it perform against the market" — a time-weighted, current-holdings-projected
view (same convention as the Master Dashboard's fund/category performance
table). It intentionally does NOT use money-weighted XIRR: that measures each
client's actual cashflow-timed return and already lives on the per-client
Performance tab. Mixing the two methodologies on one screen would make the
portfolio-vs-benchmark comparison unfair, since the benchmarks themselves are
priced point-to-point, not cashflow-matched.

Portfolio inception is a fixed date (the strategy's launch date), not each
client's individual first-deposit date — matching how a single-strategy
factsheet reports "since inception".
"""
from __future__ import annotations

from bisect import bisect_right
from datetime import datetime, timedelta

import gsheet_data
from create_client_factsheet_report import CATEGORY_ORDER

PORTFOLIO_INCEPTION = datetime(2026, 4, 13)

PERIOD_DAYS: list[tuple[str, int]] = [
    ("1M", 30),
    ("3M", 91),
    ("6M", 182),
    ("1Y", 365),
]

BENCHMARK_COLUMNS: list[tuple[str, str]] = [
    ("BSE 500", "Benchmark"),
    ("Nifty 100", "Nifty100"),
    ("Small Cap 250", "SmallCap250"),
    ("Mid Cap 150", "MidCap150"),
    ("Nifty 50", "Nifty50"),
    ("Momentum 30", "Momentum30"),
    ("Quality 30", "Quality30"),
]


def _value_on_or_before(dates: list[datetime], values: list[float], target: datetime) -> float | None:
    pos = bisect_right(dates, target) - 1
    return values[pos] if pos >= 0 else None


def _point_to_point_return(dates: list[datetime], values: list[float], start: datetime, end: datetime) -> float | None:
    v_start = _value_on_or_before(dates, values, start)
    v_end = _value_on_or_before(dates, values, end)
    if v_start and v_end and v_start > 0:
        return round((v_end / v_start - 1) * 100, 2)
    return None


def _available_periods(report_date: datetime) -> list[tuple[str, int]]:
    """Only include periods that have actually elapsed since inception —
    a '6M' column would be meaningless (not wrong, just undefined) for a
    portfolio that has only existed for 4 months."""
    elapsed_days = (report_date - PORTFOLIO_INCEPTION).days
    return [(label, days) for label, days in PERIOD_DAYS if elapsed_days >= days]


def get_overall_performance(reports: list, bse_prices: list[tuple], report_date: datetime) -> dict:
    active_reports = [r for r in reports if r.cost_value > 0]

    # ---- Totals ------------------------------------------------------------
    total_cost = sum(r.cost_value for r in active_reports)
    total_current = sum(r.current_value for r in active_reports)
    total_pl = sum(r.total_pl for r in active_reports)
    simple_return = round(total_pl / total_cost * 100, 2) if total_cost else None

    # ---- Asset allocation (category weight across the whole book) --------
    cat_values: dict[str, float] = {cat: 0.0 for cat in CATEGORY_ORDER}
    for r in active_reports:
        for row in r.category_rows:
            cat_values[row["Category"]] = cat_values.get(row["Category"], 0.0) + row["Current Value"]
    asset_allocation = [
        {"category": cat, "value": val, "pct": round(val / total_current * 100, 2) if total_current else 0.0}
        for cat, val in cat_values.items() if val > 0
    ]

    # ---- Fund-level breakdown across the whole book -----------------------
    fund_values: dict[str, dict] = {}  # isin -> {scheme, category, value}
    for r in active_reports:
        for h in r.holdings:
            if h.current_value <= 0:
                continue
            key = h.isin or h.scheme_name
            entry = fund_values.setdefault(key, {"scheme": h.scheme_name, "isin": h.isin, "category": h.category, "value": 0.0})
            entry["value"] += h.current_value
    all_funds = sorted(
        (
            {**v, "pct": round(v["value"] / total_current * 100, 2) if total_current else 0.0}
            for v in fund_values.values()
        ),
        key=lambda x: x["value"], reverse=True,
    )
    top_holdings = all_funds[:5]

    # ---- Performance: portfolio (current-weight NAV projection) + benchmarks ----
    periods = _available_periods(report_date)
    period_labels = [label for label, _ in periods] + ["Since Inception"]

    mapping = gsheet_data.load_fund_mapping()
    isin_to_fund = {m["isin"]: m["sheet_fund_name"] for m in mapping if m["sheet_fund_name"]}
    nav_index = gsheet_data.get_live_nav_index()
    nav_history_by_isin = {
        isin: nav_index[fund_name] for isin, fund_name in isin_to_fund.items() if fund_name in nav_index
    }

    portfolio_returns: dict[str, float | None] = {}
    for label, days in periods:
        start = report_date - timedelta(days=days)
        portfolio_returns[label] = _weighted_portfolio_return(all_funds, total_current, nav_history_by_isin, start, report_date)
    portfolio_returns["Since Inception"] = _weighted_portfolio_return(
        all_funds, total_current, nav_history_by_isin, PORTFOLIO_INCEPTION, report_date
    )

    bse_dates = [p[0] for p in bse_prices]
    bse_values = [p[1] for p in bse_prices]
    equity_sheet_series = gsheet_data._parse_series(gsheet_data._get_rows("equity"))

    # A list, not a dict: JSON objects have no guaranteed key order (Flask's
    # default JSON provider sorts keys alphabetically), but this row order
    # must match the sheet's own benchmark column order.
    benchmark_returns: list[dict] = []
    for display_name, column_name in BENCHMARK_COLUMNS:
        if column_name == "Benchmark":
            dates, values = bse_dates, bse_values
        else:
            series = equity_sheet_series.get(column_name, {"dates": [], "values": []})
            dates, values = series["dates"], series["values"]
        row: dict[str, float | None] = {}
        for label, days in periods:
            start = report_date - timedelta(days=days)
            row[label] = _point_to_point_return(dates, values, start, report_date)
        row["Since Inception"] = _point_to_point_return(dates, values, PORTFOLIO_INCEPTION, report_date)
        benchmark_returns.append({"name": display_name, "returns": row})

    return {
        "report_date": report_date.strftime("%d %b %Y"),
        "inception_date": PORTFOLIO_INCEPTION.strftime("%d %b %Y"),
        "totals": {
            "n_clients": len(active_reports),
            "invested": total_cost,
            "current_value": total_current,
            "gain_loss": total_pl,
            "gain_loss_pct": simple_return,
        },
        "asset_allocation": asset_allocation,
        "all_funds": all_funds,
        "top_holdings": top_holdings,
        "performance": {
            "period_labels": period_labels,
            "portfolio": portfolio_returns,
            "benchmarks": benchmark_returns,
        },
    }


def _weighted_portfolio_return(all_funds, total_current, nav_history_by_isin, start, end) -> float | None:
    """Σ(current_weight_i × fund_i point-to-point return). Skips funds with no
    priceable NAV history over the window; renormalises weights over the
    funds that ARE priceable so a couple of unmapped funds don't zero out
    the whole figure."""
    weighted_sum = 0.0
    covered_weight = 0.0
    for f in all_funds:
        hist = nav_history_by_isin.get(f["isin"])
        if not hist or not hist["dates"]:
            continue
        ret = _point_to_point_return(hist["dates"], hist["values"], start, end)
        if ret is None:
            continue
        weight = f["value"] / total_current if total_current else 0.0
        weighted_sum += weight * ret
        covered_weight += weight
    if covered_weight <= 0:
        return None
    return round(weighted_sum / covered_weight, 2)
