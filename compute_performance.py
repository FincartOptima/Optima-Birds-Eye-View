"""Client portfolio performance vs BSE 500 — simple returns since inception.

Client simple return uses total P/L over invested cost (same formula the
Consolidated tab uses), NOT the trade master's "Initial Investment" field,
which is a stale one-time value that ignores later deposits.

BSE 500 simple return is point-to-point from the client's inception date to
today. NAV/benchmark history is sourced live from Google Sheets, so both
figures self-update as those sheets get new data each day.
"""
from __future__ import annotations

from bisect import bisect_right
from datetime import datetime


def _value_on_or_before(dates: list[datetime], values: list[float], target: datetime) -> float | None:
    pos = bisect_right(dates, target) - 1
    return values[pos] if pos >= 0 else None


def _bse_simple_return_from_inception(bse_prices: list[tuple], inception_date: datetime, report_date: datetime) -> float | None:
    if not bse_prices or inception_date is None:
        return None
    dates = [p[0] for p in bse_prices]
    values = [p[1] for p in bse_prices]
    v_start = _value_on_or_before(dates, values, inception_date)
    v_end = _value_on_or_before(dates, values, report_date)
    if v_start and v_end and v_start > 0:
        return round((v_end / v_start - 1) * 100, 2)
    return None


def _resolve_inception_date(report) -> datetime | None:
    if report.custodian and report.custodian.deposits:
        return report.custodian.deposits[0].date
    if report.initial_date:
        return report.initial_date
    if report.transactions:
        return min(t.statement_date for t in report.transactions)
    return None


def get_client_performance(report, bse_prices: list[tuple], report_date: datetime) -> dict:
    inception_date = _resolve_inception_date(report)
    client_simple = round(report.total_pl / report.cost_value * 100, 2) if report.cost_value else None
    bse_simple = _bse_simple_return_from_inception(bse_prices, inception_date, report_date)
    return {
        'since_inception': {
            'client_simple_return': client_simple,
            'bse_simple_return': bse_simple,
            'inception_date': inception_date.strftime('%d %b %Y') if inception_date else None,
        },
        'report_date': report_date.strftime('%d %b %Y'),
    }
