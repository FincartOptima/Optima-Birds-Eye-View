"""Client portfolio performance vs the BSE 500 across short lookback windows.

Point-to-point NAV-based returns (no cash-flow adjustment), value-weighted by
the client's actual current holdings + uninvested cash. "Since inception"
reuses the report's XIRR, which is already cash-flow aware.

NAV history is sourced live from Google Sheets (see gsheet_data.py), so every
period return here re-derives itself from whatever data is in the sheets at
request time — no caching beyond gsheet_data's own short TTL.
"""
from __future__ import annotations

from bisect import bisect_right
from datetime import datetime, timedelta

import gsheet_data

PERIOD_DAYS: list[tuple[str, int]] = [
    ('1D',   1),
    ('7D',   7),
    ('30D',  30),
    ('120D', 120),
    ('365D', 365),
]
PERIOD_LABELS = [p[0] for p in PERIOD_DAYS]

# A period return is only reported if we can price at least this fraction of
# the client's current portfolio value at both endpoints of the window.
_MIN_COVERAGE = 0.5


def _value_on_or_before(dates: list[datetime], values: list[float], target: datetime) -> float | None:
    pos = bisect_right(dates, target) - 1
    return values[pos] if pos >= 0 else None


def _bse_returns(bse_prices: list[tuple], report_date: datetime) -> dict[str, float | None]:
    dates = [p[0] for p in bse_prices]
    values = [p[1] for p in bse_prices]
    result: dict[str, float | None] = {}
    for label, days in PERIOD_DAYS:
        from_date = report_date - timedelta(days=days)
        v_start = _value_on_or_before(dates, values, from_date)
        v_end = _value_on_or_before(dates, values, report_date)
        result[label] = round((v_end / v_start - 1) * 100, 2) if v_start and v_end and v_start > 0 else None
    return result


def _client_returns(report, nav_history_by_isin: dict, report_date: datetime) -> dict[str, float | None]:
    result: dict[str, float | None] = {}
    for label, days in PERIOD_DAYS:
        from_date = report_date - timedelta(days=days)
        start_total = report.only_cash
        end_total = report.only_cash
        covered = report.only_cash
        for h in report.holdings:
            hist = nav_history_by_isin.get(h.isin)
            if not hist or not hist['dates']:
                continue
            nav_start = _value_on_or_before(hist['dates'], hist['values'], from_date)
            nav_end = _value_on_or_before(hist['dates'], hist['values'], report_date)
            if nav_start is None or nav_end is None:
                continue
            start_total += h.units * nav_start
            end_total += h.units * nav_end
            covered += h.current_value
        enough_coverage = report.current_value > 0 and (covered / report.current_value) >= _MIN_COVERAGE
        result[label] = round((end_total / start_total - 1) * 100, 2) if enough_coverage and start_total > 0 else None
    return result


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


def get_client_performance(report, bse_prices: list[tuple], report_date: datetime) -> dict:
    mapping = gsheet_data.load_fund_mapping()
    isin_to_fund = {r['isin']: r['sheet_fund_name'] for r in mapping if r['sheet_fund_name']}
    nav_index = gsheet_data.get_live_nav_index()
    nav_history_by_isin = {
        isin: nav_index[fund_name]
        for isin, fund_name in isin_to_fund.items()
        if fund_name in nav_index
    }

    client_ret = _client_returns(report, nav_history_by_isin, report_date)
    bse_ret = _bse_returns(bse_prices, report_date)

    if report.custodian and report.custodian.deposits:
        inception_date = report.custodian.deposits[0].date
    elif report.initial_date:
        inception_date = report.initial_date
    elif report.transactions:
        inception_date = min(t.statement_date for t in report.transactions)
    else:
        inception_date = None

    to_pct = lambda v: round(v * 100, 2) if v is not None else None
    return {
        'period_labels': PERIOD_LABELS,
        'client_returns': client_ret,
        'bse_returns': bse_ret,
        'since_inception': {
            'client_simple_return': to_pct(report.simple_return),
            'bse_simple_return': _bse_simple_return_from_inception(bse_prices, inception_date, report_date),
            'inception_date': inception_date.strftime('%d %b %Y') if inception_date else None,
        },
        'report_date': report_date.strftime('%d %b %Y'),
    }
