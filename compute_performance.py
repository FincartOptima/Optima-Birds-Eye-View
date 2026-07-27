"""Client portfolio performance vs BSE 500 — SEBI-oriented simple returns.

Client simple return = (Current Value − Cost Basis) / Cost Basis. Cost Basis
is the custodian's actual cost of currently-held units + uninvested cash.

BSE 500 simple return uses the most accurate cash-flow schedule available:
  1. Custodian bank deposits (best; requires custodian statement upload)
  2. Trade Master subscription/redemption transactions (good)
  3. Cost Basis as a single lump sum at account opening date (fallback,
     lossy — assumes all money was in on day 1)

For (1) and (2), BSE units are bought/sold on each cashflow date and the
current BSE value is compared to net deposits — the fair apples-to-apples
comparison for a portfolio with staggered contributions.

NAV/benchmark history is sourced live from Google Sheets, so both figures
self-update as those sheets get new data each day.
"""
from __future__ import annotations

from bisect import bisect_right
from datetime import datetime


def _value_on_or_before(dates: list[datetime], values: list[float], target: datetime) -> float | None:
    pos = bisect_right(dates, target) - 1
    return values[pos] if pos >= 0 else None


def _resolve_inception_date(report) -> datetime | None:
    if report.custodian and report.custodian.deposits:
        return report.custodian.deposits[0].date
    if report.initial_date:
        return report.initial_date
    if report.transactions:
        return min(t.statement_date for t in report.transactions)
    return None


def _collect_bse_cashflows(report) -> tuple[list[tuple[datetime, float]], str]:
    """Return ([(date, amount) …], methodology_label). Amount > 0 for deposits,
    < 0 for withdrawals. Prefers custodian bank deposits, then trade-master
    transactions, then a single lump-sum of Cost Basis at inception."""
    if report.custodian and report.custodian.deposits:
        return (
            [(d.date, d.amount) for d in report.custodian.deposits],
            'custodian_deposits',
        )
    if report.transactions:
        cfs: list[tuple[datetime, float]] = []
        for t in report.transactions:
            sign = -1 if t.transaction_type == 'Redemption' else 1
            cfs.append((t.statement_date, sign * t.amount))
        return cfs, 'trade_master_transactions'
    inception = _resolve_inception_date(report)
    if inception and report.cost_value:
        return [(inception, report.cost_value)], 'lump_sum_at_inception'
    return [], 'unavailable'


def _bse_return_cashflow_matched(bse_prices: list[tuple], cashflows: list[tuple[datetime, float]], report_date: datetime) -> float | None:
    """Simulate the same deposit / withdrawal schedule into BSE 500 and return the
    resulting simple return (%). Returns None if the schedule can't be priced."""
    if not bse_prices or not cashflows:
        return None
    dates = [p[0] for p in bse_prices]
    values = [p[1] for p in bse_prices]
    units = 0.0
    net_deposits = 0.0
    for date, amount in cashflows:
        price = _value_on_or_before(dates, values, date)
        if not price or price <= 0:
            continue
        units += amount / price
        net_deposits += amount
    if net_deposits <= 0:
        return None
    final_price = _value_on_or_before(dates, values, report_date)
    if not final_price:
        return None
    final_value = units * final_price
    return round((final_value - net_deposits) / net_deposits * 100, 2)


_METHODOLOGY_LABELS = {
    'custodian_deposits': 'Cash-flow matched (deposits from the account statement)',
    'trade_master_transactions': 'Cash-flow matched (transactions from the tradebook)',
    'lump_sum_at_inception': 'Lump sum at account opening (no tradebook or account statement uploaded)',
    'unavailable': 'Not available',
}


def get_client_performance(report, bse_prices: list[tuple], report_date: datetime) -> dict:
    inception_date = _resolve_inception_date(report)
    client_simple = round(report.total_pl / report.cost_value * 100, 2) if report.cost_value else None
    cashflows, method = _collect_bse_cashflows(report)
    bse_simple = _bse_return_cashflow_matched(bse_prices, cashflows, report_date)
    return {
        'since_inception': {
            'client_simple_return': client_simple,
            'bse_simple_return': bse_simple,
            'bse_methodology': _METHODOLOGY_LABELS.get(method, method),
            'inception_date': inception_date.strftime('%d %b %Y') if inception_date else None,
            'cost_basis': report.cost_value,
            'current_value': report.current_value,
        },
        'report_date': report_date.strftime('%d %b %Y'),
    }
