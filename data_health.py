"""Data freshness / silent-failure visibility. Surfaces exactly when the
data actually is — snapshot holding date, last Drive fetch — and flags any
live NAV/benchmark series that's gone stale, so a feed quietly breaking
(like the Gold sheet's "Price" column, which stopped updating in Feb 2026)
shows up as a warning instead of just silently sitting there.
"""
from __future__ import annotations

from datetime import datetime

import gsheet_data

# A gap bigger than this (calendar days) between "today" and a series'
# latest data point is flagged stale. Generous enough to absorb a normal
# weekend + one holiday without false-flagging every Monday morning.
_STALE_THRESHOLD_DAYS = 5


def _latest_date_in_series(dates: list[datetime]) -> datetime | None:
    return dates[-1] if dates else None


def check_data_health(reports_cache: dict) -> dict:
    now = datetime.now()
    issues: list[dict] = []

    upload_time = reports_cache.get('upload_time')
    snapshot_holding_date = reports_cache.get('snapshot_holding_date')

    # ---- Mapped fund NAV staleness --------------------------------------
    mapping = gsheet_data.load_fund_mapping()
    nav_index = gsheet_data.get_live_nav_index()
    stale_funds: list[dict] = []
    checked_funds = 0
    for m in mapping:
        fund_name = m.get('sheet_fund_name')
        if not fund_name:
            continue
        series = nav_index.get(fund_name)
        if not series or not series.get('dates'):
            issues.append({
                'severity': 'warning',
                'source': fund_name,
                'message': f'"{fund_name}" (ISIN {m["isin"]}) has no live NAV data at all — falls back to a stale/estimated NAV elsewhere in the app.',
            })
            continue
        checked_funds += 1
        latest = _latest_date_in_series(series['dates'])
        gap_days = (now - latest).days
        if gap_days > _STALE_THRESHOLD_DAYS:
            stale_funds.append({'fund': fund_name, 'isin': m['isin'], 'latest_date': latest.strftime('%d %b %Y'), 'gap_days': gap_days})

    for sf in stale_funds:
        issues.append({
            'severity': 'error' if sf['gap_days'] > 30 else 'warning',
            'source': sf['fund'],
            'message': f'"{sf["fund"]}" (ISIN {sf["isin"]}) NAV last updated {sf["latest_date"]} — {sf["gap_days"]} days ago.',
        })

    # ---- BSE 500 benchmark staleness -------------------------------------
    try:
        bse_prices = gsheet_data.get_live_bse_prices()
        bse_latest = _latest_date_in_series([p[0] for p in bse_prices])
        bse_gap_days = (now - bse_latest).days if bse_latest else None
        if bse_gap_days and bse_gap_days > _STALE_THRESHOLD_DAYS:
            issues.append({
                'severity': 'error',
                'source': 'BSE 500 benchmark',
                'message': f'BSE 500 benchmark last updated {bse_latest.strftime("%d %b %Y")} — {bse_gap_days} days ago. Every return figure comparing to this benchmark is affected.',
            })
    except Exception as e:
        bse_latest, bse_gap_days = None, None
        issues.append({'severity': 'error', 'source': 'BSE 500 benchmark', 'message': f'Could not fetch the BSE 500 benchmark: {e}'})

    # ---- Gold/XAU spot price staleness (used by the tactical tilt rule) --
    gold_price_status = None
    try:
        gold_rows = gsheet_data._get_rows('gold')
        gold_series = gsheet_data._parse_series(gold_rows).get('Price', {'dates': []})
        gold_latest = _latest_date_in_series(gold_series['dates'])
        if gold_latest:
            gold_gap_days = (now - gold_latest).days
            gold_price_status = {'latest_date': gold_latest.strftime('%d %b %Y'), 'gap_days': gold_gap_days}
            if gold_gap_days > _STALE_THRESHOLD_DAYS:
                issues.append({
                    'severity': 'warning',
                    'source': 'Gold/XAU price feed',
                    'message': f'Gold spot price ("Price" column, gold sheet) last updated {gold_latest.strftime("%d %b %Y")} — {gold_gap_days} days ago. The BSE500/XAU tactical tilt rule cannot be computed live from this feed.',
                })
    except Exception:
        pass  # informational only — not a mapped/critical feed

    overall_severity = 'error' if any(i['severity'] == 'error' for i in issues) else ('warning' if issues else 'ok')

    return {
        'checked_at': now.strftime('%d %b %Y, %H:%M'),
        'upload_time': upload_time,
        'snapshot_holding_date': snapshot_holding_date,
        'funds_checked': checked_funds,
        'funds_stale': len(stale_funds),
        'bse_500_latest_date': bse_latest.strftime('%d %b %Y') if bse_latest else None,
        'gold_price_status': gold_price_status,
        'overall_severity': overall_severity,
        'issues': issues,
    }
