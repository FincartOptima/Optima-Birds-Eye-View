"""Live market indicators used by the Allocation Rules tab's compliance
panel — currently just Nifty 50 P/E (live, scraped from Trendlyne). Real
inflation/CPI has no scrapable live source (investing.com blocks scripted
requests with a 403, even with realistic browser headers — a hard block,
not a missing-header issue), so it stays a manually-updated value, refreshed
whenever the monthly CPI print comes out.

Nifty P/E regime thresholds match the Allocation Rules tab exactly:
  < 20            -> Aggressive
  20 <= p/e <= 50 -> Moderate
  > 50            -> Conservative
"""
from __future__ import annotations

import re
import time
import urllib.error
import urllib.request

TRENDLYNE_URL = "https://trendlyne.com/equity/PE/NIFTY/1887/nifty-50-price-to-earning-ratios/"
_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Manually-updated inflation figure — see module docstring for why this
# isn't live. Update this pair whenever a new CPI print is available.
CURRENT_INFLATION_PCT = 4.38
CURRENT_INFLATION_AS_OF = "Jul 2026 CPI print (released 12 Aug 2026)"

_CACHE_TTL_SECONDS = 300  # keep requests to Trendlyne polite
_cache: dict[str, tuple[float, float]] = {}  # 'pe' -> (fetched_at, value)

_PE_PATTERNS = [
    re.compile(r"Current PE is (\d+(?:\.\d+)?)"),
    re.compile(r'class="stock-ratio-value[^"]*">(\d+(?:\.\d+)?)<'),
]


def _fetch_nifty_pe_live() -> float:
    req = urllib.request.Request(TRENDLYNE_URL, headers={"User-Agent": _USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raise ValueError(f"Trendlyne returned HTTP {e.code}") from e
    except urllib.error.URLError as e:
        raise ValueError(f"Could not reach Trendlyne: {e.reason}") from e

    for pattern in _PE_PATTERNS:
        m = pattern.search(html)
        if m:
            return float(m.group(1))
    raise ValueError("Could not find the Nifty P/E value on the Trendlyne page (layout may have changed)")


def get_live_nifty_pe(force: bool = False) -> float:
    """Live Nifty 50 P/E, cached for a few minutes to avoid hammering
    Trendlyne on rapid repeat loads. Raises ValueError on failure — callers
    should catch this and degrade gracefully rather than showing a stale
    number as if it were live."""
    now = time.time()
    cached = _cache.get("pe")
    if not force and cached and (now - cached[0]) < _CACHE_TTL_SECONDS:
        return cached[1]
    value = _fetch_nifty_pe_live()
    _cache["pe"] = (now, value)
    return value


def pe_regime(pe: float) -> str:
    if pe < 20:
        return "Aggressive"
    if pe <= 50:
        return "Moderate"
    return "Conservative"


def get_market_status(force: bool = False) -> dict:
    """Everything the Allocation Rules tab's live-status panel needs."""
    try:
        pe = get_live_nifty_pe(force=force)
        pe_error = None
    except Exception as e:
        pe = None
        pe_error = str(e)

    return {
        "nifty_pe": pe,
        "nifty_pe_regime": pe_regime(pe) if pe is not None else None,
        "nifty_pe_error": pe_error,
        "nifty_pe_source": "Trendlyne",
        "inflation_pct": CURRENT_INFLATION_PCT,
        "inflation_as_of": CURRENT_INFLATION_AS_OF,
        "inflation_source": "Manually updated — investing.com blocks automated fetches (403)",
    }
