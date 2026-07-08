"""Live NAV / benchmark data sourced from the firm's Google Sheets.

Four sheets (Cash, Gold, Equity, Debt) each hold a date-indexed NAV history,
one column per fund. `fund_mapping.csv` links each fund's ISIN (used
throughout the rest of the app) to its column header in the relevant sheet,
so new funds can be wired up by editing that CSV alone.
"""
from __future__ import annotations

import csv
import io
import time
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
MAPPING_FILE = BASE_DIR / "fund_mapping.csv"

GSHEET_IDS = {
    "cash":   "135cjqAJK7yBp8oM-hIp8WvLmdLJmlHicb05Y7J5IfMg",
    "gold":   "1_RmaSWHumu-Mg3IGDO41VpZM6SSc51NQzcIMZJDt-Xg",
    "equity": "1pwbi3aQZOfVLvNE-Ap4PMH2vjeZ0xTfwdgcZoLMFRBU",
    "debt":   "12RQocyRpcgtgGx_TG2K_25K9Y-h2aGvhswDv_AQx_HI",
}
BENCHMARK_SOURCE = "equity"  # the equity sheet also carries the BSE 500 series
BENCHMARK_COLUMN = "Benchmark"

_CACHE_TTL_SECONDS = 600
_cache: dict[str, tuple[float, list[list[str]]]] = {}
_mapping_cache: tuple[float, list[dict]] | None = None
_MAPPING_TTL_SECONDS = 30


# ---------------------------------------------------------------------------
# Low-level fetch / parse
# ---------------------------------------------------------------------------

def _download_csv_rows(sheet_id: str) -> list[list[str]]:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    req = urllib.request.Request(url, headers={"User-Agent": "BirdsEyeView/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        raise ValueError(f"Google Sheets returned HTTP {e.code} for sheet {sheet_id}") from e
    except urllib.error.URLError as e:
        raise ValueError(f"Could not reach Google Sheets: {e.reason}") from e

    if raw[:200].lower().lstrip().startswith((b"<!doctype", b"<html")):
        raise ValueError(
            f"Sheet {sheet_id} returned an HTML page instead of CSV — "
            "make sure it's shared with 'Anyone with the link'."
        )
    text = raw.decode("utf-8-sig")
    return list(csv.reader(io.StringIO(text)))


def _get_rows(source: str, force: bool = False) -> list[list[str]]:
    now = time.time()
    cached = _cache.get(source)
    if not force and cached and (now - cached[0]) < _CACHE_TTL_SECONDS:
        return cached[1]
    rows = _download_csv_rows(GSHEET_IDS[source])
    _cache[source] = (now, rows)
    return rows


def _parse_date(text: str) -> datetime | None:
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _parse_float(text: str) -> float | None:
    text = (text or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _parse_series(rows: list[list[str]]) -> dict[str, dict]:
    """{column_header: {'dates': [...], 'values': [...]}}, sorted by date."""
    if not rows:
        return {}
    headers = rows[0]
    series: dict[str, dict] = {
        h.strip(): {"dates": [], "values": []} for h in headers[1:] if h and h.strip()
    }
    for row in rows[1:]:
        if not row:
            continue
        date_val = _parse_date(row[0])
        if not date_val:
            continue
        for col_idx, header in enumerate(headers[1:], start=1):
            header = (header or "").strip()
            if not header or header not in series or col_idx >= len(row):
                continue
            value = _parse_float(row[col_idx])
            if value is not None:
                series[header]["dates"].append(date_val)
                series[header]["values"].append(value)
    for data in series.values():
        if data["dates"]:
            paired = sorted(zip(data["dates"], data["values"]), key=lambda p: p[0])
            data["dates"], data["values"] = [p[0] for p in paired], [p[1] for p in paired]
    return series


# ---------------------------------------------------------------------------
# fund_mapping.csv
# ---------------------------------------------------------------------------

def load_fund_mapping(force: bool = False) -> list[dict]:
    global _mapping_cache
    now = time.time()
    if not force and _mapping_cache and (now - _mapping_cache[0]) < _MAPPING_TTL_SECONDS:
        return _mapping_cache[1]

    rows: list[dict] = []
    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                isin = (row.get("isin") or "").strip()
                if not isin or isin.startswith("#"):
                    continue
                rows.append({
                    "isin": isin,
                    "sheet_fund_name": (row.get("sheet_fund_name") or "").strip(),
                    "source": (row.get("source") or "").strip().lower(),
                    "category": (row.get("category") or "").strip(),
                })
    _mapping_cache = (now, rows)
    return rows


def get_category_overrides() -> dict[str, str]:
    """{isin: category}, sourced from fund_mapping.csv (no network needed)."""
    return {r["isin"]: r["category"] for r in load_fund_mapping() if r["category"]}


# ---------------------------------------------------------------------------
# Live current NAVs (per ISIN) — used by the client factsheet engine
# ---------------------------------------------------------------------------

def get_live_current_navs() -> tuple[dict[str, float], list[str]]:
    """Fetch the mapped Google Sheets and resolve {isin: latest_nav}.

    Never raises — network/parse failures are reported as warnings and the
    affected funds are simply left out (caller falls back to other sources).
    """
    mapping = load_fund_mapping()
    by_source: dict[str, list[dict]] = defaultdict(list)
    for r in mapping:
        if r["sheet_fund_name"] and r["source"]:
            by_source[r["source"]].append(r)

    navs: dict[str, float] = {}
    warnings: list[str] = []
    for source, rows in by_source.items():
        if source not in GSHEET_IDS:
            warnings.append(f"Unknown sheet source '{source}' in fund_mapping.csv")
            continue
        try:
            series = _parse_series(_get_rows(source))
        except Exception as e:
            warnings.append(f"Could not fetch the '{source}' Google Sheet: {e}")
            continue
        for r in rows:
            data = series.get(r["sheet_fund_name"])
            if not data or not data["values"]:
                warnings.append(
                    f"'{r['sheet_fund_name']}' not found in the {source} Google Sheet "
                    f"(ISIN {r['isin']})"
                )
                continue
            navs[r["isin"]] = data["values"][-1]
    return navs, warnings


def get_live_bse_prices() -> list[tuple[datetime, float]]:
    """BSE 500 benchmark series from the equity sheet's 'Benchmark' column."""
    series = _parse_series(_get_rows(BENCHMARK_SOURCE))
    data = series.get(BENCHMARK_COLUMN)
    if not data or not data["values"]:
        raise ValueError(
            f"'{BENCHMARK_COLUMN}' column not found in the {BENCHMARK_SOURCE} Google Sheet"
        )
    return list(zip(data["dates"], data["values"]))


# ---------------------------------------------------------------------------
# Live fund-name-keyed series — used by the Master Dashboard
# ---------------------------------------------------------------------------

def get_live_nav_index() -> dict[str, dict]:
    """{sheet_fund_name: {'dates': [...], 'values': [...]}} for every mapped fund."""
    mapping = load_fund_mapping()
    sources_needed = {r["source"] for r in mapping if r["sheet_fund_name"] and r["source"]}
    parsed_by_source: dict[str, dict] = {}
    for source in sources_needed:
        if source not in GSHEET_IDS:
            continue
        parsed_by_source[source] = _parse_series(_get_rows(source))

    result: dict[str, dict] = {}
    for r in mapping:
        fund, source = r["sheet_fund_name"], r["source"]
        if not fund or not source:
            continue
        data = parsed_by_source.get(source, {}).get(fund)
        if data:
            result[fund] = data
    return result


def get_live_scheme_categories() -> dict[str, str]:
    """{sheet_fund_name: category} for Master Dashboard fund/category grouping."""
    return {
        r["sheet_fund_name"]: r["category"]
        for r in load_fund_mapping() if r["sheet_fund_name"] and r["category"]
    }
