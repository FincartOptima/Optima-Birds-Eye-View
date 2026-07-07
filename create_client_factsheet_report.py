from __future__ import annotations

import csv
import io
import math
import re
import sys

# ReportLab + matplotlib font rendering on Windows can exceed the default
# 1000-frame recursion limit. Set it high at import time so it applies to
# every worker/thread.
sys.setrecursionlimit(50000)
from bisect import bisect_right
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas as PdfCanvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader


BASE_DIR = Path(__file__).resolve().parent
TRANSACTION_FILE = BASE_DIR / "Trade Allocation_Master.xlsx"
BSE_500_FILE = BASE_DIR / "BSE_DLY_BSE500, 1D (8).csv"
OPTIONAL_CURRENT_NAV_FILE = BASE_DIR / "Current_NAVs.xlsx"
OUTPUT_FILE = BASE_DIR / "Client_Factsheet_Report.xlsx"
PDF_OUTPUT_DIR = BASE_DIR / "Client_Factsheets"

REPORT_DATE = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

CATEGORY_BY_ISIN = {
    "INF109KC1RH9": "Indian Equity",
    "INF200K01RJ1": "Indian Equity",
    "INF109K01X40": "Indian Equity",
    "INF247L01CW5": "Indian Equity",
    "INF740K01QA7": "Indian Equity",
    "INF00XX01747": "Indian Equity",
    "INF194KB1AL4": "Indian Equity",
    "INF205K013T3": "Indian Equity",
    "INF846K01X06": "Foreign Equity",
    "INF846K01Y39": "Foreign Equity",
    "INF090I01JR0": "Foreign Equity",
    "INF090I01IZ5": "Foreign Equity",
    "INF174K01MP6": "Gold",
    "INF109K01U92": "Gold",
    "INF846K01DI3": "Debt",
    "INF277K017Q3": "Cash Fund",
    "INF174K01NE8": "Cash Fund",
    "INF754K01LB7": "Foreign Equity",
}

CATEGORY_ORDER = ["Indian Equity", "Foreign Equity", "Gold", "Debt", "Cash Fund", "Only Cash"]


@dataclass
class Transaction:
    source_sheet: str
    source_row: int
    statement_date: datetime
    transaction_type: str
    ucc: str
    client_name: str
    scheme_name: str
    folio_no: str
    isin: str
    amount: float
    units: float | None
    value_date: datetime | str | None
    allocation_date: datetime | str | None


@dataclass
class Holding:
    isin: str
    scheme_name: str
    category: str
    units: float = 0.0
    cost_value: float = 0.0
    realized_pl: float = 0.0
    current_nav: float | None = None
    current_value: float = 0.0
    unrealized_pl: float = 0.0
    total_pl: float = 0.0
    nav_source: str = ""


@dataclass
class ClientReport:
    client_name: str
    ucc: str
    initial_investment: float
    initial_date: datetime | None
    transactions: list[Transaction] = field(default_factory=list)
    holdings: list[Holding] = field(default_factory=list)
    only_cash: float = 0.0
    cost_value: float = 0.0
    current_value: float = 0.0
    unrealized_pl: float = 0.0
    realized_pl: float = 0.0
    total_pl: float = 0.0
    xirr: float | None = None
    benchmark_current_value: float | None = None
    benchmark_xirr: float | None = None
    category_rows: list[dict[str, Any]] = field(default_factory=list)
    top_holdings: list[Holding] = field(default_factory=list)
    performance_rows: list[tuple[datetime, float | None, float | None]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Text / data helpers
# ---------------------------------------------------------------------------

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def clean_header(value: Any) -> str:
    return clean_text(value).lower()


def to_number(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", "").strip())
        except ValueError:
            return None
    return None


def as_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                pass
    return None


def parse_sheet_date(sheet_name: str) -> datetime | None:
    if not re.fullmatch(r"\d{8}", sheet_name):
        return None
    try:
        return datetime.strptime(sheet_name, "%d%m%Y")
    except ValueError:
        return None


def header_index(headers: list[Any]) -> dict[str, int]:
    index: dict[str, int] = {}
    for pos, header in enumerate(headers):
        key = clean_header(header)
        if key and key not in index:
            index[key] = pos
    return index


def row_value(row: list[Any], headers: list[Any], *aliases: str) -> Any:
    index = header_index(headers)
    for alias in aliases:
        pos = index.get(clean_header(alias))
        if pos is not None and pos < len(row):
            return row[pos]
    return None


def safe_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned) or "Client"
    base = cleaned[:31]
    candidate = base
    suffix = 2
    while candidate in used_names:
        suffix_text = f" {suffix}"
        candidate = f"{base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def infer_category(isin: str, scheme_name: str) -> str:
    if isin in CATEGORY_BY_ISIN:
        return CATEGORY_BY_ISIN[isin]
    text = scheme_name.lower()
    if any(word in text for word in ["u.s.", " us ", "asian", "china", "global", "technology equity"]):
        return "Foreign Equity"
    if "gold" in text:
        return "Gold"
    if any(word in text for word in ["bond", "debt", "gilt"]):
        return "Debt"
    if any(word in text for word in ["liquid", "arbitrage", "money market", "overnight"]):
        return "Cash Fund"
    return "Indian Equity"


# ---------------------------------------------------------------------------
# Data readers
# ---------------------------------------------------------------------------

def read_master(workbook) -> dict[str, dict[str, Any]]:
    if "Master" not in workbook.sheetnames:
        return {}
    sheet = workbook["Master"]
    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    index = header_index(headers)
    master: dict[str, dict[str, Any]] = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_values = list(row)
        ucc_pos = index.get("ucc code")
        client_pos = index.get("client name")
        initial_pos = index.get("initial investment")
        if ucc_pos is None or client_pos is None:
            continue
        ucc = clean_text(row_values[ucc_pos] if ucc_pos < len(row_values) else None)
        client_name = clean_text(row_values[client_pos] if client_pos < len(row_values) else None)
        if not ucc or not client_name:
            continue
        master[ucc] = {
            "client_name": client_name,
            "initial_investment": to_number(row_values[initial_pos] if initial_pos is not None else None) or 0.0,
            "initial_date": as_datetime(row_values[index.get("date of opening")] if index.get("date of opening") is not None else None)
            or as_datetime(row_values[index.get("form received date")] if index.get("form received date") is not None else None),
        }
    return master


def read_transactions(workbook, master: dict[str, dict[str, Any]]) -> list[Transaction]:
    transactions: list[Transaction] = []
    for sheet in workbook.worksheets:
        statement_date = parse_sheet_date(sheet.title)
        if statement_date is None:
            continue
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue
        headers = list(rows[0])
        transaction_type = "Subscription"
        for row_number, row_tuple in enumerate(rows[1:], start=2):
            row = list(row_tuple)
            first_cell = clean_header(row[0] if row else None)
            if "redemption" in first_cell and first_cell != "redemption":
                transaction_type = "Redemption"
                continue
            if first_cell == "ucc":
                headers = row
                if any(clean_header(header) == "redemption" for header in headers):
                    transaction_type = "Redemption"
                continue
            amount = to_number(row_value(row, headers, "Amount/Units"))
            ucc = clean_text(row_value(row, headers, "UCC"))
            client_name = clean_text(row_value(row, headers, "Client Name"))
            scheme_name = clean_text(row_value(row, headers, "Scheme Name"))
            isin = clean_text(row_value(row, headers, "ISIN Code"))
            if not (amount is not None and re.fullmatch(r"CRFM\d+", ucc, re.I) and client_name and scheme_name):
                continue
            if ucc in master:
                client_name = master[ucc]["client_name"]
            transactions.append(
                Transaction(
                    source_sheet=sheet.title,
                    source_row=row_number,
                    statement_date=statement_date,
                    transaction_type=transaction_type,
                    ucc=ucc,
                    client_name=client_name,
                    scheme_name=scheme_name,
                    folio_no=clean_text(row_value(row, headers, "Folio No")),
                    isin=isin,
                    amount=amount,
                    units=to_number(row_value(row, headers, "UNITS", "Units", "Value")),
                    value_date=row_value(row, headers, "Value Date"),
                    allocation_date=row_value(row, headers, "Child Allocation", "Child Allocation ", "Redemption"),
                )
            )
    return transactions


def read_bse_prices(file_path: Path) -> list[tuple[datetime, float]]:
    if file_path.suffix.lower() == ".csv":
        return _read_bse_csv(file_path)
    return _read_bse_xlsx(file_path)


def _read_bse_csv(file_path: Path) -> list[tuple[datetime, float]]:
    prices: list[tuple[datetime, float]] = []
    with open(file_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_val = as_datetime(row.get("time", ""))
            close = to_number(row.get("close", ""))
            if date_val and close:
                prices.append((date_val, close))
    prices.sort(key=lambda x: x[0])
    return prices


def _read_bse_xlsx(file_path: Path) -> list[tuple[datetime, float]]:
    workbook = load_workbook(file_path, data_only=True, read_only=True)
    sheet = workbook.worksheets[0]
    prices: list[tuple[datetime, float]] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        date_value = as_datetime(row[0])
        price = to_number(row[1])
        if date_value and price:
            prices.append((date_value, price))
    prices.sort(key=lambda item: item[0])
    return prices


def read_current_navs(file_path: Path) -> tuple[dict[str, float], dict[str, str]]:
    if not file_path.exists():
        return {}, {}
    workbook = load_workbook(file_path, data_only=True, read_only=True)
    sheet = workbook.worksheets[0]
    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    index = header_index(headers)
    navs: dict[str, float] = {}
    categories: dict[str, str] = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_values = list(row)
        isin_pos = index.get("isin")
        nav_pos = index.get("current nav")
        category_pos = index.get("category")
        if isin_pos is None or nav_pos is None:
            continue
        isin = clean_text(row_values[isin_pos] if isin_pos < len(row_values) else None)
        nav = to_number(row_values[nav_pos] if nav_pos < len(row_values) else None)
        if isin and nav is not None:
            navs[isin] = nav
            if category_pos is not None and category_pos < len(row_values):
                category = clean_text(row_values[category_pos])
                if category:
                    categories[isin] = category
    return navs, categories


def read_client_file_csv(file_path: Path) -> tuple[dict[str, float], dict[str, str]]:
    """Read a custodian holdings-snapshot CSV (e.g. 'client file.csv').

    The snapshot has one row per (client, holding) with the live NAV in the
    'Unit Price' column. We extract a {ISIN: current_nav} map so the snapshot
    can act as the source of current valuations (replacing Current_NAVs.xlsx).
    Categories are left to CATEGORY_BY_ISIN / infer_category, so we return an
    empty category-override dict.
    """
    navs: dict[str, float] = {}
    if not file_path.exists():
        return navs, {}
    with open(file_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            isin = clean_text(row.get("ISIN"))
            nav = to_number(row.get("Unit Price"))
            if isin and nav is not None and nav > 0:
                navs[isin] = nav
    return navs, {}


def latest_observed_navs(transactions: list[Transaction]) -> dict[str, tuple[float, datetime]]:
    navs: dict[str, tuple[float, datetime]] = {}
    for transaction in transactions:
        if not transaction.isin or not transaction.units:
            continue
        nav = transaction.amount / transaction.units
        current = navs.get(transaction.isin)
        if current is None or transaction.statement_date >= current[1]:
            navs[transaction.isin] = (nav, transaction.statement_date)
    return navs


# ---------------------------------------------------------------------------
# XIRR / benchmark helpers
# ---------------------------------------------------------------------------

def xnpv(rate: float, cashflows: list[tuple[datetime, float]]) -> float:
    start = min(date for date, _ in cashflows)
    total = 0.0
    for date, amount in cashflows:
        years = (date - start).days / 365.0
        total += amount / ((1.0 + rate) ** years)
    return total


def xirr(cashflows: list[tuple[datetime, float]]) -> float | None:
    valid_cashflows = [(date, amount) for date, amount in cashflows if abs(amount) > 0.000001]
    if not valid_cashflows:
        return None
    if not any(amount > 0 for _, amount in valid_cashflows) or not any(amount < 0 for _, amount in valid_cashflows):
        return None
    low = -0.9999
    high = 10.0
    low_value = xnpv(low, valid_cashflows)
    high_value = xnpv(high, valid_cashflows)
    expansion_count = 0
    while low_value * high_value > 0 and expansion_count < 20:
        high *= 2
        high_value = xnpv(high, valid_cashflows)
        expansion_count += 1
    if low_value * high_value > 0:
        return None
    for _ in range(100):
        mid = (low + high) / 2
        mid_value = xnpv(mid, valid_cashflows)
        if abs(mid_value) < 0.000001:
            return mid
        if low_value * mid_value <= 0:
            high = mid
            high_value = mid_value
        else:
            low = mid
            low_value = mid_value
    return (low + high) / 2


def price_on_or_before(prices: list[tuple[datetime, float]], date: datetime) -> float | None:
    dates = [item[0] for item in prices]
    pos = bisect_right(dates, date) - 1
    if pos < 0:
        return None
    return prices[pos][1]


def benchmark_value_and_xirr(
    transactions: list[Transaction],
    prices: list[tuple[datetime, float]],
    report_date: datetime,
    initial_investment: float = 0.0,
    initial_date: datetime | None = None,
) -> tuple[float | None, float | None]:
    if not prices:
        return None, None
    if initial_investment > 0:
        start_date = initial_date or (min(item.statement_date for item in transactions) if transactions else report_date)
        start_price = price_on_or_before(prices, start_date)
        final_price = price_on_or_before(prices, report_date) or prices[-1][1]
        if not start_price:
            return None, None
        current_value = initial_investment * final_price / start_price
        return current_value, xirr([(start_date, -initial_investment), (report_date, current_value)])
    if not transactions:
        return None, None
    units = 0.0
    cashflows: list[tuple[datetime, float]] = []
    for transaction in sorted(transactions, key=lambda item: (item.statement_date, item.source_sheet, item.source_row)):
        price = price_on_or_before(prices, transaction.statement_date)
        if not price:
            continue
        if transaction.transaction_type == "Redemption":
            units -= transaction.amount / price
            cashflows.append((transaction.statement_date, transaction.amount))
        else:
            units += transaction.amount / price
            cashflows.append((transaction.statement_date, -transaction.amount))
    final_price = price_on_or_before(prices, report_date) or prices[-1][1]
    current_value = units * final_price
    if current_value > 0:
        cashflows.append((report_date, current_value))
    return current_value, xirr(cashflows)


def make_benchmark_series(
    transactions: list[Transaction],
    prices: list[tuple[datetime, float]],
    report_date: datetime,
) -> list[tuple[datetime, float]]:
    if not transactions:
        return []
    sorted_transactions = sorted(transactions, key=lambda item: (item.statement_date, item.source_row))
    start_date = min(item.statement_date for item in sorted_transactions)
    units = 0.0
    tx_index = 0
    series: list[tuple[datetime, float]] = []
    for date, price in prices:
        if date < start_date:
            continue
        if date > report_date:
            break
        while tx_index < len(sorted_transactions) and sorted_transactions[tx_index].statement_date <= date:
            transaction = sorted_transactions[tx_index]
            tx_price = price_on_or_before(prices, transaction.statement_date)
            if tx_price:
                if transaction.transaction_type == "Redemption":
                    units -= transaction.amount / tx_price
                else:
                    units += transaction.amount / tx_price
            tx_index += 1
        series.append((date, units * price))
    return series


def make_client_series(
    transactions: list[Transaction],
    holdings: dict[str, Holding],
    report_date: datetime,
) -> list[tuple[datetime, float]]:
    if not transactions:
        return []
    nav_history: dict[str, list[tuple[datetime, float]]] = defaultdict(list)
    for transaction in transactions:
        if transaction.units:
            nav_history[transaction.isin].append((transaction.statement_date, transaction.amount / transaction.units))
    for isin in nav_history:
        nav_history[isin].sort(key=lambda item: item[0])
        final_nav = holdings.get(isin).current_nav if holdings.get(isin) else None
        if final_nav:
            nav_history[isin].append((report_date, final_nav))
    event_dates = sorted({transaction.statement_date for transaction in transactions} | {report_date})
    holding_units: dict[str, float] = defaultdict(float)
    series: list[tuple[datetime, float]] = []
    sorted_transactions = sorted(transactions, key=lambda item: (item.statement_date, item.source_row))
    tx_index = 0
    for date in event_dates:
        while tx_index < len(sorted_transactions) and sorted_transactions[tx_index].statement_date <= date:
            transaction = sorted_transactions[tx_index]
            if transaction.units:
                if transaction.transaction_type == "Redemption":
                    holding_units[transaction.isin] -= transaction.units
                else:
                    holding_units[transaction.isin] += transaction.units
            tx_index += 1
        value = 0.0
        for isin, units in holding_units.items():
            history = nav_history.get(isin, [])
            dates = [item[0] for item in history]
            pos = bisect_right(dates, date) - 1
            if pos >= 0:
                value += units * history[pos][1]
        series.append((date, value))
    return series


# ---------------------------------------------------------------------------
# Build client reports
# ---------------------------------------------------------------------------

def build_client_reports(
    master: dict[str, dict[str, Any]],
    transactions: list[Transaction],
    bse_prices: list[tuple[datetime, float]],
    current_navs: dict[str, float],
    category_overrides: dict[str, str],
) -> list[ClientReport]:
    transactions_by_ucc: dict[str, list[Transaction]] = defaultdict(list)
    for transaction in transactions:
        transactions_by_ucc[transaction.ucc].append(transaction)
    latest_navs = latest_observed_navs(transactions)
    reports: list[ClientReport] = []
    for ucc, master_row in sorted(master.items(), key=lambda item: item[1]["client_name"].lower()):
        client_transactions = sorted(
            transactions_by_ucc.get(ucc, []),
            key=lambda item: (item.statement_date, item.source_sheet, item.source_row),
        )
        report = ClientReport(
            client_name=master_row["client_name"],
            ucc=ucc,
            initial_investment=master_row["initial_investment"],
            initial_date=master_row.get("initial_date"),
            transactions=client_transactions,
        )
        holdings_by_isin: dict[str, Holding] = {}
        for transaction in client_transactions:
            if transaction.isin not in holdings_by_isin:
                category = category_overrides.get(transaction.isin) or infer_category(transaction.isin, transaction.scheme_name)
                holdings_by_isin[transaction.isin] = Holding(
                    isin=transaction.isin,
                    scheme_name=transaction.scheme_name,
                    category=category,
                )
            holding = holdings_by_isin[transaction.isin]
            units = transaction.units or 0.0
            if not units:
                fallback_nav = current_navs.get(transaction.isin)
                if fallback_nav is None and transaction.isin in latest_navs:
                    fallback_nav = latest_navs[transaction.isin][0]
                if fallback_nav:
                    units = transaction.amount / fallback_nav
                else:
                    units = transaction.amount
            if transaction.transaction_type == "Redemption":
                avg_cost = holding.cost_value / holding.units if holding.units else 0.0
                cost_removed = min(units, holding.units) * avg_cost if units else min(transaction.amount, holding.cost_value)
                holding.units -= units
                holding.cost_value -= cost_removed
                holding.realized_pl += transaction.amount - cost_removed
            else:
                holding.units += units
                holding.cost_value += transaction.amount
        for holding in holdings_by_isin.values():
            if holding.isin in current_navs:
                holding.current_nav = current_navs[holding.isin]
                holding.nav_source = "Current_NAVs.xlsx"
            elif holding.isin in latest_navs:
                holding.current_nav = latest_navs[holding.isin][0]
                holding.nav_source = "Latest transaction NAV"
            elif holding.units and holding.cost_value:
                holding.current_nav = holding.cost_value / holding.units
                holding.nav_source = "Cost fallback"
            holding.current_value = holding.units * holding.current_nav if holding.current_nav else 0.0
            holding.unrealized_pl = holding.current_value - holding.cost_value
            holding.total_pl = holding.unrealized_pl + holding.realized_pl
        report.holdings = sorted(holdings_by_isin.values(), key=lambda item: item.current_value, reverse=True)
        report.realized_pl = sum(holding.realized_pl for holding in report.holdings)
        report.cost_value = sum(holding.cost_value for holding in report.holdings)
        report.current_value = sum(holding.current_value for holding in report.holdings)
        net_fund_investment = sum(
            transaction.amount if transaction.transaction_type == "Subscription" else -transaction.amount
            for transaction in client_transactions
        )
        report.only_cash = max(report.initial_investment - net_fund_investment, 0.0)
        report.cost_value += report.only_cash
        report.current_value += report.only_cash
        report.unrealized_pl = report.current_value - report.cost_value
        report.total_pl = report.unrealized_pl + report.realized_pl
        category_values = defaultdict(lambda: {"cost": 0.0, "current": 0.0, "unrealized": 0.0, "realized": 0.0})
        for holding in report.holdings:
            category_values[holding.category]["cost"] += holding.cost_value
            category_values[holding.category]["current"] += holding.current_value
            category_values[holding.category]["unrealized"] += holding.unrealized_pl
            category_values[holding.category]["realized"] += holding.realized_pl
        category_values["Only Cash"]["cost"] += report.only_cash
        category_values["Only Cash"]["current"] += report.only_cash
        report.category_rows = []
        for category in CATEGORY_ORDER:
            row = category_values[category]
            current = row["current"]
            report.category_rows.append(
                {
                    "Category": category,
                    "Cost Value": row["cost"],
                    "Current Value": current,
                    "Allocation %": current / report.current_value if report.current_value else 0.0,
                    "Unrealized P/L": row["unrealized"],
                    "Realized P/L": row["realized"],
                    "Total P/L": row["unrealized"] + row["realized"],
                }
            )
        report.top_holdings = [holding for holding in report.holdings if holding.current_value > 0][:5]
        portfolio_cashflows = []
        if report.initial_investment > 0:
            start_date = report.initial_date or (
                min(transaction.statement_date for transaction in client_transactions) if client_transactions else REPORT_DATE
            )
            portfolio_cashflows.append((start_date, -report.initial_investment))
        else:
            for transaction in client_transactions:
                amount = transaction.amount if transaction.transaction_type == "Redemption" else -transaction.amount
                portfolio_cashflows.append((transaction.statement_date, amount))
        if report.current_value:
            portfolio_cashflows.append((REPORT_DATE, report.current_value))
        report.xirr = xirr(portfolio_cashflows)
        report.benchmark_current_value, report.benchmark_xirr = benchmark_value_and_xirr(
            client_transactions, bse_prices, REPORT_DATE, report.initial_investment, report.initial_date
        )
        client_series = make_client_series(client_transactions, holdings_by_isin, REPORT_DATE)
        benchmark_series = make_benchmark_series(client_transactions, bse_prices, REPORT_DATE)
        benchmark_by_date = {date: value for date, value in benchmark_series}
        report.performance_rows = [
            (date, value, benchmark_by_date.get(date)) for date, value in client_series
        ]
        reports.append(report)
    return reports


# ---------------------------------------------------------------------------
# Excel output (unchanged)
# ---------------------------------------------------------------------------

def style_range_title(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def set_money(cell) -> None:
    cell.number_format = '#,##0.00'


def set_percent(cell) -> None:
    cell.number_format = '0.00%'


def write_table(sheet, start_row: int, start_col: int, headers: list[str], rows: list[list[Any]]) -> int:
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for offset, header in enumerate(headers):
        cell = sheet.cell(start_row, start_col + offset, header)
        style_range_title(cell)
        cell.border = border
    for row_offset, row_values in enumerate(rows, start=1):
        for col_offset, value in enumerate(row_values):
            cell = sheet.cell(start_row + row_offset, start_col + col_offset, value)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    return start_row + len(rows)


def write_client_sheet(workbook: Workbook, report: ClientReport, used_names: set[str]) -> None:
    sheet = workbook.create_sheet(safe_sheet_name(report.client_name, used_names))
    sheet.sheet_view.showGridLines = False
    for col in range(1, 13):
        sheet.column_dimensions[get_column_letter(col)].width = 14
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 18
    sheet.column_dimensions["D"].width = 18
    sheet.column_dimensions["E"].width = 18
    sheet.column_dimensions["G"].width = 24
    sheet.column_dimensions["H"].width = 18
    sheet.column_dimensions["I"].width = 18
    sheet.column_dimensions["J"].width = 18
    sheet.column_dimensions["K"].width = 18
    sheet.merge_cells("A1:K1")
    sheet["A1"] = "Client Portfolio Factsheet"
    sheet["A1"].font = Font(bold=True, size=18, color="1F4E78")
    sheet["A1"].alignment = Alignment(horizontal="center")
    sheet["A2"] = "Client"
    sheet["B2"] = report.client_name
    sheet["D2"] = "UCC"
    sheet["E2"] = report.ucc
    sheet["G2"] = "Report Date"
    sheet["H2"] = REPORT_DATE
    sheet["H2"].number_format = "dd-mmm-yyyy"
    for cell_ref in ["A2", "D2", "G2"]:
        sheet[cell_ref].font = Font(bold=True)
    snapshot_headers = ["Metric", "Value", "Metric", "Value"]
    snapshot_rows = [
        ["Cost Value", report.cost_value, "Current Value", report.current_value],
        ["Unrealized P/L", report.unrealized_pl, "Realized P/L", report.realized_pl],
        ["Total P/L", report.total_pl, "Portfolio XIRR", report.xirr],
        ["BSE 500 Value", report.benchmark_current_value, "BSE 500 XIRR", report.benchmark_xirr],
    ]
    write_table(sheet, 4, 1, snapshot_headers, snapshot_rows)
    for row in range(5, 9):
        set_money(sheet.cell(row, 2))
        set_money(sheet.cell(row, 4))
    set_percent(sheet["D7"])
    set_percent(sheet["D8"])
    category_headers = ["Category", "Cost Value", "Current Value", "Allocation %", "Unrealized P/L", "Realized P/L", "Total P/L"]
    category_rows = [
        [r["Category"], r["Cost Value"], r["Current Value"], r["Allocation %"], r["Unrealized P/L"], r["Realized P/L"], r["Total P/L"]]
        for r in report.category_rows
    ]
    write_table(sheet, 10, 1, category_headers, category_rows)
    for row in range(11, 11 + len(category_rows)):
        for col in [2, 3, 5, 6, 7]:
            set_money(sheet.cell(row, col))
        set_percent(sheet.cell(row, 4))
    holding_headers = ["Fund", "Category", "ISIN", "Units", "Cost Value", "Current NAV", "Current Value", "Allocation %", "Unrealized P/L", "Realized P/L", "Total P/L"]
    holding_rows = []
    for holding in report.holdings:
        holding_rows.append([
            holding.scheme_name, holding.category, holding.isin, holding.units,
            holding.cost_value, holding.current_nav, holding.current_value,
            holding.current_value / report.current_value if report.current_value else 0.0,
            holding.unrealized_pl, holding.realized_pl, holding.total_pl,
        ])
    holdings_end = write_table(sheet, 19, 1, holding_headers, holding_rows)
    for row in range(20, holdings_end + 1):
        for col in [4, 5, 6, 7, 9, 10, 11]:
            set_money(sheet.cell(row, col))
        set_percent(sheet.cell(row, 8))
    top_start = holdings_end + 3
    sheet.merge_cells(start_row=top_start, start_column=1, end_row=top_start, end_column=6)
    sheet.cell(top_start, 1, "Top 5 Holdings")
    style_range_title(sheet.cell(top_start, 1))
    top_headers = ["Rank", "Fund", "Category", "Current Value", "Allocation %", "Total P/L"]
    top_rows = [
        [rank, holding.scheme_name, holding.category, holding.current_value,
         holding.current_value / report.current_value if report.current_value else 0.0, holding.total_pl]
        for rank, holding in enumerate(report.top_holdings, start=1)
    ]
    write_table(sheet, top_start + 1, 1, top_headers, top_rows)
    for row in range(top_start + 2, top_start + 2 + len(top_rows)):
        set_money(sheet.cell(row, 4))
        set_percent(sheet.cell(row, 5))
        set_money(sheet.cell(row, 6))
    perf_start = top_start + 10
    perf_headers = ["Date", "Client Value", "BSE 500 Benchmark Value"]
    perf_rows = [[date, client_value, benchmark_value] for date, client_value, benchmark_value in report.performance_rows]
    write_table(sheet, perf_start, 1, perf_headers, perf_rows)
    for row in range(perf_start + 1, perf_start + 1 + len(perf_rows)):
        sheet.cell(row, 1).number_format = "dd-mmm-yyyy"
        set_money(sheet.cell(row, 2))
        set_money(sheet.cell(row, 3))
    if len(perf_rows) >= 2:
        chart = LineChart()
        chart.title = "Client vs BSE 500 Performance"
        chart.y_axis.title = "Value"
        chart.x_axis.title = "Date"
        data = Reference(sheet, min_col=2, max_col=3, min_row=perf_start, max_row=perf_start + len(perf_rows))
        categories = Reference(sheet, min_col=1, min_row=perf_start + 1, max_row=perf_start + len(perf_rows))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.height = 9
        chart.width = 22
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = False
        sheet.add_chart(chart, f"G{perf_start}")


def write_summary_sheet(workbook: Workbook, reports: list[ClientReport]) -> None:
    sheet = workbook.active
    sheet.title = "Summary"
    sheet.sheet_view.showGridLines = False
    sheet["A1"] = "Client Factsheet Summary"
    sheet["A1"].font = Font(bold=True, size=16, color="1F4E78")
    headers = ["Client", "UCC", "Transactions", "Cost Value", "Current Value", "Unrealized P/L", "Realized P/L", "Total P/L", "Portfolio XIRR", "BSE 500 XIRR"]
    rows = [
        [report.client_name, report.ucc, len(report.transactions), report.cost_value, report.current_value,
         report.unrealized_pl, report.realized_pl, report.total_pl, report.xirr, report.benchmark_xirr]
        for report in reports
    ]
    write_table(sheet, 3, 1, headers, rows)
    for row in range(4, 4 + len(rows)):
        for col in [4, 5, 6, 7, 8]:
            set_money(sheet.cell(row, col))
        set_percent(sheet.cell(row, 9))
        set_percent(sheet.cell(row, 10))
    widths = [42, 12, 14, 16, 16, 16, 16, 16, 14, 14]
    for col_num, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(col_num)].width = width


def write_notes_sheet(workbook: Workbook, used_names: set[str], current_navs_present: bool) -> None:
    sheet = workbook.create_sheet(safe_sheet_name("Notes", used_names))
    sheet.sheet_view.showGridLines = False
    notes = [
        "This workbook is generated from Trade Allocation_Master.xlsx and BSE_DLY_BSE500, 1D (8).csv.",
        "Fund allocation is based on current value.",
        "Cost value is the remaining cost basis after redemptions. Redemption P/L uses average cost.",
        "Only Cash is estimated as Master Initial Investment minus net fund investment, floored at zero.",
        "Portfolio XIRR uses transaction cash flows and final current value on the report date.",
        "BSE 500 XIRR uses the same subscription/redemption cash-flow dates invested in BSE 500.",
        "Client performance line uses available transaction NAV observations and current NAV.",
    ]
    if current_navs_present:
        notes.append("Current NAV was read from Current_NAVs.xlsx where ISINs matched.")
    else:
        notes.append("Current_NAVs.xlsx was not found, so current NAV is estimated from latest observed transaction NAV.")
    for row, note in enumerate(notes, start=1):
        sheet.cell(row, 1, note)
        sheet.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    sheet.column_dimensions["A"].width = 110


# =====================================================================
# PDF FACTSHEET GENERATION
# =====================================================================

_PW, _PH = A4  # 595.28, 841.89
_LM = 28
_RM = 28
_CW = _PW - _LM - _RM

_NAVY = HexColor("#14365C")
_DARK_NAVY = HexColor("#0D2440")
_GOLD = HexColor("#C5922E")
_CREAM = HexColor("#F3E9CE")
_WHITE = HexColor("#FFFFFF")
_LIGHT_BG = HexColor("#F4F6F9")
_LIGHT_BLUE = HexColor("#E8EEF6")
_BORDER = HexColor("#C2D0E0")
_TEXT_DARK = HexColor("#2C3E50")
_TEXT_MED = HexColor("#5A6C7E")
_TEXT_LIGHT = HexColor("#8899AA")
_GREEN_VAL = HexColor("#1B7A2F")
_RED_VAL = HexColor("#C0392B")

_PIE_COLORS = ["#1F4E78", "#3A7CA5", "#C5922E", "#5DADE2", "#A0B4C8"]
_BAR_COLOR = "#1F4E78"

_PDF_CAT_DISPLAY = {
    "Indian Equity": "Domestic Equity MF",
    "Foreign Equity": "International Equity",
    "Gold": "Commodities (Gold + Nat Res)",
    "Debt": "Debt Mutual Funds",
    "Cash Fund": "Cash & Equivalent",
    "Only Cash": "Cash & Equivalent",
}


def _fmt_pct(val: float | None, sign: bool = True) -> str:
    if val is None:
        return "N/A"
    pct = val * 100
    if sign and pct >= 0:
        return f"+{pct:.2f}%"
    return f"{pct:.2f}%"


def _pct_color(val: float | None):
    if val is None or val == 0:
        return _TEXT_DARK
    return _GREEN_VAL if val > 0 else _RED_VAL


def _inception(report: ClientReport) -> datetime:
    if report.initial_date:
        return report.initial_date
    if report.transactions:
        return min(t.statement_date for t in report.transactions)
    return REPORT_DATE


def _fmt_inr(val: float | None) -> str:
    if val is None or val == 0:
        return "Rs 0"
    if abs(val) >= 1e7:
        return f"Rs {val / 1e7:.2f} Cr"
    if abs(val) >= 1e5:
        return f"Rs {val / 1e5:.2f} L"
    return f"Rs {val:,.0f}"


def _aggregate_pdf_categories(report: ClientReport) -> list[tuple[str, float]]:
    agg: dict[str, float] = {}
    for row in report.category_rows:
        display_name = _PDF_CAT_DISPLAY.get(row["Category"], row["Category"])
        agg[display_name] = agg.get(display_name, 0.0) + row["Allocation %"]
    result = [(name, weight) for name, weight in agg.items() if weight > 0.001]
    result.sort(key=lambda x: x[1], reverse=True)
    return result


def _scheme_count(report: ClientReport) -> int:
    return len([h for h in report.holdings if h.current_value > 0])


def _matplotlib_pyplot():
    """Import matplotlib lazily (only the PDF chart helpers need it)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    return plt


def _make_pie_image(labels: list[str], sizes: list[float]) -> ImageReader:
    plt = _matplotlib_pyplot()
    fig, ax = plt.subplots(figsize=(2.6, 2.6))
    colors = _PIE_COLORS[: len(labels)]
    while len(colors) < len(labels):
        colors.append("#A0B4C8")
    wedges, _ = ax.pie(
        sizes,
        colors=colors,
        startangle=90,
        counterclock=False,
        wedgeprops={"width": 0.62, "edgecolor": "white", "linewidth": 1.5},
    )
    for wedge, size in zip(wedges, sizes):
        ang = (wedge.theta2 + wedge.theta1) / 2
        r = 0.72
        x = r * math.cos(math.radians(ang))
        y_pos = r * math.sin(math.radians(ang))
        ax.text(x, y_pos, f"{size * 100:.1f}%", ha="center", va="center", fontsize=6.5, fontweight="bold", color="white")
    ax.set_aspect("equal")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", transparent=True, pad_inches=0.02)
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def _make_bar_image(labels: list[str], values: list[float]) -> ImageReader:
    plt = _matplotlib_pyplot()
    fig, ax = plt.subplots(figsize=(5.2, 2.0))
    rev_labels = labels[::-1]
    rev_values = values[::-1]
    bars = ax.barh(rev_labels, rev_values, color=_BAR_COLOR, height=0.48, edgecolor="none")
    for bar, val in zip(bars, rev_values):
        ax.text(bar.get_width() + 0.25, bar.get_y() + bar.get_height() / 2, f"{val:.2f}%", va="center", fontsize=7.5, fontweight="bold")
    ax.set_title("Top 5 Allocation (%)", fontsize=9, fontweight="bold", pad=8)
    ax.set_xlim(0, max(values) * 1.3 if values else 10)
    ax.tick_params(axis="y", labelsize=6.5)
    ax.tick_params(axis="x", labelsize=7)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("Direct Schemes", fontsize=7)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", transparent=True, pad_inches=0.1)
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def _draw_pdf_header(c: PdfCanvas, y: float, report: ClientReport) -> float:
    h = 55
    c.setFillColor(_NAVY)
    c.rect(_LM, y - h, _CW, h, fill=True, stroke=False)
    # gold accent line under the header band
    c.setFillColor(_GOLD)
    c.rect(_LM, y - h - 3, _CW, 3, fill=True, stroke=False)

    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(_LM + 14, y - 20, "Fincart Direct Mutual Fund Strategic Portfolio")
    c.setFont("Helvetica", 7.5)
    c.drawString(_LM + 14, y - 32, "Portfolio Management Services  •  SEBI Reg. INP000006101")
    c.setFont("Helvetica", 7)
    c.drawString(_LM + 14, y - 44, f"Client: {report.client_name}   |   UCC: {report.ucc}")

    right_x = _LM + _CW - 14
    c.setFont("Helvetica", 7)
    c.drawRightString(right_x, y - 12, "FACTSHEET")
    c.setFont("Helvetica-Bold", 13)
    c.drawRightString(right_x, y - 27, REPORT_DATE.strftime("%B %Y"))
    c.setFont("Helvetica", 7)
    c.drawRightString(right_x, y - 39, f"As on {REPORT_DATE.strftime('%d %B %Y')}")

    return y - h


def _draw_pdf_about(c: PdfCanvas, y: float, report: ClientReport) -> float:
    h = 80
    c.setFillColor(_LIGHT_BG)
    c.rect(_LM, y - h, _CW, h, fill=True, stroke=False)

    inception_dt = _inception(report)
    inception_str = inception_dt.strftime("%d %B %Y")

    c.setFillColor(_TEXT_DARK)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(_LM + 12, y - 16, "About the Scheme:")

    c.setFillColor(_TEXT_MED)
    c.setFont("Helvetica-Oblique", 7.5)
    c.drawString(_LM + 12, y - 29, f"Multi-Asset Strategic Allocation  •  Direct Plan Mutual Fund PMS  •  Inception {inception_str}")

    c.setFillColor(_TEXT_DARK)
    c.setFont("Helvetica", 6.8)
    approach = (
        "INVESTMENT APPROACH:  Strategic multi-asset allocation across Indian equity, international equity, debt, gold and arbitrage  •  "
        "Direct-plan-only universe — preserves every bps of return  •  Diversified across market caps and themes: small-cap alpha, thematic equity, "
        "defence and focused large-cap  •  Quarterly rebalancing anchored to a defined benchmark for measurable accountability."
    )
    text_obj = c.beginText(_LM + 12, y - 42)
    text_obj.setFont("Helvetica", 6.5)
    max_w = _CW - 24
    words = approach.split()
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        if c.stringWidth(test, "Helvetica", 6.5) > max_w:
            text_obj.textLine(line)
            line = word
        else:
            line = test
    if line:
        text_obj.textLine(line)
    c.drawText(text_obj)
    return y - h


def _draw_pdf_kpi(c: PdfCanvas, y: float, report: ClientReport) -> float:
    h = 58
    box_w = _CW / 3
    inception_dt = _inception(report)
    gl_frac = (report.total_pl / report.cost_value) if report.cost_value else 0.0

    # (background, label/sub color, default value color)
    box_styles = [
        (_NAVY, _WHITE, _WHITE),
        (_DARK_NAVY, _WHITE, _WHITE),
        (_CREAM, _TEXT_MED, _TEXT_DARK),
    ]
    for i, (bg, _lbl, _val) in enumerate(box_styles):
        bx = _LM + i * box_w
        c.setFillColor(bg)
        c.rect(bx, y - h, box_w, h, fill=True, stroke=False)

    def _draw_kpi_box(idx: int, label: str, value: str, sub: str, val_color=None):
        bg, lbl_color, default_val = box_styles[idx]
        if val_color is None:
            val_color = default_val
        bx = _LM + idx * box_w
        cx = bx + box_w / 2
        c.setFillColor(lbl_color)
        c.setFont("Helvetica", 6.5)
        c.drawCentredString(cx, y - 14, label)
        c.setFillColor(val_color)
        c.setFont("Helvetica-Bold", 15)
        c.drawCentredString(cx, y - 32, value)
        c.setFillColor(lbl_color)
        c.setFont("Helvetica", 6)
        c.drawCentredString(cx, y - 46, sub)

    _draw_kpi_box(0, "INCEPTION", inception_dt.strftime("%d %b %Y"), "Portfolio start")
    _draw_kpi_box(1, "ABSOLUTE GAIN / LOSS", _fmt_pct(gl_frac), _fmt_inr(report.total_pl))
    xirr_val = _fmt_pct(report.xirr) if report.xirr is not None else "N/A"
    bm_val = _fmt_pct(report.benchmark_xirr) if report.benchmark_xirr is not None else "N/A"
    _draw_kpi_box(
        2, "XIRR  •  SINCE INCEPTION", xirr_val, f"vs BSE 500  {bm_val}",
        _pct_color(report.xirr) if report.xirr is not None else _TEXT_MED,
    )

    return y - h


def _draw_pdf_section_header(c: PdfCanvas, y: float, title: str, subtitle: str = "") -> float:
    h = 20
    c.setFillColor(_NAVY)
    c.rect(_LM, y - h, _CW, h, fill=True, stroke=False)
    # gold left edge accent
    c.setFillColor(_GOLD)
    c.rect(_LM, y - h, 4, h, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(_LM + 12, y - 14, title)
    if subtitle:
        c.setFont("Helvetica-Oblique", 7)
        c.drawString(_LM + 12 + c.stringWidth(title, "Helvetica-Bold", 9) + 12, y - 14, subtitle)
    return y - h


def _draw_pdf_strategy_tables(c: PdfCanvas, y: float, report: ClientReport) -> float:
    left_w = _CW * 0.48
    right_w = _CW * 0.52
    inception_dt = _inception(report)

    snapshot_data = [
        ("Portfolio Manager", "Credent Asset Management"),
        ("SEBI PMS Reg. No.", "INP000006101"),
        ("Strategy Type", "Multi-Asset MF Portfolio"),
        ("Benchmark", "BSE 500 (BSE_DLY_BSE500, 1D)"),
        ("Investment Universe", "Direct Plan Mutual Funds"),
        ("Custodian / Fund A/c", "Nuvama Wealth"),
        ("Min. Investment", "Rs 50 Lakh"),
    ]

    c.setFont("Helvetica-Bold", 8.5)
    c.setFillColor(_NAVY)
    c.drawString(_LM + 10, y - 16, "STRATEGY SNAPSHOT")

    row_h = 16
    ty = y - 28
    for label, value in snapshot_data:
        c.setFillColor(_LIGHT_BLUE)
        c.rect(_LM, ty - row_h, left_w, row_h, fill=True, stroke=False)
        c.setStrokeColor(_BORDER)
        c.setLineWidth(0.3)
        c.rect(_LM, ty - row_h, left_w, row_h, fill=False, stroke=True)
        c.setFillColor(_TEXT_DARK)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(_LM + 8, ty - 11, label)
        c.setFont("Helvetica", 7)
        c.drawString(_LM + 120, ty - 11, value)
        ty -= row_h

    rx = _LM + left_w + 8
    perf_w = right_w - 8
    c.setFont("Helvetica-Bold", 8.5)
    c.setFillColor(_NAVY)
    c.drawString(rx + 6, y - 16, "PERFORMANCE (XIRR)")

    hdr_y = y - 28
    c.setFillColor(_NAVY)
    c.rect(rx, hdr_y - 18, perf_w, 18, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(rx + 8, hdr_y - 12, "Returns (XIRR)")
    c.drawRightString(rx + perf_w - 8, hdr_y - 12, "Since Inception*")

    row_y = hdr_y - 18
    for label, val in [("Client Portfolio", report.xirr), ("BSE 500", report.benchmark_xirr)]:
        c.setFillColor(_LIGHT_BLUE)
        c.rect(rx, row_y - 20, perf_w, 20, fill=True, stroke=False)
        c.setStrokeColor(_BORDER)
        c.setLineWidth(0.3)
        c.rect(rx, row_y - 20, perf_w, 20, fill=False, stroke=True)
        c.setFillColor(_TEXT_DARK)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(rx + 8, row_y - 13, label)
        c.setFillColor(_pct_color(val))
        c.setFont("Helvetica-Bold", 8)
        c.drawRightString(rx + perf_w - 8, row_y - 13, _fmt_pct(val))
        row_y -= 20

    c.setFillColor(_TEXT_MED)
    c.setFont("Helvetica-Oblique", 5.8)
    note_y = row_y - 14
    c.drawString(rx + 6, note_y, "*Absolute for period <1Y; ≥1Y annualised.")
    c.drawString(rx + 6, note_y - 9, "Gross of fees, XIRR method.")

    total_h = y - min(ty, note_y - 9) + 4
    return y - max(140, total_h)


def _draw_pdf_allocation(c: PdfCanvas, y: float, report: ClientReport) -> float:
    cat_data = _aggregate_pdf_categories(report)
    if not cat_data:
        return y - 30

    labels = [name for name, _ in cat_data]
    sizes = [weight for _, weight in cat_data]
    scheme_count = _scheme_count(report)

    pie_img = _make_pie_image(labels, sizes)
    pie_w = 155
    pie_h = 155
    c.drawImage(pie_img, _LM + 5, y - pie_h - 5, width=pie_w, height=pie_h, mask="auto")

    tx = _LM + pie_w + 20
    tw = _CW - pie_w - 25

    hdr_h = 16
    c.setFillColor(_NAVY)
    c.rect(tx, y - hdr_h, tw, hdr_h, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(tx + 6, y - 11, "Asset Class")
    c.drawRightString(tx + tw - 6, y - 11, "Weight")

    row_h = 15
    ry = y - hdr_h
    for i, (name, weight) in enumerate(cat_data):
        bg = _LIGHT_BLUE if i % 2 == 0 else _WHITE
        c.setFillColor(bg)
        c.rect(tx, ry - row_h, tw, row_h, fill=True, stroke=False)
        c.setStrokeColor(_BORDER)
        c.setLineWidth(0.3)
        c.rect(tx, ry - row_h, tw, row_h, fill=False, stroke=True)
        c.setFillColor(_TEXT_DARK)
        c.setFont("Helvetica", 7)
        c.drawString(tx + 6, ry - 10, name)
        c.setFont("Helvetica-Bold", 7)
        c.drawRightString(tx + tw - 6, ry - 10, f"{weight * 100:.2f}%")
        ry -= row_h

    c.setFillColor(_NAVY)
    c.rect(tx, ry - row_h, tw, row_h, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(tx + 6, ry - 10, f"TOTAL  •  {scheme_count} underlying schemes")
    c.drawRightString(tx + tw - 6, ry - 10, "100.00%")

    total_h = max(pie_h + 10, (hdr_h + row_h * (len(cat_data) + 1)) + 5)
    return y - total_h


def _draw_pdf_top_holdings(c: PdfCanvas, y: float, report: ClientReport) -> float:
    if not report.top_holdings:
        return y - 30

    labels = [h.scheme_name[:45] for h in report.top_holdings]
    values = [
        (h.current_value / report.current_value * 100) if report.current_value else 0.0
        for h in report.top_holdings
    ]

    bar_img = _make_bar_image(labels, values)
    img_w = _CW - 10
    img_h = 145
    c.drawImage(bar_img, _LM + 5, y - img_h - 5, width=img_w, height=img_h, mask="auto")
    return y - img_h - 10


def _draw_pdf_footer(c: PdfCanvas, page: int, total: int = 2) -> None:
    fy = 18
    c.setStrokeColor(_NAVY)
    c.setLineWidth(0.5)
    c.line(_LM, fy + 10, _LM + _CW, fy + 10)
    c.setFillColor(_TEXT_MED)
    c.setFont("Helvetica", 5.5)
    c.drawString(
        _LM,
        fy,
        "Credent Asset Management Services Pvt Ltd  |  SEBI PMS INP000006101  |  ",
    )
    c.setFont("Helvetica-Bold", 5.5)
    bw = c.stringWidth(
        "Credent Asset Management Services Pvt Ltd  |  SEBI PMS INP000006101  |  ",
        "Helvetica", 5.5,
    )
    c.drawString(_LM + bw, fy, "STRICTLY CONFIDENTIAL — FOR PRIVATE CIRCULATION ONLY")
    c.setFont("Helvetica", 5.5)
    c.drawRightString(_LM + _CW, fy, f"Page {page} of {total}")


def _draw_pdf_disclaimer(c: PdfCanvas) -> None:
    y = _PH - 15
    h = 50
    c.setFillColor(_NAVY)
    c.rect(_LM, y - h, _CW, h, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(_LM + 14, y - 22, "DISCLAIMER & STATUTORY DISCLOSURES")
    c.setFont("Helvetica-Oblique", 7.5)
    c.drawString(_LM + 14, y - 36, f"Credent Fincart Direct Mutual Fund Strategic Portfolio  •  {REPORT_DATE.strftime('%B %Y')}")

    paragraphs = [
        ("General Risk.", "Prospective clients are expected to carefully consider all relevant risk factors — including financial condition, risk-return profile, tax implications, and investment objectives — before making any investment decision. Past performance is not indicative of future results."),
        ("Market & Mutual Fund Risk.", "All products marketed by Credent Asset Management Services Pvt Ltd are subject to market risks including settlement risk, economic risk, political risk, business risk, sector risk and concentration risk. The strategy invests exclusively in direct plan mutual fund schemes; the NAV of underlying schemes is exposed to equity, debt, currency and commodity market fluctuations. Mutual fund investments are subject to market risks; please read all scheme-related documents carefully before investing."),
        ("Performance Attribution.", "Past performance of the Portfolio Manager, the strategy, or similar products does not guarantee future performance. Returns shown are computed using the XIRR methodology. Direct plan returns are gross of all distribution / commission charges but net of fund-level expenses (TER). Individual investor returns may differ owing to timing of inflows/outflows and portfolio composition differences."),
        ("Tax, Legal & Advisory.", "Prospective investors are advised to seek independent professional legal, accounting and tax advice before deciding on any investment or disinvestment. Credent Asset Management may have financial or business interests that could affect the objectivity of the views expressed in this document. The document is for informational purposes only and does not constitute an offer, solicitation, recommendation, or invitation to purchase or sell any investment product."),
        ("SEBI Disclosure.", "This document has neither been approved nor disapproved by SEBI, nor has SEBI certified the accuracy or adequacy of its contents. The portfolio data and performance figures shown are as uploaded on the SEBI / APMI website as on the cover date. For peer comparison within the strategy category, please refer to: www.apmiindia.org/apmi/welcomeiaperformance.htm?action=PMSmenu. Clients have the option to onboard Credent Asset Management PMS either directly or through a SEBI-registered distributor."),
        ("Liability & Confidentiality.", "Credent Asset Management is not liable for any losses arising from investment / disinvestment decisions made on the basis of communications received, nor for losses arising from incorrect or misleading instructions issued by clients, whether oral or written. This document is strictly confidential and intended only for the addressee; unauthorised reproduction, distribution, or use of the contents is prohibited. For the full Disclosure Document and risk factors, please contact the Portfolio Manager or visit www.credentglobal.com."),
    ]

    ty = y - h - 20
    max_w = _CW - 24

    for bold_title, body in paragraphs:
        full_text = f"{bold_title} {body}"
        c.setFont("Helvetica", 6.8)

        lines: list[tuple[str, bool]] = []
        bold_end = len(bold_title)
        words = full_text.split()
        line = ""
        current_pos = 0
        for word in words:
            test = f"{line} {word}".strip()
            if c.stringWidth(test, "Helvetica", 6.8) > max_w:
                lines.append((line, current_pos <= bold_end))
                line = word
            else:
                line = test
            current_pos = len(line)
        if line:
            lines.append((line, False))

        for line_text, _ in lines:
            if line_text.startswith(bold_title):
                c.setFont("Helvetica-Bold", 6.8)
                c.setFillColor(_TEXT_DARK)
                c.drawString(_LM + 12, ty, bold_title)
                rest = line_text[len(bold_title):]
                bw = c.stringWidth(bold_title, "Helvetica-Bold", 6.8)
                c.setFont("Helvetica", 6.8)
                c.drawString(_LM + 12 + bw, ty, rest)
            else:
                c.setFont("Helvetica", 6.8)
                c.setFillColor(_TEXT_DARK)
                c.drawString(_LM + 12, ty, line_text)
            ty -= 9.5

        ty -= 8

    # Footer block
    block_h = 45
    block_y = 50
    c.setFillColor(_NAVY)
    c.rect(_LM, block_y, _CW, block_h, fill=True, stroke=False)
    c.setFillColor(_WHITE)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(_PW / 2, block_y + 32, "Credent Asset Management Services Pvt Ltd")
    c.setFont("Helvetica", 6)
    c.drawCentredString(_PW / 2, block_y + 21, "306, B2, Shubham Centre, Cardinal Gracious Road, Chakala, Andheri (East), Mumbai – 400099, Maharashtra")
    c.drawCentredString(_PW / 2, block_y + 11, "PAN: AAGCB0414J   |   SEBI PMS: INP000006101   |   SEBI RIA: INA000018294   |   www.credentglobal.com")


def generate_client_pdf(report: ClientReport, output_path: Path) -> None:
    if sys.getrecursionlimit() < 50000:
        sys.setrecursionlimit(50000)
    c = PdfCanvas(str(output_path), pagesize=A4)

    # --- Page 1 ---
    y = _PH - 12
    y = _draw_pdf_header(c, y, report)
    y -= 5
    y = _draw_pdf_about(c, y, report)
    y -= 5
    y = _draw_pdf_kpi(c, y, report)
    y -= 6
    y = _draw_pdf_section_header(c, y, "STRATEGY SNAPSHOT & PERFORMANCE", "XIRR  •  Since Inception")
    y = _draw_pdf_strategy_tables(c, y, report)
    y -= 6
    y = _draw_pdf_section_header(c, y, "ASSET ALLOCATION", "Multi-asset, multi-cap diversification")
    y = _draw_pdf_allocation(c, y, report)
    y -= 6
    y = _draw_pdf_section_header(c, y, "TOP 5 HOLDINGS", "By weight")
    y = _draw_pdf_top_holdings(c, y, report)
    _draw_pdf_footer(c, 1)

    c.showPage()

    # --- Page 2: Disclaimer ---
    _draw_pdf_disclaimer(c)
    _draw_pdf_footer(c, 2)

    c.save()


def generate_all_pdfs(reports: list[ClientReport]) -> None:
    PDF_OUTPUT_DIR.mkdir(exist_ok=True)
    count = 0
    for report in reports:
        if report.cost_value <= 0:
            continue
        safe_name = re.sub(r"[^\w\s\-]", "", report.client_name).strip()
        safe_name = re.sub(r"\s+", "_", safe_name)
        pdf_path = PDF_OUTPUT_DIR / f"{safe_name}_Factsheet_{REPORT_DATE.strftime('%b_%Y')}.pdf"
        generate_client_pdf(report, pdf_path)
        count += 1
        print(f"  PDF: {pdf_path.name}")
    print(f"Generated {count} PDF factsheets in {PDF_OUTPUT_DIR}")


# =====================================================================
# MAIN
# =====================================================================

def build_report() -> None:
    source_workbook = load_workbook(TRANSACTION_FILE, data_only=True)
    master = read_master(source_workbook)
    transactions = read_transactions(source_workbook, master)
    bse_prices = read_bse_prices(BSE_500_FILE)
    current_navs, category_overrides = read_current_navs(OPTIONAL_CURRENT_NAV_FILE)

    reports = build_client_reports(master, transactions, bse_prices, current_navs, category_overrides)

    workbook = Workbook()
    write_summary_sheet(workbook, reports)
    used_names = {"Summary"}
    for report in reports:
        write_client_sheet(workbook, report, used_names)
    write_notes_sheet(workbook, used_names, bool(current_navs))
    workbook.save(OUTPUT_FILE)

    print(f"Created Excel: {OUTPUT_FILE}")
    print(f"Clients: {len(reports)}")
    print(f"Transactions used: {len(transactions)}")
    print(f"BSE 500 source: {BSE_500_FILE.name}")
    print(f"Report date: {REPORT_DATE.strftime('%d %B %Y')}")
    print(f"Current NAV source: {'Current_NAVs.xlsx' if current_navs else 'latest observed transaction NAV'}")

    print("\nGenerating PDF factsheets...")
    generate_all_pdfs(reports)


if __name__ == "__main__":
    build_report()
