"""Parser for the custodian account statement (.xls, one sheet per client,
sheet name = numeric account CODE).

Extracts per-client:
  - Corpus deposit dates and amounts (real external cash inflows)
  - Account summary holdings (total cost, total value)
  - Fee deductions (FA fees, custody charges, management fees)
  - Implied cash = deposits - fees - fund_purchases + fund_redemptions

The account CODE in each sheet (e.g. 6777) is mapped to a UCC (e.g. CRFM002)
via the client file CSV's CODE → Client code columns.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import xlrd


@dataclass
class CorpusDeposit:
    date: datetime
    amount: float


@dataclass
class CustodianStatement:
    ucc: str
    client_name: str
    account_code: str
    as_of_date: datetime | None
    deposits: list[CorpusDeposit] = field(default_factory=list)
    total_contribution: float = 0.0
    total_fees: float = 0.0
    holdings_cost: float = 0.0
    holdings_value: float = 0.0
    cash: float = 0.0


_DATE_RE = re.compile(r"(\d{2})/(\d{2})/(\d{4})")


def _text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _to_number(value: Any) -> float | None:
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


def _parse_date(text: str) -> datetime | None:
    m = _DATE_RE.search(text)
    if not m:
        return None
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _parse_sheet(sheet) -> tuple[str, str, CustodianStatement] | None:
    """Parse a single sheet. Returns (account_code, client_name, statement) or None."""
    account_code = None
    client_name = None
    as_of_date = None
    deposits: list[CorpusDeposit] = []
    total_fees = 0.0
    holdings_cost = 0.0
    holdings_value = 0.0

    in_summary = False
    in_transactions = False
    pending_amount: float | None = None

    for r in range(sheet.nrows):
        col0 = _text(sheet.cell_value(r, 0) if sheet.ncols > 0 else None)

        if col0.startswith("As of"):
            parsed = _parse_date(col0)
            if parsed:
                as_of_date = parsed

        if col0.startswith("Account :"):
            parts = col0.split(":", 1)[-1].strip().split(None, 1)
            if parts:
                account_code = parts[0].strip()
                if len(parts) > 1:
                    client_name = parts[1].strip()

        if col0.startswith("Account Summary"):
            in_summary = True
            continue

        # Totals row in Account Summary: col0 empty, has cost/value/units.
        # Stop looking once we hit the transaction ledger header.
        if in_summary and not in_transactions and not col0 and sheet.ncols > 19:
            cost_val = _to_number(sheet.cell_value(r, 16) if sheet.ncols > 16 else None)
            value_val = _to_number(sheet.cell_value(r, 19) if sheet.ncols > 19 else None)
            units_val = _to_number(sheet.cell_value(r, 12) if sheet.ncols > 12 else None)
            if cost_val and value_val and units_val and cost_val > 10000:
                holdings_cost = cost_val
                holdings_value = value_val

        # Transaction ledger section
        if col0 == "Date" and _text(sheet.cell_value(r, 3) if sheet.ncols > 3 else None) == "Transactions":
            in_transactions = True
            continue

        if in_transactions:
            amount_col19 = _to_number(sheet.cell_value(r, 19) if sheet.ncols > 19 else None)
            if amount_col19 is not None and not col0:
                pending_amount = amount_col19
                continue

            if col0 and pending_amount is not None:
                date = _parse_date(col0)
                desc = _text(sheet.cell_value(r, 3) if sheet.ncols > 3 else None)

                if "Corpus Deposits" in desc or "Corpus Received" in desc:
                    if date:
                        deposits.append(CorpusDeposit(date=date, amount=pending_amount))
                elif any(kw in desc for kw in ("Fund Accountant Fees", "Custody Charges", "Management Fees")):
                    total_fees += pending_amount

                pending_amount = None

    if not account_code:
        return None

    total_contribution = sum(d.amount for d in deposits)

    stmt = CustodianStatement(
        ucc="",  # filled later via CODE→UCC mapping
        client_name=client_name or account_code,
        account_code=account_code,
        as_of_date=as_of_date,
        deposits=deposits,
        total_contribution=total_contribution,
        total_fees=total_fees,
        holdings_cost=holdings_cost,
        holdings_value=holdings_value,
        cash=max(total_contribution - total_fees - holdings_cost, 0.0),
    )
    return account_code, client_name or account_code, stmt


def validate_custodian_statement(path: Path) -> list[str]:
    """Return a list of format errors. Empty list = valid."""
    errors: list[str] = []
    try:
        wb = xlrd.open_workbook(str(path))
    except Exception as exc:
        return [f"Could not open as an .xls file: {exc}"]

    if not wb.sheet_names():
        return ["The file has no sheets."]

    found_account = False
    found_deposit = False
    for name in wb.sheet_names():
        sheet = wb.sheet_by_name(name)
        for r in range(min(sheet.nrows, 60)):
            col0 = _text(sheet.cell_value(r, 0) if sheet.ncols > 0 else None)
            if col0.startswith("Account :"):
                found_account = True
            col3 = _text(sheet.cell_value(r, 3) if sheet.ncols > 3 else None)
            if "Corpus Deposits" in col3 or "Corpus Received" in col3:
                found_deposit = True
        if found_account and found_deposit:
            break

    if not found_account:
        errors.append(
            "No client account sheets found — expected at least one sheet "
            "with an 'Account : <code> <name>' row."
        )
    if not found_deposit:
        errors.append("No 'Corpus Deposits' entries found in any sheet.")
    return errors


def read_custodian_statement(
    path: Path,
    code_to_ucc: dict[str, str] | None = None,
) -> tuple[dict[str, CustodianStatement], list[str]]:
    """Parse every client sheet.

    Returns ({ucc: CustodianStatement}, warnings).
    code_to_ucc maps account CODE (sheet name, e.g. "6777") to UCC
    (e.g. "CRFM002"). Built from the client file CSV.
    """
    code_to_ucc = code_to_ucc or {}
    warnings: list[str] = []
    result: dict[str, CustodianStatement] = {}
    wb = xlrd.open_workbook(str(path))

    for name in wb.sheet_names():
        sheet = wb.sheet_by_name(name)
        try:
            parsed = _parse_sheet(sheet)
        except Exception as exc:
            warnings.append(f"Could not parse sheet '{name}': {exc}")
            continue
        if parsed is None:
            warnings.append(f"Sheet '{name}' has no recognizable 'Account :' row — skipped.")
            continue

        code, client_name, stmt = parsed
        ucc = code_to_ucc.get(code)
        if not ucc:
            warnings.append(
                f"Sheet '{name}' (code {code}, {client_name}): no UCC mapping found "
                f"in client file — skipped. Upload the client CSV to enable mapping."
            )
            continue

        stmt.ucc = ucc
        result[ucc] = stmt

    return result, warnings
