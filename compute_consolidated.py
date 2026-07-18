"""Consolidated client view computed entirely from the custodian holdings
snapshot CSV (e.g. 'client file.csv').

The snapshot has one row per (client, holding) with cost, market value,
gain/loss and % assets already provided by the custodian — so this view needs
no transaction history and is accurate for every client. Performance metric is
absolute Gain/Loss (₹ and %). Cash-in-hand counts only the literal uninvested
CASH rows (not liquid funds).
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

from create_client_factsheet_report import to_number, clean_text, infer_category

# Display category for each fund, reusing the factsheet taxonomy where possible.
_CASH_SYMBOLS = {"CASH"}


def _split_name_ucc(name_field: str, client_code: str) -> tuple[str, str]:
    """'SOUMIT  BISWAS - CRFM002' -> ('Soumit Biswas', 'CRFM002')."""
    ucc = clean_text(client_code)
    name = clean_text(name_field)
    # Strip a trailing ' - CRFMxxx' if present
    m = re.match(r"^(.*?)\s*-\s*(CRFM\d+)\s*$", name, re.I)
    if m:
        name = m.group(1).strip()
        if not ucc:
            ucc = m.group(2).upper()
    return name.title() if name.isupper() else name, ucc


def compute_consolidated(csv_path: Path, current_navs: dict[str, float] | None = None) -> dict:
    """Build consolidated view. When *current_navs* is supplied, each holding's
    market value is recomputed as units × live NAV (instead of using the CSV's
    stale valuation) so this tab matches the factsheet."""
    current_navs = current_navs or {}
    if not csv_path.exists():
        return {"clients": [], "totals": {}, "report_date": ""}

    clients: dict[str, dict] = {}
    report_date = ""

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ucc = clean_text(row.get("Client code"))
            if not ucc:
                continue
            if not report_date:
                report_date = clean_text(row.get("Holding Date"))

            if ucc not in clients:
                name, ucc_parsed = _split_name_ucc(row.get("NAME", ""), ucc)
                clients[ucc] = {
                    "ucc": ucc_parsed or ucc,
                    "name": name,
                    "cost": 0.0,
                    "market_value": 0.0,
                    "gain_loss": 0.0,
                    "cash": 0.0,
                    "holdings": [],
                }

            c = clients[ucc]
            symbol = clean_text(row.get("SYMBOLCODE")).upper()
            scheme = clean_text(row.get("SYMBOLNAME"))
            isin = clean_text(row.get("ISIN"))
            cost = to_number(row.get("Total Cost")) or 0.0
            mkt = to_number(row.get("Market Value")) or 0.0
            gl = to_number(row.get("Gain/Loss")) or 0.0
            gl_pct = to_number(row.get("Gain/Loss %"))
            pct_assets = to_number(row.get("% Assets")) or 0.0
            units = to_number(row.get("QUANTITY")) or 0.0

            is_cash = symbol in _CASH_SYMBOLS or (not isin and scheme.lower() == "cash")
            if is_cash:
                c["cost"] += cost
                c["market_value"] += mkt
                c["gain_loss"] += gl
                c["cash"] += mkt
                continue  # don't list bare cash as a fund holding

            if isin and isin in current_navs and units > 0:
                mkt = units * current_navs[isin]
                gl = mkt - cost
                gl_pct = (gl / cost * 100) if cost else 0.0

            c["cost"] += cost
            c["market_value"] += mkt
            c["gain_loss"] += gl

            c["holdings"].append({
                "scheme": scheme,
                "isin": isin,
                "category": infer_category(isin, scheme),
                "units": units,
                "cost": cost,
                "market_value": mkt,
                "gain_loss": gl,
                "gain_loss_pct": gl_pct if gl_pct is not None else (gl / cost * 100 if cost else 0.0),
                "pct_assets": pct_assets,
            })

    # Finalise per-client derived figures
    client_list = []
    for c in clients.values():
        c["holdings"].sort(key=lambda h: h["market_value"], reverse=True)
        c["gain_loss_pct"] = (c["gain_loss"] / c["cost"] * 100) if c["cost"] else 0.0
        c["cash_pct"] = (c["cash"] / c["market_value"] * 100) if c["market_value"] else 0.0
        client_list.append(c)

    client_list.sort(key=lambda x: x["market_value"], reverse=True)

    total_cost = sum(c["cost"] for c in client_list)
    total_mkt = sum(c["market_value"] for c in client_list)
    total_gl = sum(c["gain_loss"] for c in client_list)
    total_cash = sum(c["cash"] for c in client_list)

    totals = {
        "n_clients": len(client_list),
        "cost": total_cost,
        "market_value": total_mkt,
        "gain_loss": total_gl,
        "gain_loss_pct": (total_gl / total_cost * 100) if total_cost else 0.0,
        "cash": total_cash,
        "cash_pct": (total_cash / total_mkt * 100) if total_mkt else 0.0,
    }

    from datetime import datetime
    display_date = datetime.today().strftime("%d/%m/%Y") if current_navs else report_date
    return {"clients": client_list, "totals": totals, "report_date": display_date}
