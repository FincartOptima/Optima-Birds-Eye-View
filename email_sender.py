"""Sends factsheet emails to clients via Gmail SMTP.

Credentials come from environment variables — never hardcoded, never
committed:
  SMTP_EMAIL         — the sending Gmail address (e.g. reports@yourfirm.com,
                        must be a real Gmail / Google Workspace mailbox)
  SMTP_APP_PASSWORD  — a 16-character Gmail App Password for that address
                        (requires 2-Step Verification enabled on the
                        account; a regular login password will NOT work
                        with smtplib). Generate one at
                        myaccount.google.com/apppasswords.

Nothing in this module sends anything on its own — every send is triggered
by an explicit user action (clicking Send in the Mailing tab).
"""
from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

_SMTP_HOST = "smtp.gmail.com"
_SMTP_PORT = 465

DISCLAIMER_TEXT = (
    "This email and any attachments are confidential and intended solely for the addressee. "
    "Mutual fund investments are subject to market risks; please read all scheme-related documents "
    "carefully before investing. Past performance is not indicative of future results. This "
    "communication is for informational purposes only and does not constitute investment advice, "
    "an offer, or a solicitation to buy or sell any investment product. Credent Asset Management "
    "Services Pvt Ltd | SEBI PMS Reg. INP000006101 | www.credentglobal.com"
)


def is_configured() -> bool:
    return bool(os.environ.get("SMTP_EMAIL") and os.environ.get("SMTP_APP_PASSWORD"))


def send_factsheet_email(to_email: str, client_name: str, pdf_bytes: bytes, pdf_filename: str, report_month: str) -> None:
    """Send one client's factsheet PDF by email. Raises on any failure —
    callers should catch this per-recipient so one bad address doesn't
    abort the whole batch."""
    sender = os.environ.get("SMTP_EMAIL")
    app_password = os.environ.get("SMTP_APP_PASSWORD")
    if not sender or not app_password:
        raise ValueError("Email sending is not configured (SMTP_EMAIL / SMTP_APP_PASSWORD not set).")

    first_name = client_name.split()[0].title() if client_name.split() else client_name

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = f"Your Portfolio Factsheet — {report_month}"

    body = (
        f"Dear {first_name},\n\n"
        f"Please find attached your portfolio factsheet for {report_month}, showing how your "
        f"portfolio has performed over the period.\n\n"
        f"For any queries, please contact your wealth manager.\n\n"
        f"Regards,\n"
        f"Credent Asset Management Services\n\n"
        f"---\n"
        f"{DISCLAIMER_TEXT}"
    )
    msg.attach(MIMEText(body, "plain"))

    attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
    attachment.add_header("Content-Disposition", "attachment", filename=pdf_filename)
    msg.attach(attachment)

    with smtplib.SMTP_SSL(_SMTP_HOST, _SMTP_PORT, timeout=30) as server:
        server.login(sender, app_password)
        server.send_message(msg)
