from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import io
import re
import tempfile
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from collections import defaultdict

from openpyxl import load_workbook
from create_client_factsheet_report import (
    read_master, read_transactions,
    read_current_navs, read_client_file_csv, build_client_reports, generate_client_pdf,
    validate_trade_master, validate_client_csv,
    REPORT_DATE, CATEGORY_ORDER,
)
from custodian_statement import validate_custodian_statement, read_custodian_statement
import csv as csv_mod
import gsheet_data

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max file size

# Use system temp dir so it works both locally and on cloud (Render, Railway, etc.)
_TEMP_DIR = Path(tempfile.gettempdir()) / "bev_uploads"
_TEMP_DIR.mkdir(exist_ok=True)
app.config['UPLOAD_FOLDER'] = str(_TEMP_DIR)

# Global state: store reports and metadata in session
reports_cache = {}
current_file = None


def _format_cell_date(value):
    """Format a raw trade-master cell value (datetime, string, or blank) for
    display without relying on Flask's default datetime JSON encoding."""
    if isinstance(value, datetime):
        return value.strftime('%d %b %Y')
    if value is None:
        return ''
    return str(value).strip()


def _build_all_holdings(report):
    rows = [
        {
            'name': h.scheme_name,
            'category': h.category,
            'isin': h.isin,
            'units': h.units,
            'cost_value': h.cost_value,
            'current_nav': h.current_nav,
            'current_value': h.current_value,
            'allocation_pct': (h.current_value / report.current_value * 100) if report.current_value else 0,
            'unrealized_pl': h.unrealized_pl,
            'realized_pl': h.realized_pl,
            'total_pl': h.total_pl,
            'is_cash': False,
        }
        for h in report.holdings
    ]
    if report.only_cash > 0:
        rows.append({
            'name': 'Cash (Uninvested)',
            'category': 'Only Cash',
            'isin': None,
            'units': None,
            'cost_value': report.only_cash,
            'current_nav': None,
            'current_value': report.only_cash,
            'allocation_pct': (report.only_cash / report.current_value * 100) if report.current_value else 0,
            'unrealized_pl': 0.0,
            'realized_pl': 0.0,
            'total_pl': 0.0,
            'is_cash': True,
        })
    return rows


@app.route('/')
def index():
    return render_template('index.html')


def _fetch_gsheet(url, fmt='csv'):
    """Download a Google Sheet as CSV or XLSX bytes.

    The sheet must be shared with 'Anyone with the link'.
    """
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if m:
        sheet_id = m.group(1)
    elif re.match(r'^[a-zA-Z0-9_-]{20,}$', url.strip()):
        sheet_id = url.strip()
    else:
        raise ValueError(
            'Could not find a Google Sheet ID. Paste the full URL '
            '(e.g. https://docs.google.com/spreadsheets/d/…/edit)')

    export_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format={fmt}'
    req = urllib.request.Request(export_url, headers={'User-Agent': 'BirdsEyeView/1.0'})
    try:
        resp = urllib.request.urlopen(req, timeout=60)
    except urllib.error.HTTPError as e:
        if e.code == 403:
            raise ValueError('Access denied — make sure the sheet is shared with "Anyone with the link"')
        raise ValueError(f'Google Sheets returned HTTP {e.code}')

    data = resp.read()
    if fmt == 'csv' and data[:100].lower().startswith((b'<!doctype', b'<html')):
        raise ValueError(
            'Got an HTML page instead of data — the sheet is probably not shared publicly. '
            'Set sharing to "Anyone with the link".')
    return data


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle upload of the client holdings CSV and/or the trade master xlsx.

    Each source can come from a file upload OR a Google Sheet URL.
    The CSV alone is enough for the Client Consolidated tab. The xlsx (trade
    master) additionally enables the per-client factsheet (XIRR) and the
    Master Dashboard. At least one source must be provided.
    """
    xlsx = request.files.get('file')
    snap = request.files.get('nav_file')
    custodian_file = request.files.get('custodian_file')
    nav_url = request.form.get('nav_sheet_url', '').strip()
    master_url = request.form.get('master_sheet_url', '').strip()

    has_xlsx = bool(xlsx and xlsx.filename)
    has_snap = bool(snap and snap.filename)
    has_custodian = bool(custodian_file and custodian_file.filename)

    if has_xlsx and not xlsx.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'The trade master must be an Excel file (.xlsx, .xls)'}), 400
    if has_snap and not snap.filename.lower().endswith('.csv'):
        return jsonify({'error': 'The client holdings snapshot must be a .csv file'}), 400
    if has_custodian and not custodian_file.filename.lower().endswith('.xls'):
        return jsonify({'error': 'The custodian portfolio statement must be an old-format Excel file (.xls)'}), 400

    if not has_xlsx and not has_snap and not has_custodian and not nav_url and not master_url:
        return jsonify({'error': 'No file or Google Sheet URL provided.'}), 400

    try:
        # Reset prior state
        reports_cache.pop('csv_path', None)
        reports_cache['reports'] = []
        reports_cache['bse_prices'] = []
        reports_cache['failed_transactions'] = []

        # ---- Resolve the holdings CSV source ----
        snap_path = None
        if has_snap:
            snap_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(snap.filename))
            snap.save(snap_path)
        elif nav_url:
            print(f"Fetching holdings CSV from Google Sheet: {nav_url}")
            csv_bytes = _fetch_gsheet(nav_url, 'csv')
            snap_path = os.path.join(app.config['UPLOAD_FOLDER'], 'gsheet_holdings.csv')
            with open(snap_path, 'wb') as f:
                f.write(csv_bytes)
            print(f"Holdings CSV fetched ({len(csv_bytes):,} bytes)")

        current_navs, category_overrides = {}, {}
        if snap_path:
            csv_errors = validate_client_csv(Path(snap_path))
            if csv_errors:
                return jsonify({'error': 'Invalid client file format:\n• ' + '\n• '.join(csv_errors)}), 400
            reports_cache['csv_path'] = snap_path
            current_navs, category_overrides = read_client_file_csv(Path(snap_path))
            print(f"Snapshot NAVs: {len(current_navs)}")

        # ---- Resolve the custodian account statement source (optional) ----
        custodian_data = {}
        if has_custodian:
            custodian_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(custodian_file.filename))
            custodian_file.save(custodian_path)
            custodian_errors = validate_custodian_statement(Path(custodian_path))
            if custodian_errors:
                return jsonify({'error': 'Invalid custodian account statement format:\n• ' + '\n• '.join(custodian_errors)}), 400

            # Build CODE → UCC mapping from the client CSV (if uploaded)
            code_to_ucc: dict[str, str] = {}
            if snap_path:
                with open(snap_path, newline="", encoding="utf-8-sig") as f:
                    for row in csv_mod.DictReader(f):
                        code = (row.get("CODE") or "").strip()
                        ucc = (row.get("Client code") or "").strip()
                        if code and ucc and code not in code_to_ucc:
                            code_to_ucc[code] = ucc

            custodian_data, custodian_warnings = read_custodian_statement(
                Path(custodian_path), code_to_ucc=code_to_ucc,
            )
            print(f"Custodian account statement: {len(custodian_data)} client accounts parsed")
            for w in custodian_warnings:
                print(f"[custodian] {w}")

        # ---- Resolve the trade master XLSX source ----
        xlsx_path = None
        if has_xlsx:
            xlsx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(xlsx.filename))
            xlsx.save(xlsx_path)
        elif master_url:
            print(f"Fetching trade master XLSX from Google Sheet: {master_url}")
            xlsx_bytes = _fetch_gsheet(master_url, 'xlsx')
            xlsx_path = os.path.join(app.config['UPLOAD_FOLDER'], 'gsheet_master.xlsx')
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_bytes)
            print(f"Trade master fetched ({len(xlsx_bytes):,} bytes)")

        client_list = []
        if xlsx_path:
            print(f"Processing trade master: {xlsx_path}")
            source_workbook = load_workbook(xlsx_path, data_only=True)

            xlsx_errors = validate_trade_master(source_workbook)
            if xlsx_errors:
                return jsonify({'error': 'Invalid Trade Master format:\n• ' + '\n• '.join(xlsx_errors)}), 400

            master = read_master(source_workbook)
            if not master:
                return jsonify({'error': 'No valid client data found in the trade master file'}), 400
            transactions, failed_transactions = read_transactions(source_workbook, master)
            reports_cache['failed_transactions'] = failed_transactions

            # BSE 500 benchmark: live from the Google Sheet's 'Benchmark' column
            try:
                bse_prices = gsheet_data.get_live_bse_prices()
                print(f"BSE 500 benchmark: {len(bse_prices)} points from live Google Sheet")
            except Exception as e:
                return jsonify({'error': f'Could not fetch the BSE 500 benchmark from Google Sheets: {e}'}), 400

            # Current NAV precedence (highest wins): custodian snapshot Unit Price
            # > live Google Sheets > local Current_NAVs.xlsx > latest txn NAV (inside build_client_reports).
            snapshot_navs, snapshot_categories = current_navs, category_overrides
            try:
                live_navs, gsheet_warnings = gsheet_data.get_live_current_navs()
            except Exception as e:
                live_navs, gsheet_warnings = {}, [str(e)]
            local_navs, local_categories = read_current_navs(Path(__file__).resolve().parent / "Current_NAVs.xlsx")

            current_navs = {**local_navs, **live_navs, **snapshot_navs}
            category_overrides = {
                **local_categories,
                **gsheet_data.get_category_overrides(),
                **snapshot_categories,
            }
            if live_navs:
                print(f"Live Google Sheet NAVs: {len(live_navs)}")
            for w in gsheet_warnings:
                print(f"[gsheet] {w}")

            reports = build_client_reports(
                master, transactions, bse_prices, current_navs, category_overrides,
                custodian_data=custodian_data,
            )
            reports_cache['reports'] = reports
            reports_cache['bse_prices'] = bse_prices
            reports_cache['file_path'] = xlsx_path

            valid_reports = [r for r in reports if r.cost_value > 0]
            total_portfolio_value = sum(r.current_value for r in valid_reports)
            portfolio_cat_values = defaultdict(float)
            for r in valid_reports:
                for row in r.category_rows:
                    portfolio_cat_values[row['Category']] += row['Current Value']
            reports_cache['portfolio_allocation'] = {
                cat: (portfolio_cat_values.get(cat, 0.0) / total_portfolio_value * 100)
                for cat in CATEGORY_ORDER
                if portfolio_cat_values.get(cat, 0.0) > 0
            } if total_portfolio_value else {}

            client_list = [
                {'id': i, 'name': r.client_name, 'ucc': r.ucc,
                 'cost_value': r.cost_value, 'current_value': r.current_value}
                for i, r in enumerate(reports) if r.cost_value > 0
            ]

        has_any_master = bool(xlsx_path)
        has_any_snap = bool(snap_path)
        reports_cache['upload_time'] = datetime.now().isoformat()
        return jsonify({
            'success': True,
            'clients': client_list,
            'total_clients': len(client_list),
            'has_master': has_any_master,
            'has_snapshot': has_any_snap,
            'has_custodian': bool(custodian_data),
            'message': f'Loaded {len(client_list)} clients successfully'
        })

    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
        msg = str(e)
        if 'not a zip file' in msg.lower() or 'badzipfile' in msg.lower():
            msg = 'The Trade Master file is not a valid Excel (.xlsx) file. Please check you uploaded the correct file.'
        return jsonify({'error': msg}), 400


@app.route('/api/client/<int:client_id>')
def get_client_data(client_id):
    """Get detailed client data and PDF preview"""
    if 'reports' not in reports_cache:
        return jsonify({'error': 'No data loaded. Please upload a file first.'}), 400

    reports = reports_cache['reports']
    valid_ids = [i for i, r in enumerate(reports) if r.cost_value > 0]

    if client_id not in valid_ids:
        return jsonify({'error': 'Invalid client ID'}), 404

    report = reports[client_id]

    # Prepare structured data from report
    inception_date = report.initial_date or (
        min(t.statement_date for t in report.transactions) if report.transactions else datetime.now()
    )

    client_data = {
        'id': client_id,
        'name': report.client_name,
        'ucc': report.ucc,
        'inception_date': inception_date.strftime('%d %b %Y'),
        'report_date': REPORT_DATE.strftime('%d %b %Y'),

        # Key metrics
        'metrics': {
            'cost_value': report.cost_value,
            'current_value': report.current_value,
            'unrealized_pl': report.unrealized_pl,
            'realized_pl': report.realized_pl,
            'total_pl': report.total_pl,
            'portfolio_xirr': report.xirr,
            'simple_return': report.simple_return,
            'benchmark_xirr': report.benchmark_xirr,
            'benchmark_value': report.benchmark_current_value,
            'custodian_cash': report.custodian.cash if report.custodian else None,
            'custodian_contribution': report.custodian.total_contribution if report.custodian else None,
        },

        # Category breakdown
        'categories': [
            {
                'name': row['Category'],
                'cost_value': row['Cost Value'],
                'current_value': row['Current Value'],
                'allocation_pct': row['Allocation %'] * 100,
                'unrealized_pl': row['Unrealized P/L'],
                'realized_pl': row['Realized P/L'],
                'total_pl': row['Total P/L'],
            }
            for row in report.category_rows
        ],

        # Top holdings
        'top_holdings': [
            {
                'rank': i + 1,
                'name': h.scheme_name,
                'category': h.category,
                'units': h.units,
                'cost_value': h.cost_value,
                'current_nav': h.current_nav,
                'current_value': h.current_value,
                'allocation_pct': (h.current_value / report.current_value * 100) if report.current_value else 0,
                'unrealized_pl': h.unrealized_pl,
                'realized_pl': h.realized_pl,
                'total_pl': h.total_pl,
            }
            for i, h in enumerate(report.top_holdings)
        ],

        # All holdings (fund rows + optional cash row)
        'all_holdings': _build_all_holdings(report),

        # Grand total row
        'grand_total': {
            'cost_value': report.cost_value,
            'current_value': report.current_value,
            'total_pl': report.total_pl,
        },

        # Portfolio-wide category allocation for comparison pie chart
        'portfolio_allocation': reports_cache.get('portfolio_allocation', {}),
    }

    return jsonify(client_data)


@app.route('/api/client/<int:client_id>/download_pdf')
def download_client_pdf(client_id):
    """Generate and download PDF for a client"""
    if 'reports' not in reports_cache:
        return jsonify({'error': 'No data loaded. Please upload a file first.'}), 400

    reports = reports_cache['reports']
    valid_ids = [i for i, r in enumerate(reports) if r.cost_value > 0]

    if client_id not in valid_ids:
        return jsonify({'error': 'Invalid client ID'}), 404

    try:
        report = reports[client_id]

        # Write PDF to a system temp file (works on cloud too)
        temp_pdf_path = str(_TEMP_DIR / f"temp_{client_id}.pdf")
        generate_client_pdf(report, Path(temp_pdf_path))

        # Read generated PDF and return
        with open(temp_pdf_path, 'rb') as f:
            pdf_buffer = io.BytesIO(f.read())

        # Clean up temp file
        os.remove(temp_pdf_path)

        pdf_buffer.seek(0)
        safe_name = report.client_name.replace(' ', '_').replace('/', '_')
        filename = f"{safe_name}_Factsheet_{REPORT_DATE.strftime('%b_%Y')}.pdf"

        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(f"=== PDF GENERATION ERROR ===\n{tb}\n===========================")
        # Find the deepest "File ..." frame that lives in our own code so
        # the error modal shows where the failure originated.
        our_frames = [
            line.strip() for line in tb.splitlines()
            if line.lstrip().startswith('File "') and 'site-packages' not in line
        ]
        detail = our_frames[-1] if our_frames else ''
        return jsonify({'error': f'Error generating PDF: {str(e)}\n{detail}'}), 500


@app.route('/api/master')
def get_master():
    """Return aggregated master dashboard data."""
    if 'reports' not in reports_cache:
        return jsonify({'error': 'No data loaded. Please upload a file first.'}), 400

    try:
        from compute_master import get_master_data
        from create_client_factsheet_report import REPORT_DATE
        data = get_master_data(
            reports_cache['reports'],
            reports_cache['bse_prices'],
            REPORT_DATE,
        )
        return jsonify(data)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/failed-transactions')
def get_failed_transactions():
    """Return trade-master rows with a blank Value column (column K), grouped
    by client. These are excluded from every other calculation."""
    if 'reports' not in reports_cache:
        return jsonify({'error': 'No data loaded. Please upload a file first.'}), 400

    failed = reports_cache.get('failed_transactions', [])

    by_client = defaultdict(list)
    for t in failed:
        by_client[t.ucc].append({
            'ucc': t.ucc,
            'client_name': t.client_name,
            'value_date': _format_cell_date(t.value_date),
            'scheme_name': t.scheme_name,
            'amount_units': t.amount_units,
            'reason': t.reason,
        })

    overview = sorted(
        (
            {'ucc': ucc, 'client_name': rows[0]['client_name'], 'count': len(rows)}
            for ucc, rows in by_client.items()
        ),
        key=lambda r: r['client_name'].lower(),
    )

    return jsonify({
        'overview': overview,
        'by_client': by_client,
        'total_count': len(failed),
    })


@app.route('/api/consolidated')
def get_consolidated():
    """Return the CSV-driven consolidated client view (gain/loss + cash)."""
    csv_path = reports_cache.get('csv_path')
    if not csv_path or not os.path.exists(csv_path):
        return jsonify({'error': 'No client holdings snapshot (CSV) was uploaded. Upload the client file to see the consolidated view.'}), 400
    try:
        from compute_consolidated import compute_consolidated
        return jsonify(compute_consolidated(Path(csv_path)))
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def format_currency(value):
    """Format value as INR currency"""
    if not value:
        return "₹0"
    if abs(value) >= 1e7:
        return f"₹{value/1e7:.2f}Cr"
    if abs(value) >= 1e5:
        return f"₹{value/1e5:.2f}L"
    return f"₹{value:,.0f}"


def format_percentage(value):
    """Format value as percentage"""
    if value is None:
        return "N/A"
    return f"{value*100:.2f}%"


app.jinja_env.filters['currency'] = format_currency
app.jinja_env.filters['percentage'] = format_percentage


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(debug=debug, host='0.0.0.0', port=port)
