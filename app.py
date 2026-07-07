from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import io
import tempfile
from pathlib import Path
from datetime import datetime
from collections import defaultdict

from openpyxl import load_workbook
from create_client_factsheet_report import (
    read_master, read_transactions, read_bse_prices,
    read_current_navs, read_client_file_csv, build_client_reports, generate_client_pdf,
    REPORT_DATE, CATEGORY_ORDER,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max file size

# Use system temp dir so it works both locally and on cloud (Render, Railway, etc.)
_TEMP_DIR = Path(tempfile.gettempdir()) / "bev_uploads"
_TEMP_DIR.mkdir(exist_ok=True)
app.config['UPLOAD_FOLDER'] = str(_TEMP_DIR)

# Global state: store reports and metadata in session
reports_cache = {}
current_file = None


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


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle upload of the client holdings CSV and/or the trade master xlsx.

    The CSV alone is enough for the Client Consolidated tab. The xlsx (trade
    master) additionally enables the per-client factsheet (XIRR) and the
    Master Dashboard. At least one file must be provided.
    """
    xlsx = request.files.get('file')
    snap = request.files.get('nav_file')
    has_xlsx = bool(xlsx and xlsx.filename)
    has_snap = bool(snap and snap.filename)

    if not has_xlsx and not has_snap:
        return jsonify({'error': 'No file provided. Upload the client holdings CSV (and optionally the trade master).'}), 400
    if has_xlsx and not xlsx.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'The trade master must be an Excel file (.xlsx, .xls)'}), 400
    if has_snap and not snap.filename.lower().endswith('.csv'):
        return jsonify({'error': 'The client holdings snapshot must be a .csv file'}), 400

    try:
        # Reset prior state
        reports_cache.pop('csv_path', None)
        reports_cache['reports'] = []
        reports_cache['bse_prices'] = []

        # Save and parse the holdings snapshot (current NAVs + consolidated view)
        current_navs, category_overrides = {}, {}
        if has_snap:
            snap_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(snap.filename))
            snap.save(snap_path)
            reports_cache['csv_path'] = snap_path
            current_navs, category_overrides = read_client_file_csv(Path(snap_path))
            print(f"Snapshot NAVs: {len(current_navs)}")

        client_list = []
        if has_xlsx:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(xlsx.filename))
            xlsx.save(filepath)
            print(f"Processing trade master: {filepath}")
            source_workbook = load_workbook(filepath, data_only=True)
            master = read_master(source_workbook)
            if not master:
                return jsonify({'error': 'No valid client data found in the trade master file'}), 400
            transactions = read_transactions(source_workbook, master)

            bse_file = Path(__file__).resolve().parent / "BSE_DLY_BSE500, 1D (8).csv"
            if not bse_file.exists():
                return jsonify({'error': f'BSE 500 data file not found: {bse_file}'}), 400
            bse_prices = read_bse_prices(bse_file)

            # Fall back to bundled Current_NAVs.xlsx only if no snapshot NAVs
            if not current_navs:
                current_navs, category_overrides = read_current_navs(Path(__file__).resolve().parent / "Current_NAVs.xlsx")

            reports = build_client_reports(master, transactions, bse_prices, current_navs, category_overrides)
            reports_cache['reports'] = reports
            reports_cache['bse_prices'] = bse_prices
            reports_cache['file_path'] = filepath

            # Portfolio-wide category allocation (used by client pie chart comparison)
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

        reports_cache['upload_time'] = datetime.now().isoformat()
        return jsonify({
            'success': True,
            'clients': client_list,
            'total_clients': len(client_list),
            'has_master': has_xlsx,
            'has_snapshot': has_snap,
            'message': f'Loaded {len(client_list)} clients successfully'
        })

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500


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
            'benchmark_xirr': report.benchmark_xirr,
            'benchmark_value': report.benchmark_current_value,
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
