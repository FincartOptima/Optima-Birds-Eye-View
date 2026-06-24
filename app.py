from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import io
import tempfile
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from create_client_factsheet_report import (
    read_master, read_transactions, read_bse_prices,
    read_current_navs, build_client_reports, generate_client_pdf,
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


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload and generate reports"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Only Excel files (.xlsx, .xls) are supported'}), 400

    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Process the file
        print(f"Processing file: {filepath}")
        source_workbook = load_workbook(filepath, data_only=True)
        master = read_master(source_workbook)

        if not master:
            return jsonify({'error': 'No valid client data found in the file'}), 400

        transactions = read_transactions(source_workbook, master)

        # Read BSE prices
        bse_file = Path(__file__).resolve().parent / "BSE_DLY_BSE500, 1D (8).csv"
        if not bse_file.exists():
            return jsonify({'error': f'BSE 500 data file not found: {bse_file}'}), 400

        bse_prices = read_bse_prices(bse_file)
        current_navs, category_overrides = read_current_navs(Path(__file__).resolve().parent / "Current_NAVs.xlsx")

        # Build reports
        reports = build_client_reports(master, transactions, bse_prices, current_navs, category_overrides)

        # Cache reports
        reports_cache['reports'] = reports
        reports_cache['file_path'] = filepath
        reports_cache['upload_time'] = datetime.now().isoformat()

        # Prepare client list
        client_list = [
            {
                'id': i,
                'name': report.client_name,
                'ucc': report.ucc,
                'cost_value': report.cost_value,
                'current_value': report.current_value,
            }
            for i, report in enumerate(reports)
            if report.cost_value > 0  # Only show clients with investments
        ]

        return jsonify({
            'success': True,
            'clients': client_list,
            'total_clients': len(client_list),
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
        'report_date': datetime(2026, 6, 10).strftime('%d %b %Y'),

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

        # All holdings
        'all_holdings': [
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
            }
            for h in report.holdings
        ],
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
        filename = f"{safe_name}_Factsheet_Jun_2026.pdf"

        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print(f"Error generating PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error generating PDF: {str(e)}'}), 500


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
