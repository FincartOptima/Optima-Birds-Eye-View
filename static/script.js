// ============================================================
// Global State
// ============================================================

let clientsList = [];
let currentClientId = null;
let currentClientData = null;
let activeTab = 'client';
let masterDataLoaded = false;
let masterData = null;
let consolidatedLoaded = false;
let consolidatedData = null;
let hasMasterFile = false;

// ============================================================
// Utility Functions
// ============================================================

function formatCurrency(value) {
    if (!value) return '₹0';
    if (Math.abs(value) >= 1e7) return `₹${(value / 1e7).toFixed(2)}Cr`;
    if (Math.abs(value) >= 1e5) return `₹${(value / 1e5).toFixed(2)}L`;
    return `₹${value.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`;
}

function formatPercentage(value) {
    if (value === null || value === undefined) return 'N/A';
    return `${(value * 100).toFixed(2)}%`;
}

function formatNumber(value) {
    if (!value) return '0';
    return value.toLocaleString('en-IN', { maximumFractionDigits: 2 });
}

function showLoading(message = 'Processing...') {
    document.getElementById('loadingText').textContent = message;
    document.getElementById('loadingModal').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingModal').style.display = 'none';
}

function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorModal').style.display = 'flex';
}

function closeErrorModal() {
    document.getElementById('errorModal').style.display = 'none';
}

// ============================================================
// File Upload
// ============================================================

document.getElementById('csvInput').addEventListener('change', (e) => {
    const f = e.target.files[0];
    document.getElementById('csvPicked').textContent = f ? f.name : 'No file';
});
document.getElementById('fileInput').addEventListener('change', (e) => {
    const f = e.target.files[0];
    document.getElementById('xlsxPicked').textContent = f ? f.name : 'No file';
});

async function processUpload() {
    const csvFile = document.getElementById('csvInput').files[0];
    const xlsxFile = document.getElementById('fileInput').files[0];

    if (!csvFile && !xlsxFile) {
        showError('Please choose at least the Client Holdings CSV file.');
        return;
    }

    showLoading('Uploading and processing files...');

    const formData = new FormData();
    if (xlsxFile) formData.append('file', xlsxFile);
    if (csvFile)  formData.append('nav_file', csvFile);

    try {
        const response = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Upload failed');

        clientsList = data.clients || [];
        hasMasterFile = !!xlsxFile && clientsList.length > 0;
        document.getElementById('fileInfo').textContent = `✓ Loaded ${data.total_clients || 0} clients`;

        initializeUI();
        hideLoading();
    } catch (error) {
        hideLoading();
        showError(error.message);
    }
}

// ============================================================
// UI Initialization
// ============================================================

function initializeUI() {
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('mainSection').style.display = 'block';

    // Populate client select
    const select = document.getElementById('clientSelect');
    select.innerHTML = '';
    clientsList.forEach((client) => {
        const option = document.createElement('option');
        option.value = client.id;
        option.textContent = `${client.name} (${client.ucc})`;
        select.appendChild(option);
    });

    if (clientsList.length > 0) {
        select.value = clientsList[0].id;
        loadClient();
    }

    setupClientSearch();

    // Enable/disable tabs that require the trade master
    const clientBtn = document.getElementById('tabClientBtn');
    const masterBtn = document.getElementById('tabMasterBtn');
    clientBtn.disabled = !hasMasterFile;
    masterBtn.disabled = !hasMasterFile;
    clientBtn.classList.toggle('disabled', !hasMasterFile);
    masterBtn.classList.toggle('disabled', !hasMasterFile);

    // Land on Client Consolidated (always available from the CSV)
    switchTab('consolidated');
}

function setupClientSearch() {
    const searchInput = document.getElementById('clientSearch');
    const select = document.getElementById('clientSelect');

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        Array.from(select.options).forEach((option) => {
            const text = option.textContent.toLowerCase();
            option.style.display = text.includes(query) ? '' : 'none';
        });
        const firstVisible = Array.from(select.options).find(opt => opt.style.display !== 'none');
        if (firstVisible) select.value = firstVisible.value;
    });
}

// ============================================================
// Tab Switching
// ============================================================

function switchTab(tab) {
    if ((tab === 'client' || tab === 'master') && !hasMasterFile) {
        showError('This view needs the Trade Allocation Master file. Re-upload including the Excel file to enable it.');
        return;
    }
    activeTab = tab;

    const tabs = {
        client:       { btn: 'tabClientBtn',       content: 'clientTabContent' },
        master:       { btn: 'tabMasterBtn',       content: 'masterTabContent' },
        consolidated: { btn: 'tabConsolidatedBtn', content: 'consolidatedTabContent' },
    };

    for (const [key, ids] of Object.entries(tabs)) {
        const btn = document.getElementById(ids.btn);
        const content = document.getElementById(ids.content);
        const isActive = key === tab;
        btn.classList.toggle('active', isActive);
        content.style.display = isActive ? 'block' : 'none';
    }

    // The client/search toolbar only applies to the per-client factsheet
    const showToolbar = tab === 'client';
    document.getElementById('toolbarCenter').style.display = showToolbar ? 'flex' : 'none';
    document.getElementById('toolbarRight').style.display  = showToolbar ? 'flex' : 'none';

    if (tab === 'master' && !masterDataLoaded) loadMasterData();
    if (tab === 'consolidated' && !consolidatedLoaded) loadConsolidated();
}

// ============================================================
// Load Client Data
// ============================================================

async function loadClient() {
    const select = document.getElementById('clientSelect');
    const clientId = parseInt(select.value);
    if (!clientId && clientId !== 0) return;

    showLoading('Loading client data...');

    try {
        const response = await fetch(`/api/client/${clientId}`);
        if (!response.ok) throw new Error('Failed to load client data');

        currentClientData = await response.json();
        currentClientId = clientId;
        renderClientData();
        hideLoading();

    } catch (error) {
        hideLoading();
        showError(error.message);
    }
}

// ============================================================
// Render Client Data
// ============================================================

function renderClientData() {
    if (!currentClientData) return;
    const data = currentClientData;

    document.getElementById('clientName').textContent = data.name;
    document.getElementById('clientUCC').textContent = `UCC: ${data.ucc}`;
    document.getElementById('inceptionDate').textContent = data.inception_date;
    document.getElementById('reportDate').textContent = data.report_date;

    document.getElementById('costValue').textContent = formatCurrency(data.metrics.cost_value);
    document.getElementById('currentValue').textContent = formatCurrency(data.metrics.current_value);

    const totalPLEl = document.getElementById('totalPL');
    totalPLEl.textContent = formatCurrency(data.metrics.total_pl);
    totalPLEl.className = 'fs-kpi-value ' + glClass(data.metrics.total_pl);

    const glPct = data.metrics.cost_value ? (data.metrics.total_pl / data.metrics.cost_value) * 100 : 0;
    const glPctEl = document.getElementById('gainLossPct');
    glPctEl.textContent = signedPct(glPct);
    glPctEl.className = 'fs-kpi-sub ' + glClass(glPct);

    const xirr = document.getElementById('portfolioXIRR');
    if (data.metrics.portfolio_xirr !== null) {
        xirr.textContent = formatPercentage(data.metrics.portfolio_xirr);
        xirr.className = 'fs-kpi-value ' + glClass(data.metrics.portfolio_xirr);
    } else {
        xirr.textContent = 'N/A';
        xirr.className = 'fs-kpi-value';
    }

    const bxirr = document.getElementById('benchmarkXIRR');
    bxirr.textContent = 'BSE 500: ' + (data.metrics.benchmark_xirr !== null ? formatPercentage(data.metrics.benchmark_xirr) : 'N/A');

    renderCategoriesTable(data.categories);
    renderTopHoldingsTable(data.top_holdings);
    renderAllHoldingsTable(data.all_holdings);
}

function renderCategoriesTable(categories) {
    const tbody = document.getElementById('categoriesTable');
    if (!categories || categories.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="loading">No data available</td></tr>';
        return;
    }
    tbody.innerHTML = categories.map(cat => `
        <tr>
            <td><strong>${cat.name}</strong></td>
            <td class="number">${formatCurrency(cat.cost_value)}</td>
            <td class="number">${formatCurrency(cat.current_value)}</td>
            <td class="number"><strong>${cat.allocation_pct.toFixed(2)}%</strong></td>
            <td class="number ${cat.unrealized_pl >= 0 ? 'positive' : 'negative'}">${formatCurrency(cat.unrealized_pl)}</td>
            <td class="number ${cat.total_pl >= 0 ? 'positive' : 'negative'}">${formatCurrency(cat.total_pl)}</td>
        </tr>
    `).join('');
}

function renderTopHoldingsTable(holdings) {
    const tbody = document.getElementById('topHoldingsTable');
    if (!holdings || holdings.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="loading">No holdings</td></tr>';
        return;
    }
    tbody.innerHTML = holdings.map(h => `
        <tr>
            <td><strong>${h.rank}</strong></td>
            <td>${h.name}</td>
            <td>${h.category}</td>
            <td class="number">${formatCurrency(h.current_value)}</td>
            <td class="number"><strong>${h.allocation_pct.toFixed(2)}%</strong></td>
            <td class="number ${h.total_pl >= 0 ? 'positive' : 'negative'}">${formatCurrency(h.total_pl)}</td>
        </tr>
    `).join('');
}

function renderAllHoldingsTable(holdings) {
    const tbody = document.getElementById('allHoldingsTable');
    if (!holdings || holdings.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="loading">No holdings</td></tr>';
        return;
    }
    tbody.innerHTML = holdings.map(h => `
        <tr>
            <td>${h.name}</td>
            <td>${h.category}</td>
            <td><code>${h.isin}</code></td>
            <td class="number">${formatNumber(h.units)}</td>
            <td class="number">${formatCurrency(h.cost_value)}</td>
            <td class="number">${formatNumber(h.current_nav)}</td>
            <td class="number">${formatCurrency(h.current_value)}</td>
            <td class="number"><strong>${h.allocation_pct.toFixed(2)}%</strong></td>
            <td class="number ${h.total_pl >= 0 ? 'positive' : 'negative'}">${formatCurrency(h.total_pl)}</td>
        </tr>
    `).join('');
}

// ============================================================
// Download PDF
// ============================================================

async function downloadPDF() {
    if (currentClientId === null) { showError('Please select a client first'); return; }

    showLoading('Generating PDF...');
    try {
        const response = await fetch(`/api/client/${currentClientId}/download_pdf`);
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Failed to generate PDF');
        }
        const contentDisposition = response.headers.get('content-disposition');
        let filename = `factsheet_${currentClientId}.pdf`;
        if (contentDisposition) {
            const m = contentDisposition.match(/filename="?([^"]+)"?/);
            if (m) filename = m[1];
        }
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        hideLoading();
    } catch (error) {
        hideLoading();
        showError(error.message);
    }
}

// ============================================================
// Reset App
// ============================================================

function resetApp() {
    clientsList = [];
    currentClientId = null;
    currentClientData = null;
    masterDataLoaded = false;
    masterData = null;
    consolidatedLoaded = false;
    consolidatedData = null;
    hasMasterFile = false;

    document.getElementById('uploadSection').style.display = 'block';
    document.getElementById('mainSection').style.display = 'none';
    document.getElementById('fileInfo').textContent = '';
    document.getElementById('fileInput').value = '';
    document.getElementById('csvInput').value = '';
    document.getElementById('csvPicked').textContent = 'No file';
    document.getElementById('xlsxPicked').textContent = 'No file';
    document.getElementById('clientSearch').value = '';
}

// ============================================================
// Master Dashboard — Load
// ============================================================

async function loadMasterData() {
    showLoading('Loading master dashboard data...');
    try {
        const response = await fetch('/api/master');
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || 'Failed to load master data');
        }
        masterData = await response.json();
        masterDataLoaded = true;
        renderMasterDashboard();
        hideLoading();
    } catch (error) {
        hideLoading();
        showError(error.message);
    }
}

// ============================================================
// Master Dashboard — Render
// ============================================================

function renderMasterDashboard() {
    if (!masterData) return;
    document.getElementById('masterReportDate').textContent = masterData.report_date;
    renderFundPerfTable();
    renderCatPerfTable();
    renderCashSummary();
    renderFundMatrix();
    renderCatMatrix();
}

/* Return cell: green positive, red negative, gray N/A */
function returnCell(val) {
    if (val === null || val === undefined) {
        return '<td class="number na-cell">N/A</td>';
    }
    const cls  = val >= 0 ? 'positive' : 'negative';
    const sign = val >= 0 ? '+' : '';
    return `<td class="number ${cls}">${sign}${val.toFixed(2)}%</td>`;
}

function renderFundPerfTable() {
    const periods = masterData.period_labels;
    const bse = masterData.bse_returns;

    document.getElementById('fundPerfHead').innerHTML = `
        <tr>
            <th style="min-width:200px">Fund</th>
            ${periods.map(p => `<th style="min-width:70px;text-align:right">${p}</th>`).join('')}
        </tr>`;

    let html = '';
    let currentCat = null;

    for (const fund of masterData.fund_performance) {
        if (fund.category !== currentCat) {
            currentCat = fund.category;
            html += `<tr class="cat-group-header"><td colspan="${1 + periods.length}">${currentCat}</td></tr>`;
        }
        html += `<tr>
            <td class="fund-name-cell">${fund.name}</td>
            ${periods.map(p => returnCell(fund.returns[p])).join('')}
        </tr>`;
    }

    html += `<tr class="bse-row">
        <td><strong>BSE 500 (Benchmark)</strong></td>
        ${periods.map(p => returnCell(bse[p])).join('')}
    </tr>`;

    document.getElementById('fundPerfBody').innerHTML = html;
}

function renderCatPerfTable() {
    const periods = masterData.period_labels;
    const bse = masterData.bse_returns;

    document.getElementById('catPerfHead').innerHTML = `
        <tr>
            <th style="min-width:180px">Category</th>
            ${periods.map(p => `<th style="min-width:70px;text-align:right">${p}</th>`).join('')}
        </tr>`;

    let html = masterData.category_performance.map(cat => `
        <tr>
            <td><strong>${cat.name}</strong></td>
            ${periods.map(p => returnCell(cat.returns[p])).join('')}
        </tr>
    `).join('');

    html += `<tr class="bse-row">
        <td><strong>BSE 500 (Benchmark)</strong></td>
        ${periods.map(p => returnCell(bse[p])).join('')}
    </tr>`;

    document.getElementById('catPerfBody').innerHTML = html;
}

function renderCashSummary() {
    const cs = masterData.cash_summary;

    document.getElementById('cashSummaryDiv').innerHTML = `
        <div class="cash-summary-cards">
            <div class="cash-card">
                <span class="cash-label">Total Uninvested Cash</span>
                <span class="cash-amount">${formatCurrency(cs.total_cash)}</span>
                <span class="cash-pct">${cs.cash_pct.toFixed(2)}% of AUM</span>
            </div>
            <div class="cash-card">
                <span class="cash-label">Total AUM</span>
                <span class="cash-amount">${formatCurrency(cs.total_aum)}</span>
                <span class="cash-pct">&nbsp;</span>
            </div>
        </div>
        <div class="table-responsive" style="margin-top:20px">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Client</th>
                        <th>Portfolio Value</th>
                        <th>Uninvested Cash</th>
                        <th>Cash %</th>
                    </tr>
                </thead>
                <tbody>
                    ${cs.per_client.map(c => `
                        <tr>
                            <td>${c.client}</td>
                            <td class="number">${formatCurrency(c.portfolio_value)}</td>
                            <td class="number">${formatCurrency(c.cash_amount)}</td>
                            <td class="number">${c.cash_pct.toFixed(2)}%</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>`;
}

/* Heat-map cell: white → navy gradient based on value magnitude */
function heatCell(val, max) {
    if (!val || val === 0) return '<td class="matrix-cell zero-cell">—</td>';
    const t   = Math.min(val / Math.max(max, 1), 1);
    const r   = Math.round(235 + (20 - 235) * t);
    const g   = Math.round(240 + (54 - 240) * t);
    const b   = Math.round(247 + (92 - 247) * t);
    const bg  = `rgb(${r},${g},${b})`;
    const fg  = t > 0.55 ? '#ffffff' : '#14365C';
    return `<td class="matrix-cell" style="background:${bg};color:${fg}">${val.toFixed(1)}%</td>`;
}

function renderFundMatrix() {
    const { clients, funds, data } = masterData.client_fund_matrix;

    let max = 0;
    for (const vals of Object.values(data))
        for (const v of vals) if (v > max) max = v;

    document.getElementById('fundMatrixHead').innerHTML = `
        <tr>
            <th style="min-width:200px">Fund</th>
            ${clients.map(c => `<th class="client-col">${c}</th>`).join('')}
        </tr>`;

    document.getElementById('fundMatrixBody').innerHTML = funds.map(fund => `
        <tr>
            <td class="fund-name-cell" title="${fund}">${fund}</td>
            ${data[fund].map(v => heatCell(v, max)).join('')}
        </tr>
    `).join('');
}

function renderCatMatrix() {
    const { clients, categories, data } = masterData.client_category_matrix;

    let max = 0;
    for (const vals of Object.values(data))
        for (const v of vals) if (v > max) max = v;

    document.getElementById('catMatrixHead').innerHTML = `
        <tr>
            <th style="min-width:180px">Category</th>
            ${clients.map(c => `<th class="client-col">${c}</th>`).join('')}
        </tr>`;

    document.getElementById('catMatrixBody').innerHTML = categories.map(cat => `
        <tr>
            <td><strong>${cat}</strong></td>
            ${(data[cat] || []).map(v => heatCell(v, max)).join('')}
        </tr>
    `).join('');
}

// ============================================================
// Client Consolidated — Load & Render
// ============================================================

async function loadConsolidated() {
    showLoading('Loading consolidated view...');
    try {
        const response = await fetch('/api/consolidated');
        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.error || 'Failed to load consolidated data');
        }
        consolidatedData = await response.json();
        consolidatedLoaded = true;
        renderConsolidated();
        hideLoading();
    } catch (error) {
        hideLoading();
        showError(error.message);
    }
}

function glClass(v) { return v >= 0 ? 'positive' : 'negative'; }
function signedPct(v) { return `${v >= 0 ? '+' : ''}${v.toFixed(2)}%`; }

function renderConsolidated() {
    if (!consolidatedData) return;
    const { clients, totals, report_date } = consolidatedData;

    document.getElementById('consReportDate').textContent = report_date || '—';

    // KPI cards
    document.getElementById('consKpis').innerHTML = `
        <div class="kpi-tile">
            <span class="kpi-tile-label">Clients</span>
            <span class="kpi-tile-value">${totals.n_clients}</span>
        </div>
        <div class="kpi-tile">
            <span class="kpi-tile-label">Total Invested</span>
            <span class="kpi-tile-value">${formatCurrency(totals.cost)}</span>
        </div>
        <div class="kpi-tile">
            <span class="kpi-tile-label">Current Value</span>
            <span class="kpi-tile-value">${formatCurrency(totals.market_value)}</span>
        </div>
        <div class="kpi-tile">
            <span class="kpi-tile-label">Total Gain / Loss</span>
            <span class="kpi-tile-value ${glClass(totals.gain_loss)}">${formatCurrency(totals.gain_loss)}</span>
            <span class="kpi-tile-sub ${glClass(totals.gain_loss_pct)}">${signedPct(totals.gain_loss_pct)}</span>
        </div>
        <div class="kpi-tile">
            <span class="kpi-tile-label">Total Cash in Hand</span>
            <span class="kpi-tile-value">${formatCurrency(totals.cash)}</span>
            <span class="kpi-tile-sub">${totals.cash_pct.toFixed(2)}% of AUM</span>
        </div>`;

    // Per-client rows + hidden detail rows
    let html = '';
    clients.forEach((c, idx) => {
        html += `
            <tr class="cons-client-row" onclick="toggleClientDetail(${idx})">
                <td class="exp-cell"><span class="exp-caret" id="caret-${idx}">▸</span></td>
                <td><strong>${c.name}</strong><br><span class="muted">${c.ucc}</span></td>
                <td class="num">${formatCurrency(c.cost)}</td>
                <td class="num">${formatCurrency(c.market_value)}</td>
                <td class="num ${glClass(c.gain_loss)}">${formatCurrency(c.gain_loss)}</td>
                <td class="num ${glClass(c.gain_loss_pct)}">${signedPct(c.gain_loss_pct)}</td>
                <td class="num">${formatCurrency(c.cash)}</td>
                <td class="num">${c.cash_pct.toFixed(2)}%</td>
            </tr>
            <tr class="cons-detail-row" id="detail-${idx}" style="display:none">
                <td></td>
                <td colspan="7">${renderClientDetail(c)}</td>
            </tr>`;
    });
    document.getElementById('consBody').innerHTML = html ||
        '<tr><td colspan="8" class="loading">No client data</td></tr>';
}

function renderClientDetail(c) {
    if (!c.holdings || c.holdings.length === 0) {
        return `<div class="detail-empty">No fund holdings — entire portfolio (${formatCurrency(c.cash)}) is in cash.</div>`;
    }
    const rows = c.holdings.map(h => `
        <tr>
            <td>${h.scheme}</td>
            <td>${h.category}</td>
            <td class="num">${h.pct_assets.toFixed(2)}%</td>
            <td class="num">${formatCurrency(h.cost)}</td>
            <td class="num">${formatCurrency(h.market_value)}</td>
            <td class="num ${glClass(h.gain_loss)}">${formatCurrency(h.gain_loss)}</td>
            <td class="num ${glClass(h.gain_loss_pct)}">${signedPct(h.gain_loss_pct)}</td>
        </tr>`).join('');
    return `
        <table class="data-table detail-table">
            <thead>
                <tr>
                    <th>Fund</th><th>Category</th><th class="num">% Assets</th>
                    <th class="num">Invested</th><th class="num">Current Value</th>
                    <th class="num">Gain / Loss</th><th class="num">Gain / Loss %</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>`;
}

function toggleClientDetail(idx) {
    const row = document.getElementById(`detail-${idx}`);
    const caret = document.getElementById(`caret-${idx}`);
    const open = row.style.display !== 'none';
    row.style.display = open ? 'none' : 'table-row';
    caret.textContent = open ? '▸' : '▾';
}

// ============================================================
// Init
// ============================================================

window.addEventListener('DOMContentLoaded', () => {
    console.log('Client Factsheet Viewer loaded');
});
