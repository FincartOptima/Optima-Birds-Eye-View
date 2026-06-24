// ============================================================
// Global State
// ============================================================

let clientsList = [];
let currentClientId = null;
let currentClientData = null;

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

document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    showLoading('Uploading and processing file...');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || 'Upload failed');
        }

        clientsList = data.clients;
        document.getElementById('fileInfo').textContent = `✓ Loaded ${data.total_clients} clients`;

        // Initialize UI
        initializeUI();

        hideLoading();

    } catch (error) {
        hideLoading();
        showError(error.message);
    }
});

// ============================================================
// UI Initialization
// ============================================================

function initializeUI() {
    // Hide upload section, show main section
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('mainSection').style.display = 'block';

    // Populate client select
    const select = document.getElementById('clientSelect');
    select.innerHTML = '';

    clientsList.forEach((client, index) => {
        const option = document.createElement('option');
        option.value = client.id;
        option.textContent = `${client.name} (${client.ucc})`;
        select.appendChild(option);
    });

    // Load first client by default
    if (clientsList.length > 0) {
        select.value = clientsList[0].id;
        loadClient();
    }

    // Setup search functionality
    setupClientSearch();
}

function setupClientSearch() {
    const searchInput = document.getElementById('clientSearch');
    const select = document.getElementById('clientSelect');

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const select = document.getElementById('clientSelect');

        // Filter and update options
        Array.from(select.options).forEach((option, index) => {
            if (index === 0) return; // Skip default option
            const text = option.textContent.toLowerCase();
            option.style.display = text.includes(query) ? '' : 'none';
        });

        // Auto-select first visible option
        const firstVisible = Array.from(select.options).find(opt => opt.style.display !== 'none');
        if (firstVisible) {
            select.value = firstVisible.value;
        }
    });
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

        if (!response.ok) {
            throw new Error('Failed to load client data');
        }

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

    // Update client card
    document.getElementById('clientName').textContent = data.name;
    document.getElementById('clientUCC').textContent = `UCC: ${data.ucc}`;
    document.getElementById('inceptionDate').textContent = data.inception_date;
    document.getElementById('reportDate').textContent = data.report_date;

    // Update KPI cards
    document.getElementById('costValue').textContent = formatCurrency(data.metrics.cost_value);
    document.getElementById('currentValue').textContent = formatCurrency(data.metrics.current_value);

    const totalPLElement = document.getElementById('totalPL');
    totalPLElement.textContent = formatCurrency(data.metrics.total_pl);
    totalPLElement.className = 'kpi-value ' + (data.metrics.total_pl >= 0 ? 'positive' : 'negative');

    const portfolioXirrElement = document.getElementById('portfolioXIRR');
    portfolioXirrElement.textContent = data.metrics.portfolio_xirr !== null ? formatPercentage(data.metrics.portfolio_xirr) : 'N/A';
    if (data.metrics.portfolio_xirr !== null) {
        portfolioXirrElement.className = 'kpi-value ' + (data.metrics.portfolio_xirr >= 0 ? 'positive' : 'negative');
    }

    const benchmarkXirrElement = document.getElementById('benchmarkXIRR');
    benchmarkXirrElement.textContent = data.metrics.benchmark_xirr !== null ? formatPercentage(data.metrics.benchmark_xirr) : 'N/A';
    if (data.metrics.benchmark_xirr !== null) {
        benchmarkXirrElement.className = 'kpi-value ' + (data.metrics.benchmark_xirr >= 0 ? 'positive' : 'negative');
    }

    // Render categories table
    renderCategoriesTable(data.categories);

    // Render top holdings
    renderTopHoldingsTable(data.top_holdings);

    // Render all holdings
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
            <td class="number ${cat.unrealized_pl >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(cat.unrealized_pl)}
            </td>
            <td class="number ${cat.total_pl >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(cat.total_pl)}
            </td>
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
            <td class="number ${h.total_pl >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(h.total_pl)}
            </td>
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
            <td class="number ${h.total_pl >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(h.total_pl)}
            </td>
        </tr>
    `).join('');
}

// ============================================================
// Download PDF
// ============================================================

async function downloadPDF() {
    if (currentClientId === null) {
        showError('Please select a client first');
        return;
    }

    showLoading('Generating PDF...');

    try {
        const response = await fetch(`/api/client/${currentClientId}/download_pdf`);

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Failed to generate PDF');
        }

        // Get filename from Content-Disposition header
        const contentDisposition = response.headers.get('content-disposition');
        let filename = `factsheet_${currentClientId}.pdf`;
        if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
            if (filenameMatch) filename = filenameMatch[1];
        }

        // Download the PDF
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
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
    // Clear state
    clientsList = [];
    currentClientId = null;
    currentClientData = null;

    // Reset UI
    document.getElementById('uploadSection').style.display = 'block';
    document.getElementById('mainSection').style.display = 'none';
    document.getElementById('fileInfo').textContent = '';
    document.getElementById('fileInput').value = '';
    document.getElementById('clientSearch').value = '';
}

// ============================================================
// Initialize on Page Load
// ============================================================

window.addEventListener('DOMContentLoaded', () => {
    // Page is ready
    console.log('Client Factsheet Viewer loaded');
});
