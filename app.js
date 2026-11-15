/**
 * Main Frontend Application Orchestrator
 * Connects the UI to Audit Engine, SA 520 Engine, Export Service,
 * and optional Colab REST API Backend.
 */

(function () {
  'use strict';

  // Application State
  const state = {
    theme: localStorage.getItem('audit_theme') || 'dark',
    engineMode: localStorage.getItem('audit_engine_mode') || 'client',
    colabUrl: localStorage.getItem('audit_colab_url') || '',
    apAuditData: null,
    sa520Data: null,
    ledgerPage: 1,
    ledgerPageSize: 25,
    ledgerFilter: '',
    uploadTarget: 'ap', // 'ap' or 'screener'
    selectedUploadFile: null,
    charts: {}
  };

  // DOM Elements Cache
  const elements = {
    themeToggleBtn: document.getElementById('themeToggleBtn'),
    engineStatusPill: document.getElementById('engineStatusPill'),
    engineStatusText: document.getElementById('engineStatusText'),
    openSettingsBtn: document.getElementById('openSettingsBtn'),
    navBtns: document.querySelectorAll('.nav-btn'),
    moduleSections: document.querySelectorAll('.module-section'),

    // AP Module
    btnGenerateMockAP: document.getElementById('btnGenerateMockAP'),
    btnUploadAP: document.getElementById('btnUploadAP'),
    cutoffDateInput: document.getElementById('cutoffDateInput'),
    btnExportAPExcel: document.getElementById('btnExportAPExcel'),
    tabBtns: document.querySelectorAll('.tab-btn'),
    subtabContents: document.querySelectorAll('.subtab-content'),

    // AP KPIs
    kpiTotalSpend: document.getElementById('kpiTotalSpend'),
    kpiTotalVouchers: document.getElementById('kpiTotalVouchers'),
    kpiDuplicateExposure: document.getElementById('kpiDuplicateExposure'),
    kpiDuplicateCount: document.getElementById('kpiDuplicateCount'),
    kpiTdsShortfall: document.getElementById('kpiTdsShortfall'),
    kpiTdsCount: document.getElementById('kpiTdsCount'),
    kpiDisallowance30: document.getElementById('kpiDisallowance30'),
    kpiInvalidGstCount: document.getElementById('kpiInvalidGstCount'),
    kpiInvalidGstExposure: document.getElementById('kpiInvalidGstExposure'),
    kpiCutoffCount: document.getElementById('kpiCutoffCount'),
    kpiCutoffExposure: document.getElementById('kpiCutoffExposure'),

    // AP Badges
    badgeVendorCount: document.getElementById('badgeVendorCount'),
    badgeDuplicateCount: document.getElementById('badgeDuplicateCount'),
    badgeTdsCount: document.getElementById('badgeTdsCount'),
    badgeGstCount: document.getElementById('badgeGstCount'),
    badgeCutoffCount: document.getElementById('badgeCutoffCount'),
    badgeFullCount: document.getElementById('badgeFullCount'),

    // AP Tables
    tableVendorSummary: document.getElementById('tableVendorSummary').querySelector('tbody'),
    tableDuplicates: document.getElementById('tableDuplicates').querySelector('tbody'),
    tableTdsShortfalls: document.getElementById('tableTdsShortfalls').querySelector('tbody'),
    tableInvalidGst: document.getElementById('tableInvalidGst').querySelector('tbody'),
    tableCutoff: document.getElementById('tableCutoff').querySelector('tbody'),
    fullLedgerBody: document.getElementById('fullLedgerBody'),
    ledgerSearchInput: document.getElementById('ledgerSearchInput'),
    ledgerPaginationInfo: document.getElementById('ledgerPaginationInfo'),
    btnPrevPage: document.getElementById('btnPrevPage'),
    btnNextPage: document.getElementById('btnNextPage'),

    // SA 520 Module
    btnLoad3YrBaseline: document.getElementById('btnLoad3YrBaseline'),
    btnLoadInfosys10Yr: document.getElementById('btnLoadInfosys10Yr'),
    btnUploadScreener: document.getElementById('btnUploadScreener'),
    btnExportSA520Excel: document.getElementById('btnExportSA520Excel'),
    dupontNetMargin: document.getElementById('dupontNetMargin'),
    dupontAssetTurnover: document.getElementById('dupontAssetTurnover'),
    dupontEquityMultiplier: document.getElementById('dupontEquityMultiplier'),
    dupontCompositeRoe: document.getElementById('dupontCompositeRoe'),
    sa520AlertsGrid: document.getElementById('sa520AlertsGrid'),
    scorecardHead: document.getElementById('scorecardHead'),
    scorecardBody: document.getElementById('scorecardBody'),

    // Colab Bridge Module
    colabUrlInput: document.getElementById('colabUrlInput'),
    btnTestColabPing: document.getElementById('btnTestColabPing'),
    colabPingResult: document.getElementById('colabPingResult'),
    btnCopyColabCode: document.getElementById('btnCopyColabCode'),
    colabCodeBox: document.getElementById('colabCodeBox'),
    engineModeRadios: document.querySelectorAll('input[name="engineMode"]'),

    // Modals
    uploadModal: document.getElementById('uploadModal'),
    uploadModalTitle: document.getElementById('uploadModalTitle'),
    btnCloseUploadModal: document.getElementById('btnCloseUploadModal'),
    btnCancelUpload: document.getElementById('btnCancelUpload'),
    btnConfirmUpload: document.getElementById('btnConfirmUpload'),
    fileDropzone: document.getElementById('fileDropzone'),
    modalFileInput: document.getElementById('modalFileInput'),
    uploadFileInfo: document.getElementById('uploadFileInfo'),
    uploadFileName: document.getElementById('uploadFileName'),
    uploadFileSize: document.getElementById('uploadFileSize'),
    toastContainer: document.getElementById('toastContainer')
  };

  // Utilities
  function formatINR(val) {
    if (val === null || val === undefined || isNaN(val)) return '₹0.00';
    return '₹' + Number(val).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatPct(val) {
    if (val === null || val === undefined || isNaN(val)) return '0.00%';
    return Number(val).toFixed(2) + '%';
  }

  function showToast(msg, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '⚠️';
    toast.innerHTML = `<span>${icon}</span> <span>${msg}</span>`;
    elements.toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 250);
    }, 4000);
  }

  // Theme Management
  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    elements.themeToggleBtn.innerHTML = theme === 'dark' ? '&#9728;' : '&#9790;';
    localStorage.setItem('audit_theme', theme);
  }

  // Navigation Switching
  function switchModule(targetId) {
    elements.navBtns.forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-target') === targetId);
    });
    elements.moduleSections.forEach(sec => {
      sec.classList.toggle('active', sec.id === targetId);
    });
  }

  function switchSubtab(targetId) {
    elements.tabBtns.forEach(btn => {
      btn.classList.toggle('active', btn.getAttribute('data-subtab') === targetId);
    });
    elements.subtabContents.forEach(content => {
      content.style.display = content.id === targetId ? 'block' : 'none';
    });
  }

  // ==============================================================
  // AP LEDGER AUDIT MODULE
  // ==============================================================

  function runAPAudit(rows) {
    const yearEnd = elements.cutoffDateInput.value || '2026-03-31';
    try {
      const results = window.APAuditEngine.analyzeAPLedger(rows, yearEnd);
      state.apAuditData = results;
      renderAPResults(results);
      showToast(`Audit completed! Analyzed ${results.kpis.totalVouchers} vouchers.`, 'success');
    } catch (err) {
      console.error(err);
      showToast(err.message, 'error');
    }
  }

  function renderAPResults(data) {
    const { kpis, exceptions, fullLedger } = data;

    // Render KPIs
    elements.kpiTotalSpend.textContent = formatINR(kpis.totalSpend);
    elements.kpiTotalVouchers.textContent = kpis.totalVouchers;
    elements.kpiDuplicateExposure.textContent = formatINR(kpis.duplicateExposure);
    elements.kpiDuplicateCount.textContent = kpis.duplicateCount;
    elements.kpiTdsShortfall.textContent = formatINR(kpis.totalTdsShortfall);
    elements.kpiTdsCount.textContent = kpis.tdsShortfallCount;
    elements.kpiDisallowance30.textContent = formatINR(kpis.totalDisallowance30);
    elements.kpiInvalidGstCount.textContent = kpis.invalidGstinCount;
    elements.kpiInvalidGstExposure.textContent = formatINR(kpis.invalidGstExposure);
    elements.kpiCutoffCount.textContent = kpis.cutoffErrorCount;
    elements.kpiCutoffExposure.textContent = formatINR(kpis.cutoffExposure);

    // Badges
    elements.badgeVendorCount.textContent = exceptions.vendorSummary.length;
    elements.badgeDuplicateCount.textContent = exceptions.duplicates.length;
    elements.badgeTdsCount.textContent = exceptions.tdsShortfalls.length;
    elements.badgeGstCount.textContent = exceptions.invalidGstins.length;
    elements.badgeCutoffCount.textContent = exceptions.cutoffErrors.length;
    elements.badgeFullCount.textContent = fullLedger.length;

    // Render Sub-Tables
    renderVendorSummaryTable(exceptions.vendorSummary);
    renderDuplicatesTable(exceptions.duplicates);
    renderTdsShortfallsTable(exceptions.tdsShortfalls);
    renderInvalidGstTable(exceptions.invalidGstins);
    renderCutoffTable(exceptions.cutoffErrors);

    state.ledgerPage = 1;
    renderFullLedgerTable();
  }

  function renderVendorSummaryTable(vendors) {
    if (!vendors || vendors.length === 0) {
      elements.tableVendorSummary.innerHTML = `<tr><td colspan="8" class="text-center" style="padding: 2rem; color: var(--text-muted);">No vendor records found.</td></tr>`;
      return;
    }

    elements.tableVendorSummary.innerHTML = vendors.map(v => {
      let riskBadge = `<span class="badge badge-success">Compliant</span>`;
      if (v.risk_score >= 60) {
        riskBadge = `<span class="badge badge-danger">High Risk (${v.risk_score})</span>`;
      } else if (v.risk_score > 0) {
        riskBadge = `<span class="badge badge-warning">Moderate Risk (${v.risk_score})</span>`;
      }

      return `
        <tr>
          <td><strong>${v.vendor}</strong></td>
          <td><code>${v.gstin}</code></td>
          <td class="text-center">${v.voucher_count}</td>
          <td class="text-right">${formatINR(v.total_spend)}</td>
          <td class="text-right">${formatINR(v.total_gst_claimed)}</td>
          <td class="text-right" style="color: ${v.total_tds_shortfall > 0 ? 'var(--accent-rose)' : 'inherit'};">
            ${formatINR(v.total_tds_shortfall)}
          </td>
          <td class="text-right" style="color: ${v.total_disallowance_30pct > 0 ? 'var(--accent-rose)' : 'inherit'}; font-weight: 600;">
            ${formatINR(v.total_disallowance_30pct)}
          </td>
          <td class="text-center">${riskBadge}</td>
        </tr>
      `;
    }).join('');
  }

  function renderDuplicatesTable(duplicates) {
    if (!duplicates || duplicates.length === 0) {
      elements.tableDuplicates.innerHTML = `<tr><td colspan="8" class="text-center" style="padding: 2rem; color: var(--text-muted);">Zero duplicate invoices detected across the ledger!</td></tr>`;
      return;
    }

    elements.tableDuplicates.innerHTML = duplicates.map(r => `
      <tr>
        <td><strong>${r.voucher_no}</strong></td>
        <td><code style="color: var(--accent-rose); font-weight: 700;">${r.invoice_no}</code></td>
        <td>${r.vendor}</td>
        <td><code>${r.gstin}</code></td>
        <td>${r.invoice_date || '-'}</td>
        <td>${r.grn_date || '-'}</td>
        <td class="text-right" style="font-weight: 600;">${formatINR(r.amount)}</td>
        <td class="text-center">
          <span class="badge badge-danger">&#9888; Double Payment Risk</span>
        </td>
      </tr>
    `).join('');
  }

  function renderTdsShortfallsTable(shortfalls) {
    if (!shortfalls || shortfalls.length === 0) {
      elements.tableTdsShortfalls.innerHTML = `<tr><td colspan="9" class="text-center" style="padding: 2rem; color: var(--text-muted);">No statutory TDS shortfalls found. 100% tax compliant!</td></tr>`;
      return;
    }

    elements.tableTdsShortfalls.innerHTML = shortfalls.map(r => `
      <tr>
        <td><strong>${r.voucher_no}</strong></td>
        <td>${r.vendor}</td>
        <td><span class="badge badge-info">${r.section}</span></td>
        <td class="text-right">${formatINR(r.amount)}</td>
        <td class="text-center">${(r.tds_rate * 100).toFixed(0)}%</td>
        <td class="text-right">${formatINR(r.expected_tds)}</td>
        <td class="text-right" style="color: var(--accent-rose);">${formatINR(r.tds_deducted)}</td>
        <td class="text-right" style="color: var(--accent-rose); font-weight: 700;">${formatINR(r.tds_shortfall)}</td>
        <td class="text-right" style="color: var(--accent-rose); font-weight: 700;">${formatINR(r.disallowance_30pct)}</td>
      </tr>
    `).join('');
  }

  function renderInvalidGstTable(invalidGstins) {
    if (!invalidGstins || invalidGstins.length === 0) {
      elements.tableInvalidGst.innerHTML = `<tr><td colspan="7" class="text-center" style="padding: 2rem; color: var(--text-muted);">All reported vendor GSTINs passed the 15-character statutory checksum test!</td></tr>`;
      return;
    }

    elements.tableInvalidGst.innerHTML = invalidGstins.map(r => {
      let diagnostic = "Corrupted GSTIN / State Code Structure";
      if (r.gstin === "MISSING" || !r.gstin) {
        diagnostic = "Missing GSTIN / Unregistered Vendor";
      } else if (r.gstin.length === 14) {
        diagnostic = "Incomplete 14 characters (Checksum digit missing)";
      } else if (r.gstin.includes("INVALID")) {
        diagnostic = "Invalid PAN structure / Dummy GSTIN";
      }

      return `
        <tr>
          <td><strong>${r.voucher_no}</strong></td>
          <td>${r.vendor}</td>
          <td><code style="color: var(--accent-rose); font-weight: 700;">${r.gstin}</code></td>
          <td class="text-right">${formatINR(r.amount)}</td>
          <td class="text-right" style="color: var(--accent-rose);">${formatINR(r.gst_claimed)}</td>
          <td>${diagnostic}</td>
          <td class="text-center">
            <span class="badge badge-danger">Disallow ITC</span>
          </td>
        </tr>
      `;
    }).join('');
  }

  function renderCutoffTable(cutoffErrors) {
    if (!cutoffErrors || cutoffErrors.length === 0) {
      elements.tableCutoff.innerHTML = `<tr><td colspan="7" class="text-center" style="padding: 2rem; color: var(--text-muted);">No year-end cut-off errors detected! All invoices match delivery periods.</td></tr>`;
      return;
    }

    elements.tableCutoff.innerHTML = cutoffErrors.map(r => {
      const isEarly = r.cutoff_status.includes("Early");
      return `
        <tr>
          <td><strong>${r.voucher_no}</strong></td>
          <td>${r.vendor}</td>
          <td>${r.invoice_date}</td>
          <td>${r.grn_date}</td>
          <td class="text-right" style="font-weight: 600;">${formatINR(r.amount)}</td>
          <td>${r.cutoff_status}</td>
          <td class="text-center">
            <span class="badge ${isEarly ? 'badge-danger' : 'badge-warning'}">
              ${isEarly ? 'Expense Overstated' : 'Unrecorded Liability'}
            </span>
          </td>
        </tr>
      `;
    }).join('');
  }

  function renderFullLedgerTable() {
    if (!state.apAuditData || !state.apAuditData.fullLedger) return;

    let filtered = state.apAuditData.fullLedger;
    if (state.ledgerFilter) {
      const q = state.ledgerFilter.toLowerCase();
      filtered = filtered.filter(r =>
        r.voucher_no.toLowerCase().includes(q) ||
        r.vendor.toLowerCase().includes(q) ||
        r.invoice_no.toLowerCase().includes(q) ||
        r.gstin.toLowerCase().includes(q)
      );
    }

    const total = filtered.length;
    const totalPages = Math.ceil(total / state.ledgerPageSize) || 1;
    if (state.ledgerPage > totalPages) state.ledgerPage = totalPages;

    const startIdx = (state.ledgerPage - 1) * state.ledgerPageSize;
    const pageRows = filtered.slice(startIdx, startIdx + state.ledgerPageSize);

    elements.ledgerPaginationInfo.textContent = `Showing ${Math.min(startIdx + 1, total)} to ${Math.min(startIdx + pageRows.length, total)} of ${total} vouchers (Page ${state.ledgerPage} of ${totalPages})`;

    elements.btnPrevPage.disabled = state.ledgerPage <= 1;
    elements.btnNextPage.disabled = state.ledgerPage >= totalPages;

    if (pageRows.length === 0) {
      elements.fullLedgerBody.innerHTML = `<tr><td colspan="13" class="text-center" style="padding: 2rem; color: var(--text-muted);">No matching vouchers found.</td></tr>`;
      return;
    }

    elements.fullLedgerBody.innerHTML = pageRows.map(r => {
      const hasAnomaly = r.is_duplicate || r.tds_shortfall > 0 || !r.is_valid_gstin || r.is_cutoff_error;
      return `
        <tr style="${hasAnomaly ? 'background: rgba(244, 63, 94, 0.04);' : ''}">
          <td><strong>${r.voucher_no}</strong></td>
          <td>${r.vendor}</td>
          <td><code>${r.gstin}</code></td>
          <td>${r.invoice_no}</td>
          <td>${r.invoice_date}</td>
          <td>${r.grn_date}</td>
          <td>${r.section}</td>
          <td class="text-right">${formatINR(r.amount)}</td>
          <td class="text-right">${formatINR(r.gst_claimed)}</td>
          <td class="text-right">${formatINR(r.tds_deducted)}</td>
          <td class="text-right">${formatINR(r.expected_tds)}</td>
          <td class="text-center">
            <span class="badge ${r.is_valid_gstin ? 'badge-success' : 'badge-danger'}">
              ${r.is_valid_gstin ? 'Valid' : 'Invalid'}
            </span>
          </td>
          <td class="text-center">
            <span class="badge ${r.is_cutoff_error ? 'badge-danger' : 'badge-success'}">
              ${r.is_cutoff_error ? 'Error' : 'Compliant'}
            </span>
          </td>
        </tr>
      `;
    }).join('');
  }

  // ==============================================================
  // SA 520 ANALYTICAL PROCEDURES & DUPONT ROE MODULE
  // ==============================================================

  function runSA520Analysis(records) {
    try {
      const results = window.SA520Engine.analyzeFinancialData(records);
      state.sa520Data = results;
      renderSA520Results(results);
      showToast(`SA 520 Analytical Procedures run across ${results.years.length} financial years.`, 'success');
    } catch (err) {
      console.error(err);
      showToast(err.message, 'error');
    }
  }

  function renderSA520Results(results) {
    const { years, scorecard } = results;
    const latestYear = years[years.length - 1];

    // 1. Render DuPont Breakdown Cards (for the latest year)
    elements.dupontNetMargin.textContent = formatPct(latestYear.ratios.netMarginPct);
    elements.dupontAssetTurnover.textContent = (latestYear.ratios.assetTurnover || 0).toFixed(2) + 'x';
    elements.dupontEquityMultiplier.textContent = (latestYear.ratios.equityMultiplier || 0).toFixed(2) + 'x';
    elements.dupontCompositeRoe.textContent = formatPct(latestYear.ratios.dupontRoePct);

    // 2. Render Scorecard Table (Transposed: Ratios as rows, Years as columns)
    const yearHeaders = years.map(y => `<th class="text-right">${y.financial_year}</th>`).join('');
    elements.scorecardHead.innerHTML = `
      <tr>
        <th>Audit & Financial Ratio</th>
        <th class="text-center">Unit</th>
        ${yearHeaders}
        <th class="text-right">YoY Variance</th>
      </tr>
    `;

    elements.scorecardBody.innerHTML = scorecard.map(row => {
      const yearCells = years.map(y => {
        const val = row[y.financial_year];
        let display = '-';
        if (val !== null && val !== undefined) {
          display = typeof val === 'number' ? val.toFixed(2) : val;
        }
        return `<td class="text-right">${display}</td>`;
      }).join('');

      let yoyDisplay = '-';
      let yoyStyle = '';
      if (row.yoy_change !== null && row.yoy_change !== undefined) {
        yoyDisplay = (row.yoy_change > 0 ? '+' : '') + row.yoy_change.toFixed(2);
        if (row.ratio.includes('Margin') || row.ratio.includes('ROE')) {
          yoyStyle = row.yoy_change < 0 ? 'color: var(--accent-rose); font-weight: 700;' : 'color: var(--accent-emerald); font-weight: 700;';
        }
      }

      return `
        <tr>
          <td><strong>${row.ratio}</strong></td>
          <td class="text-center"><span class="badge badge-info">${row.unit}</span></td>
          ${yearCells}
          <td class="text-right" style="${yoyStyle}">${yoyDisplay}</td>
        </tr>
      `;
    }).join('');

    // 3. Render SA 520 Alerts Grid
    renderSA520AlertsGrid(years);

    // 4. Render Analytical Charts
    renderAnalyticalCharts(years);
  }

  function renderSA520AlertsGrid(years) {
    const allAlerts = [];
    years.forEach(y => {
      y.flags.forEach(f => {
        allAlerts.push({ year: y.financial_year, ...f });
      });
    });

    if (allAlerts.length === 0) {
      elements.sa520AlertsGrid.innerHTML = `
        <div class="alert-item success" style="grid-column: 1 / -1;">
          <div class="alert-icon">✅</div>
          <div class="alert-content">
            <h4>Clean Financial Health Profile</h4>
            <p>All core audit ratios (Current Ratio, Debt-to-Equity, Debtor Days, Interest Cover) are compliant with statutory and banking norms.</p>
          </div>
        </div>
      `;
      return;
    }

    elements.sa520AlertsGrid.innerHTML = allAlerts.map(a => `
      <div class="alert-item ${a.severity === 'danger' || a.severity === 'high' ? 'danger' : 'warning'}">
        <div class="alert-icon">${a.severity === 'danger' || a.severity === 'high' ? '🚨' : '⚠️'}</div>
        <div class="alert-content">
          <h4>${a.title} &bull; ${a.year}</h4>
          <p>${a.detail}</p>
        </div>
      </div>
    `).join('');
  }

  function renderAnalyticalCharts(years) {
    if (typeof Chart === 'undefined') return;

    const labels = years.map(y => y.financial_year);
    const roeValues = years.map(y => y.ratios.dupontRoePct);
    const netMargins = years.map(y => y.ratios.netMarginPct);
    const ebitMargins = years.map(y => y.ratios.ebitMarginPct);
    const recDays = years.map(y => y.ratios.receivablesDays);
    const invDays = years.map(y => y.ratios.inventoryDays);
    const currentRatios = years.map(y => y.ratios.currentRatio);
    const debtToEquities = years.map(y => y.ratios.debtToEquity);

    // Destroy existing charts to prevent canvas memory leaks
    Object.values(state.charts).forEach(c => c && c.destroy());

    const chartOptions = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: state.theme === 'dark' ? '#94a3b8' : '#475569', font: { size: 11 } } }
      },
      scales: {
        x: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#94a3b8' } },
        y: { grid: { color: 'rgba(255, 255, 255, 0.05)' }, ticks: { color: '#94a3b8' } }
      }
    };

    // Chart 1: DuPont ROE Trend
    const ctxDupont = document.getElementById('chartDupont').getContext('2d');
    state.charts.dupont = new Chart(ctxDupont, {
      type: 'line',
      data: {
        labels,
        datasets: [{
          label: 'DuPont Composite ROE %',
          data: roeValues,
          borderColor: '#06b6d4',
          backgroundColor: 'rgba(6, 182, 212, 0.15)',
          fill: true,
          tension: 0.3,
          borderWidth: 3
        }]
      },
      options: chartOptions
    });

    // Chart 2: Margins Trend
    const ctxMargins = document.getElementById('chartMargins').getContext('2d');
    state.charts.margins = new Chart(ctxMargins, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'EBIT Margin %', data: ebitMargins, borderColor: '#3b82f6', tension: 0.3, borderWidth: 2 },
          { label: 'Net Margin %', data: netMargins, borderColor: '#10b981', tension: 0.3, borderWidth: 2 }
        ]
      },
      options: chartOptions
    });

    // Chart 3: Working Capital Cycle
    const ctxWc = document.getElementById('chartWorkingCapital').getContext('2d');
    state.charts.wc = new Chart(ctxWc, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Receivables Days (DSO)', data: recDays, backgroundColor: '#f59e0b', borderRadius: 4 },
          { label: 'Inventory Days (DIO)', data: invDays, backgroundColor: '#8b5cf6', borderRadius: 4 }
        ]
      },
      options: chartOptions
    });

    // Chart 4: Solvency & Liquidity
    const ctxSolvency = document.getElementById('chartSolvency').getContext('2d');
    state.charts.solvency = new Chart(ctxSolvency, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Current Ratio (x)', data: currentRatios, borderColor: '#10b981', borderWidth: 2 },
          { label: 'Debt to Equity (x)', data: debtToEquities, borderColor: '#ef4444', borderWidth: 2 }
        ]
      },
      options: chartOptions
    });
  }

  // ==============================================================
  // MODAL & FILE UPLOAD MANAGEMENT
  // ==============================================================

  function openUploadModal(target) {
    state.uploadTarget = target;
    state.selectedUploadFile = null;
    elements.modalFileInput.value = '';
    elements.uploadFileInfo.style.display = 'none';
    elements.btnConfirmUpload.disabled = true;

    if (target === 'ap') {
      elements.uploadModalTitle.textContent = 'Upload ERP Accounts Payable Ledger';
    } else {
      elements.uploadModalTitle.textContent = 'Upload Screener.in Financial Excel (.xlsx)';
    }

    elements.uploadModal.classList.add('open');
  }

  function closeUploadModal() {
    elements.uploadModal.classList.remove('open');
  }

  function handleFileSelected(file) {
    if (!file) return;
    state.selectedUploadFile = file;
    elements.uploadFileName.textContent = file.name;
    elements.uploadFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;
    elements.uploadFileInfo.style.display = 'block';
    elements.btnConfirmUpload.disabled = false;
  }

  function processUploadedFile() {
    const file = state.selectedUploadFile;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (state.uploadTarget === 'ap') {
          // Parse first sheet as AP ledger rows
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(firstSheet);
          runAPAudit(rows);
        } else {
          // Parse Screener workbook
          const results = window.SA520Engine.parseScreenerWorkbook(workbook);
          state.sa520Data = results;
          renderSA520Results(results);
          showToast(`Extracted Screener financial blocks successfully.`, 'success');
        }
        closeUploadModal();
      } catch (err) {
        console.error(err);
        showToast("Error processing file: " + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ==============================================================
  // COLAB BACKEND PING & CONNECTOR
  // ==============================================================

  async function testColabBackendPing() {
    const url = elements.colabUrlInput.value.trim().replace(/\/+$/, '');
    if (!url) {
      showToast('Please enter a valid Colab URL or ngrok endpoint.', 'error');
      return;
    }

    elements.colabPingResult.textContent = 'Testing connection...';
    elements.colabPingResult.style.color = 'var(--text-muted)';

    try {
      const startTime = performance.now();
      const res = await fetch(`${url}/`, { method: 'GET' });
      const latency = Math.round(performance.now() - startTime);

      if (res.ok) {
        const json = await res.json();
        elements.colabPingResult.textContent = `Connected! Latency: ${latency}ms - ${json.message || json.service || 'Online'}`;
        elements.colabPingResult.style.color = 'var(--accent-emerald)';
        showToast('Successfully connected to Colab backend!', 'success');
      } else {
        throw new Error(`HTTP ${res.status}`);
      }
    } catch (err) {
      elements.colabPingResult.textContent = `Connection failed: ${err.message}. Ensure ngrok is running and accepts requests.`;
      elements.colabPingResult.style.color = 'var(--accent-rose)';
      showToast('Could not reach Colab endpoint. Check URL or firewall.', 'error');
    }
  }

  // ==============================================================
  // EVENT LISTENERS INITIALIZATION
  // ==============================================================

  function initEventListeners() {
    // Theme Toggle
    elements.themeToggleBtn.addEventListener('click', () => {
      applyTheme(state.theme === 'dark' ? 'light' : 'dark');
      if (state.sa520Data) {
        renderAnalyticalCharts(state.sa520Data.years);
      }
    });

    // Navigation Switcher
    elements.navBtns.forEach(btn => {
      btn.addEventListener('click', () => switchModule(btn.getAttribute('data-target')));
    });

    // Settings Button shortcuts to Colab tab
    elements.openSettingsBtn.addEventListener('click', () => switchModule('module-colab'));
    elements.engineStatusPill.addEventListener('click', () => switchModule('module-colab'));

    // Subtab Switcher
    elements.tabBtns.forEach(btn => {
      btn.addEventListener('click', () => switchSubtab(btn.getAttribute('data-subtab')));
    });

    // AP Module Buttons
    elements.btnGenerateMockAP.addEventListener('click', () => {
      const mockRows = window.APAuditEngine.generateMockLedger(500, 42);
      runAPAudit(mockRows);
    });

    elements.btnUploadAP.addEventListener('click', () => openUploadModal('ap'));

    elements.cutoffDateInput.addEventListener('change', () => {
      if (state.apAuditData && state.apAuditData.fullLedger) {
        runAPAudit(state.apAuditData.fullLedger);
      }
    });

    elements.btnExportAPExcel.addEventListener('click', () => {
      if (!state.apAuditData) {
        showToast('Please load or generate an AP ledger before exporting.', 'error');
        return;
      }
      window.ExportService.exportMasterAPAudit(state.apAuditData);
      showToast('Master 6-Tab AP Audit Workbook downloaded!', 'success');
    });

    // Ledger Search & Pagination
    elements.ledgerSearchInput.addEventListener('input', (e) => {
      state.ledgerFilter = e.target.value.trim();
      state.ledgerPage = 1;
      renderFullLedgerTable();
    });

    elements.btnPrevPage.addEventListener('click', () => {
      if (state.ledgerPage > 1) {
        state.ledgerPage--;
        renderFullLedgerTable();
      }
    });

    elements.btnNextPage.addEventListener('click', () => {
      state.ledgerPage++;
      renderFullLedgerTable();
    });

    // SA 520 Module Buttons
    elements.btnLoad3YrBaseline.addEventListener('click', () => {
      runSA520Analysis(window.SA520Engine.baseline3Year);
    });

    elements.btnLoadInfosys10Yr.addEventListener('click', () => {
      // Direct load of Infosys 10-year model
      const records = window.SA520Engine.infosys10Year.map(item => ({
        financial_year: item.year,
        revenue: 100000.0,
        net_profit: 100000.0 * (item.netMargin / 100),
        ebit: 100000.0 * (item.ebitMargin / 100),
        shareholders_equity: 100000.0 * (item.netMargin / 100) / (item.dupontRoe / 100),
        total_assets: 100000.0 / item.assetTurnover,
        total_debt: (100000.0 * (item.netMargin / 100) / (item.dupontRoe / 100)) * item.debtToEquity,
        current_liabilities: 20000.0,
        trade_receivables: (item.receivablesDays * 100000.0) / 365,
        inventory: (item.inventoryDays * 100000.0) / 365,
        cash_bank: 15000.0
      }));
      runSA520Analysis(records);
    });

    elements.btnUploadScreener.addEventListener('click', () => openUploadModal('screener'));

    elements.btnExportSA520Excel.addEventListener('click', () => {
      if (!state.sa520Data) {
        showToast('Please load financial data before exporting SA 520 scorecard.', 'error');
        return;
      }
      window.ExportService.exportSA520Audit(state.sa520Data);
      showToast('SA 520 Analytical Scorecard downloaded!', 'success');
    });

    // Colab Bridge Listeners
    elements.btnTestColabPing.addEventListener('click', testColabBackendPing);

    elements.btnCopyColabCode.addEventListener('click', () => {
      navigator.clipboard.writeText(elements.colabCodeBox.textContent);
      showToast('Colab server code copied to clipboard!', 'success');
    });

    elements.engineModeRadios.forEach(radio => {
      radio.addEventListener('change', (e) => {
        state.engineMode = e.target.value;
        localStorage.setItem('audit_engine_mode', state.engineMode);
        elements.engineStatusText.textContent = state.engineMode === 'client' ? 'Browser Engine (Instant)' : 'Colab API (Active)';
        showToast(`Switched execution mode to ${state.engineMode === 'client' ? 'Browser Engine' : 'Colab Backend API'}.`, 'info');
      });
    });

    // Modal Upload Listeners
    elements.btnCloseUploadModal.addEventListener('click', closeUploadModal);
    elements.btnCancelUpload.addEventListener('click', closeUploadModal);
    elements.fileDropzone.addEventListener('click', () => elements.modalFileInput.click());

    elements.modalFileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleFileSelected(e.target.files[0]);
      }
    });

    elements.fileDropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      elements.fileDropzone.classList.add('dragover');
    });

    elements.fileDropzone.addEventListener('dragleave', () => {
      elements.fileDropzone.classList.remove('dragover');
    });

    elements.fileDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      elements.fileDropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleFileSelected(e.dataTransfer.files[0]);
      }
    });

    elements.btnConfirmUpload.addEventListener('click', processUploadedFile);
  }

  // Self-Initialization on page load
  window.addEventListener('DOMContentLoaded', () => {
    applyTheme(state.theme);
    initEventListeners();

    // Auto-load 500-row ERP dataset as default showcase
    const initialMock = window.APAuditEngine.generateMockLedger(500, 42);
    runAPAudit(initialMock);

    // Auto-load 3-year baseline model for SA 520 showcase
    runSA520Analysis(window.SA520Engine.baseline3Year);
  });

})();
