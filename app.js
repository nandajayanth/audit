/**
 * Main Frontend Application Orchestrator
 * Connects the UI to Audit Engine, SA 520 Engine, and Export Service.
 */

(function () {
  'use strict';

  // Application State
  const state = {
    theme: localStorage.getItem('audit_theme') || 'dark',
    apAuditData: null,
    cutoffData: null,
    sa520Data: null,
    ledgerPage: 1,
    ledgerPageSize: 25,
    ledgerFilter: '',
    chartDupont: null,
    rawCutoffRows: null
  };

  // DOM Elements Cache
  const elements = {
    themeToggleBtn: document.getElementById('themeToggleBtn'),
    navBtns: document.querySelectorAll('.nav-btn'),
    moduleSections: document.querySelectorAll('.module-section'),
    tabBtns: document.querySelectorAll('.tab-btn'),
    toastContainer: document.getElementById('toastContainer'),

    // Module 1: AP Audit
    btnTestAPSample: document.getElementById('btnTestAPSample'),
    btnUploadAP: document.getElementById('btnUploadAP'),
    inputAPFile: document.getElementById('inputAPFile'),
    btnExportAPExcel: document.getElementById('btnExportAPExcel'),
    kpiTotalSpend: document.getElementById('kpiTotalSpend'),
    kpiTotalVouchers: document.getElementById('kpiTotalVouchers'),
    kpiDuplicateExposure: document.getElementById('kpiDuplicateExposure'),
    kpiDuplicateCount: document.getElementById('kpiDuplicateCount'),
    kpiTdsShortfall: document.getElementById('kpiTdsShortfall'),
    kpiTdsCount: document.getElementById('kpiTdsCount'),
    kpiDisallowance30: document.getElementById('kpiDisallowance30'),
    kpiInvalidGstCount: document.getElementById('kpiInvalidGstCount'),
    kpiInvalidGstExposure: document.getElementById('kpiInvalidGstExposure'),
    kpiCleanCount: document.getElementById('kpiCleanCount'),
    kpiCleanPct: document.getElementById('kpiCleanPct'),
    badgeVendorCount: document.getElementById('badgeVendorCount'),
    badgeDuplicateCount: document.getElementById('badgeDuplicateCount'),
    badgeTdsCount: document.getElementById('badgeTdsCount'),
    badgeGstCount: document.getElementById('badgeGstCount'),
    badgeFullCount: document.getElementById('badgeFullCount'),
    tableVendorSummary: document.getElementById('tableVendorSummary').querySelector('tbody'),
    tableDuplicates: document.getElementById('tableDuplicates').querySelector('tbody'),
    tableTdsShortfalls: document.getElementById('tableTdsShortfalls').querySelector('tbody'),
    tableInvalidGst: document.getElementById('tableInvalidGst').querySelector('tbody'),
    fullLedgerBody: document.getElementById('fullLedgerBody'),
    ledgerSearchInput: document.getElementById('ledgerSearchInput'),
    ledgerPaginationInfo: document.getElementById('ledgerPaginationInfo'),
    btnPrevPage: document.getElementById('btnPrevPage'),
    btnNextPage: document.getElementById('btnNextPage'),

    // Module 2: Cut-Off Testing
    btnTestCutoffSample: document.getElementById('btnTestCutoffSample'),
    btnUploadCutoff: document.getElementById('btnUploadCutoff'),
    inputCutoffFile: document.getElementById('inputCutoffFile'),
    cutoffDatePicker: document.getElementById('cutoffDatePicker'),
    btnExportCutoffExcel: document.getElementById('btnExportCutoffExcel'),
    kpiCutoffTestedCount: document.getElementById('kpiCutoffTestedCount'),
    kpiCutoffTotalErrors: document.getElementById('kpiCutoffTotalErrors'),
    kpiCutoffTotalErrorAmount: document.getElementById('kpiCutoffTotalErrorAmount'),
    kpiCutoffEarlyCount: document.getElementById('kpiCutoffEarlyCount'),
    kpiCutoffEarlyAmount: document.getElementById('kpiCutoffEarlyAmount'),
    kpiCutoffLateCount: document.getElementById('kpiCutoffLateCount'),
    kpiCutoffLateAmount: document.getElementById('kpiCutoffLateAmount'),
    badgeCutoffExceptionsCount: document.getElementById('badgeCutoffExceptionsCount'),
    badgeCutoffAllCount: document.getElementById('badgeCutoffAllCount'),
    tableCutoffExceptions: document.getElementById('tableCutoffExceptions').querySelector('tbody'),
    tableCutoffAll: document.getElementById('tableCutoffAll').querySelector('tbody'),

    // Module 3: SA 520 & DuPont ROE
    btnTestInfosysSample: document.getElementById('btnTestInfosysSample'),
    btnLoad3YrBaseline: document.getElementById('btnLoad3YrBaseline'),
    btnUploadScreener: document.getElementById('btnUploadScreener'),
    inputScreenerFile: document.getElementById('inputScreenerFile'),
    btnExportSA520Excel: document.getElementById('btnExportSA520Excel'),
    dupontNetMargin: document.getElementById('dupontNetMargin'),
    dupontAssetTurnover: document.getElementById('dupontAssetTurnover'),
    dupontEquityMultiplier: document.getElementById('dupontEquityMultiplier'),
    dupontCompositeRoe: document.getElementById('dupontCompositeRoe'),
    kpiLatestRoe: document.getElementById('kpiLatestRoe'),
    kpiLatestNetMargin: document.getElementById('kpiLatestNetMargin'),
    kpiLatestAssetTurnover: document.getElementById('kpiLatestAssetTurnover'),
    kpiLatestEquityMultiplier: document.getElementById('kpiLatestEquityMultiplier'),
    kpiLatestDebtToEquity: document.getElementById('kpiLatestDebtToEquity'),
    badgeSa520AlertCount: document.getElementById('badgeSa520AlertCount'),
    chartCanvas: document.getElementById('chartDupontTrend'),
    scorecardHead: document.getElementById('scorecardHead'),
    scorecardBody: document.getElementById('scorecardBody'),
    sa520AlertsBody: document.getElementById('sa520AlertsBody'),
    sa520FullHead: document.getElementById('sa520FullHead'),
    sa520FullBody: document.getElementById('sa520FullBody')
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
    if (!elements.toastContainer) return;
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
    if (elements.themeToggleBtn) {
      elements.themeToggleBtn.innerHTML = theme === 'dark' ? '&#9728;' : '&#9790;';
    }
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

  function setupSubtabs() {
    elements.tabBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const parentModule = btn.closest('.module-section');
        const targetSubtab = btn.getAttribute('data-subtab');
        if (!parentModule || !targetSubtab) return;

        parentModule.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        parentModule.querySelectorAll('.subtab-content').forEach(c => c.style.display = 'none');

        btn.classList.add('active');
        const contentEl = document.getElementById(targetSubtab);
        if (contentEl) contentEl.style.display = 'block';
      });
    });
  }

  // =========================================================================
  // MODULE 1: AP FORENSIC AUDIT
  // =========================================================================
  function renderAPAuditResults(results) {
    if (!results) return;
    state.apAuditData = results;
    const summary = results.summary || results.kpis || {};
    const exceptions = results.exceptions || {};
    const fullLedger = results.fullLedger || [];

    // Update KPI cards
    elements.kpiTotalSpend.textContent = formatINR(summary.totalSpend || 0);
    elements.kpiTotalVouchers.textContent = (summary.totalVouchers || 0).toLocaleString();
    elements.kpiDuplicateExposure.textContent = formatINR(summary.duplicateExposure || 0);
    elements.kpiDuplicateCount.textContent = summary.duplicateCount || 0;
    elements.kpiTdsShortfall.textContent = formatINR(summary.totalTdsShortfall || 0);
    elements.kpiTdsCount.textContent = summary.tdsShortfallCount || 0;
    elements.kpiDisallowance30.textContent = formatINR(summary.totalDisallowance30 || 0);
    elements.kpiInvalidGstCount.textContent = summary.invalidGstinCount || 0;
    elements.kpiInvalidGstExposure.textContent = Number(summary.invalidGstExposure || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 });

    const cleanCount = fullLedger.filter(r => !r.is_duplicate && r.tds_shortfall === 0 && r.is_valid_gstin).length;
    elements.kpiCleanCount.textContent = cleanCount.toLocaleString();
    const cleanPct = (summary.totalVouchers || 0) > 0 ? ((cleanCount / summary.totalVouchers) * 100).toFixed(1) : 0;
    elements.kpiCleanPct.textContent = `${cleanPct}%`;

    // Badges
    const vendorSummary = exceptions.vendorSummary || [];
    const duplicates = exceptions.duplicates || [];
    const tdsShortfalls = exceptions.tdsShortfalls || [];
    const invalidGstins = exceptions.invalidGstins || [];

    elements.badgeVendorCount.textContent = vendorSummary.length;
    elements.badgeDuplicateCount.textContent = duplicates.length;
    elements.badgeTdsCount.textContent = tdsShortfalls.length;
    elements.badgeGstCount.textContent = invalidGstins.length;
    elements.badgeFullCount.textContent = fullLedger.length;

    // Render Sub-Tab 1: Vendor Summary
    if (exceptions.vendorSummary.length === 0) {
      elements.tableVendorSummary.innerHTML = '<tr><td colspan="8" class="text-center" style="padding: 2rem; color: var(--text-muted);">No vendor records found.</td></tr>';
    } else {
      elements.tableVendorSummary.innerHTML = exceptions.vendorSummary.map(v => {
        let badgeHtml = '<span class="badge badge-success">Compliant</span>';
        if (v.has_missing_name) badgeHtml = '<span class="badge badge-danger">Missing Name</span>';
        else if (v.has_invalid_gst) badgeHtml = '<span class="badge badge-danger">Broken GSTIN</span>';
        else if (v.total_tds_shortfall > 0) badgeHtml = '<span class="badge badge-warning">TDS Non-Compliant</span>';

        return `
          <tr>
            <td><strong>${v.vendor}</strong></td>
            <td><code>${v.gstin}</code></td>
            <td class="text-center">${v.voucher_count}</td>
            <td class="text-right">${formatINR(v.total_spend)}</td>
            <td class="text-right">${formatINR(v.total_gst_claimed)}</td>
            <td class="text-right" style="color: ${v.total_tds_shortfall > 0 ? 'var(--accent-rose)' : 'inherit'}; font-weight: ${v.total_tds_shortfall > 0 ? '700' : 'normal'};">${formatINR(v.total_tds_shortfall)}</td>
            <td class="text-right" style="color: ${v.total_disallowance_30pct > 0 ? 'var(--accent-purple)' : 'inherit'}; font-weight: ${v.total_disallowance_30pct > 0 ? '700' : 'normal'};">${formatINR(v.total_disallowance_30pct)}</td>
            <td class="text-center">${badgeHtml}</td>
          </tr>
        `;
      }).join('');
    }

    // Render Sub-Tab 2: Duplicate Invoices
    if (exceptions.duplicates.length === 0) {
      elements.tableDuplicates.innerHTML = '<tr><td colspan="8" class="text-center" style="padding: 2rem; color: var(--text-muted);">No duplicate invoices detected. Zero double-payment risk.</td></tr>';
    } else {
      elements.tableDuplicates.innerHTML = exceptions.duplicates.map(r => `
        <tr>
          <td><code>${r.voucher_no}</code></td>
          <td><strong style="color: var(--accent-rose);">${r.invoice_no}</strong></td>
          <td>${r.vendor}</td>
          <td><code>${r.gstin}</code></td>
          <td>${r.invoice_date}</td>
          <td>${r.grn_date}</td>
          <td class="text-right font-mono" style="font-weight: 700;">${formatINR(r.amount)}</td>
          <td class="text-center"><span class="badge badge-danger">Double Payment Risk</span></td>
        </tr>
      `).join('');
    }

    // Render Sub-Tab 3: TDS Shortfalls
    if (exceptions.tdsShortfalls.length === 0) {
      elements.tableTdsShortfalls.innerHTML = '<tr><td colspan="9" class="text-center" style="padding: 2rem; color: var(--text-muted);">No TDS shortfalls found. 100% compliant with Sections 194C and 194J.</td></tr>';
    } else {
      elements.tableTdsShortfalls.innerHTML = exceptions.tdsShortfalls.map(r => `
        <tr>
          <td><code>${r.voucher_no}</code></td>
          <td>${r.vendor}</td>
          <td><span class="badge badge-info">Sec ${r.section}</span></td>
          <td class="text-right font-mono">${formatINR(r.amount)}</td>
          <td class="text-center">${(r.tds_rate * 100).toFixed(0)}%</td>
          <td class="text-right font-mono">${formatINR(r.expected_tds)}</td>
          <td class="text-right font-mono" style="color: var(--accent-rose); font-weight: 600;">${formatINR(r.tds_deducted)}</td>
          <td class="text-right font-mono" style="color: var(--accent-rose); font-weight: 700;">${formatINR(r.tds_shortfall)}</td>
          <td class="text-right font-mono" style="color: var(--accent-purple); font-weight: 700;">${formatINR(r.disallowance_30pct)}</td>
        </tr>
      `).join('');
    }

    // Render Sub-Tab 4: Invalid GSTINs
    if (exceptions.invalidGstins.length === 0) {
      elements.tableInvalidGst.innerHTML = '<tr><td colspan="7" class="text-center" style="padding: 2rem; color: var(--text-muted);">All vendor GSTINs compliant with 15-character statutory regex.</td></tr>';
    } else {
      elements.tableInvalidGst.innerHTML = exceptions.invalidGstins.map(r => {
        let reason = "Invalid PAN or Checksum Pattern";
        if (r.gstin === "MISSING" || !r.gstin) reason = "GSTIN Missing in Voucher";
        else if (r.gstin.length !== 15) reason = `Invalid Length (${r.gstin.length} chars, expected 15)`;

        return `
          <tr>
            <td><code>${r.voucher_no}</code></td>
            <td>${r.vendor}</td>
            <td><code style="color: var(--accent-rose);">${r.gstin}</code></td>
            <td class="text-right font-mono">${formatINR(r.amount)}</td>
            <td class="text-right font-mono" style="color: var(--accent-rose); font-weight: 700;">${formatINR(r.gst_claimed)}</td>
            <td><span class="highlight-danger">${reason}</span></td>
            <td class="text-center"><span class="badge badge-danger">Ineligible ITC</span></td>
          </tr>
        `;
      }).join('');
    }

    // Render Sub-Tab 5: Full Ledger
    renderFullLedgerTable();
  }

  function renderFullLedgerTable() {
    if (!state.apAuditData) return;
    const all = state.apAuditData.fullLedger;
    const filter = (state.ledgerFilter || '').toLowerCase().trim();

    const filtered = filter ? all.filter(r => {
      return (r.voucher_no && r.voucher_no.toLowerCase().includes(filter)) ||
             (r.vendor && r.vendor.toLowerCase().includes(filter)) ||
             (r.invoice_no && r.invoice_no.toLowerCase().includes(filter)) ||
             (r.gstin && r.gstin.toLowerCase().includes(filter));
    }) : all;

    const totalRows = filtered.length;
    const totalPages = Math.ceil(totalRows / state.ledgerPageSize) || 1;
    if (state.ledgerPage > totalPages) state.ledgerPage = totalPages;
    if (state.ledgerPage < 1) state.ledgerPage = 1;

    const startIdx = (state.ledgerPage - 1) * state.ledgerPageSize;
    const endIdx = Math.min(startIdx + state.ledgerPageSize, totalRows);
    const pageRows = filtered.slice(startIdx, endIdx);

    elements.ledgerPaginationInfo.textContent = `Showing ${totalRows === 0 ? 0 : startIdx + 1} - ${endIdx} of ${totalRows} vouchers (Page ${state.ledgerPage} of ${totalPages})`;
    elements.btnPrevPage.disabled = state.ledgerPage <= 1;
    elements.btnNextPage.disabled = state.ledgerPage >= totalPages;

    if (pageRows.length === 0) {
      elements.fullLedgerBody.innerHTML = '<tr><td colspan="12" class="text-center" style="padding: 2rem; color: var(--text-muted);">No vouchers matching filter.</td></tr>';
      return;
    }

    elements.fullLedgerBody.innerHTML = pageRows.map(r => `
      <tr>
        <td><code>${r.voucher_no}</code></td>
        <td>${r.vendor}</td>
        <td><code>${r.gstin}</code></td>
        <td>${r.invoice_no}</td>
        <td>${r.invoice_date}</td>
        <td>${r.grn_date}</td>
        <td><span class="badge badge-info">${r.section}</span></td>
        <td class="text-right font-mono">${formatINR(r.amount)}</td>
        <td class="text-right font-mono">${formatINR(r.gst_claimed)}</td>
        <td class="text-right font-mono">${formatINR(r.tds_deducted)}</td>
        <td class="text-right font-mono">${formatINR(r.expected_tds)}</td>
        <td class="text-center">${r.is_valid_gstin ? '<span class="badge badge-success">Valid</span>' : '<span class="badge badge-danger">Invalid</span>'}</td>
      </tr>
    `).join('');
  }

  // Load AP Mock Data
  function loadAPMockData() {
    try {
      const mockRows = APAuditEngine.generateMockLedger(500, 42);
      const results = APAuditEngine.analyzeAPLedger(mockRows, "2026-03-31");
      renderAPAuditResults(results);
      showToast("500-Row Sample AP Ledger analyzed successfully!", "success");
    } catch (e) {
      showToast(`Error generating AP sample: ${e.message}`, "error");
    }
  }

  // =========================================================================
  // MODULE 2: YEAR-END CUT-OFF TESTING
  // =========================================================================
  function renderCutoffResults(cutoffRes) {
    if (!cutoffRes) return;
    state.cutoffData = cutoffRes;
    const summary = cutoffRes.summary || cutoffRes.kpis || {};
    const exceptions = cutoffRes.exceptions || [];
    const allVouchers = cutoffRes.allVouchers || [];

    elements.kpiCutoffTestedCount.textContent = (summary.testedCount || 0).toLocaleString();
    elements.kpiCutoffTotalErrors.textContent = (summary.totalErrors || 0).toLocaleString();
    elements.kpiCutoffTotalErrorAmount.textContent = Number(summary.totalErrorAmt || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 });
    elements.kpiCutoffEarlyCount.textContent = (summary.earlyCount || 0).toLocaleString();
    elements.kpiCutoffEarlyAmount.textContent = Number(summary.earlyAmt || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 });
    elements.kpiCutoffLateCount.textContent = (summary.lateCount || 0).toLocaleString();
    elements.kpiCutoffLateAmount.textContent = Number(summary.lateAmt || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 });

    elements.badgeCutoffExceptionsCount.textContent = exceptions.length;
    elements.badgeCutoffAllCount.textContent = allVouchers.length;

    // Render Exceptions Table
    if (exceptions.length === 0) {
      elements.tableCutoffExceptions.innerHTML = '<tr><td colspan="8" class="text-center" style="padding: 2rem; color: var(--text-muted);">No cut-off errors detected for year-end boundary.</td></tr>';
    } else {
      elements.tableCutoffExceptions.innerHTML = exceptions.map(r => {
        const isEarly = r.cutoff_status.includes("Early");
        const severityBadge = isEarly
          ? '<span class="badge badge-danger">Expense Overstated in Current FY</span>'
          : '<span class="badge badge-warning">Expense Understated / Unrecorded</span>';

        return `
          <tr>
            <td><code>${r.voucher_no}</code></td>
            <td><strong>${r.vendor}</strong></td>
            <td>${r.invoice_no}</td>
            <td>${r.invoice_date}</td>
            <td>${r.grn_date}</td>
            <td class="text-right font-mono" style="font-weight: 700; color: var(--accent-rose);">${formatINR(r.amount)}</td>
            <td><strong style="color: ${isEarly ? 'var(--accent-rose)' : 'var(--accent-amber)'};">${r.cutoff_status}</strong></td>
            <td class="text-center">${severityBadge}</td>
          </tr>
        `;
      }).join('');
    }

    // Render All Tested Vouchers Table
    if (allVouchers.length === 0) {
      elements.tableCutoffAll.innerHTML = '<tr><td colspan="7" class="text-center" style="padding: 2rem; color: var(--text-muted);">No records loaded.</td></tr>';
    } else {
      elements.tableCutoffAll.innerHTML = allVouchers.map(r => {
        let statusBadge = '<span class="badge badge-success">COMPLIANT</span>';
        if (r.cutoff_status.includes("Early")) {
          statusBadge = '<span class="badge badge-danger">ERROR: Early Booking</span>';
        } else if (r.cutoff_status.includes("Late")) {
          statusBadge = '<span class="badge badge-warning">ERROR: Late Booking</span>';
        }

        return `
          <tr>
            <td><code>${r.voucher_no}</code></td>
            <td>${r.vendor}</td>
            <td>${r.invoice_no}</td>
            <td>${r.invoice_date}</td>
            <td>${r.grn_date}</td>
            <td class="text-right font-mono">${formatINR(r.amount)}</td>
            <td>${statusBadge}</td>
          </tr>
        `;
      }).join('');
    }
  }

  function loadCutoffSampleData() {
    const sampleRows = [
      { voucher_no: "V-201", vendor: "Apex Supplies", invoice_no: "INV-901", invoice_date: "2026-03-27", grn_date: "2026-03-29", amount: "120000.00" },
      { voucher_no: "V-202", vendor: "Southern Logistics", invoice_no: "INV-902", invoice_date: "2026-03-30", grn_date: "2026-04-03", amount: "450000.00" },
      { voucher_no: "V-203", vendor: "Deccan Tech", invoice_no: "INV-903", invoice_date: "2026-03-31", grn_date: "2026-04-05", amount: "310000.00" },
      { voucher_no: "V-204", vendor: "Metro Services", invoice_no: "INV-904", invoice_date: "2026-04-02", grn_date: "2026-03-28", amount: "280000.00" },
      { voucher_no: "V-205", vendor: "Coastal Engineering", invoice_no: "INV-905", invoice_date: "2026-04-03", grn_date: "2026-04-04", amount: "95000.00" },
      { voucher_no: "V-206", vendor: "Vanguard Legal", invoice_no: "INV-906", invoice_date: "2026-03-25", grn_date: "2026-03-26", amount: "160000.00" },
      { voucher_no: "V-207", vendor: "Apex Supplies", invoice_no: "INV-907", invoice_date: "2026-04-04", grn_date: "2026-03-29", amount: "520000.00" },
      { voucher_no: "V-208", vendor: "QuickFix Maint", invoice_no: "INV-908", invoice_date: "2026-03-29", grn_date: "2026-03-31", amount: "75000.00" },
      { voucher_no: "V-209", vendor: "Southern Logistics", invoice_no: "INV-909", invoice_date: "2026-03-31", grn_date: "2026-04-02", amount: "340000.00" },
      { voucher_no: "V-210", vendor: "Deccan Tech", invoice_no: "INV-910", invoice_date: "2026-04-05", grn_date: "2026-04-07", amount: "210000.00" }
    ];

    state.rawCutoffRows = sampleRows;
    const yearEnd = elements.cutoffDatePicker ? elements.cutoffDatePicker.value : "2026-03-31";
    const res = APAuditEngine.analyzeCutoffLedger(sampleRows, yearEnd);
    renderCutoffResults(res);
    showToast("Sample Cut-Off Ledger tested successfully!", "success");
  }

  // =========================================================================
  // MODULE 3: SA 520 ANALYTICAL PROCEDURES & DUPONT ROE
  // =========================================================================
  function renderSA520Results(saResults) {
    if (!saResults) return;
    state.sa520Data = saResults;
    const years = saResults.years || [];
    const scorecard = saResults.scorecard || [];
    const summary = saResults.summary || saResults.kpis || {};
    const latest = summary.latestYear || (years.length > 0 ? years[years.length - 1] : null);

    if (!latest) return;

    // DuPont 3-Stage Ribbon
    elements.dupontNetMargin.textContent = formatPct(latest.ratios.netMarginPct);
    elements.dupontAssetTurnover.textContent = Number(latest.ratios.assetTurnover).toFixed(2) + 'x';
    elements.dupontEquityMultiplier.textContent = Number(latest.ratios.equityMultiplier).toFixed(2) + 'x';
    elements.dupontCompositeRoe.textContent = formatPct(latest.ratios.dupontRoePct);

    // Latest KPIs
    elements.kpiLatestRoe.textContent = formatPct(latest.ratios.dupontRoePct);
    elements.kpiLatestNetMargin.textContent = formatPct(latest.ratios.netMarginPct);
    elements.kpiLatestAssetTurnover.textContent = Number(latest.ratios.assetTurnover).toFixed(2) + 'x';
    elements.kpiLatestEquityMultiplier.textContent = Number(latest.ratios.equityMultiplier).toFixed(2) + 'x';
    elements.kpiLatestDebtToEquity.textContent = Number(latest.ratios.debtToEquity).toFixed(2) + 'x';

    // Badge
    elements.badgeSa520AlertCount.textContent = summary.totalAlerts;

    // Render Chart: DuPont ROE % vs Net Margin %
    renderTrendChart(years);

    // Render Sub-Tab 1: 10-Year Scorecard Table (Transposed: Ratios as rows, Years as columns)
    const yearLabels = years.map(y => y.financial_year);
    elements.scorecardHead.innerHTML = `
      <tr>
        <th style="min-width: 220px;">Financial Ratio / Metric</th>
        <th style="min-width: 60px;" class="text-center">Unit</th>
        ${yearLabels.map(yl => `<th class="text-right" style="min-width: 90px;">${yl}</th>`).join('')}
        <th class="text-right" style="min-width: 90px;">YoY Change</th>
      </tr>
    `;

    elements.scorecardBody.innerHTML = scorecard.map(row => {
      const isRoe = row.ratio.includes("DuPont ROE");
      const isMargin = row.ratio.includes("Margin");
      const highlightClass = isRoe ? 'style="font-weight: 700; background: rgba(31, 78, 120, 0.15);"' : (isMargin ? 'style="font-weight: 600;"' : '');

      return `
        <tr ${highlightClass}>
          <td><strong>${row.ratio}</strong></td>
          <td class="text-center font-mono" style="color: var(--text-muted);">${row.unit}</td>
          ${yearLabels.map(yl => {
            const val = row[yl];
            const fmt = (val === null || val === undefined) ? '-' : (row.unit === '%' ? formatPct(val) : Number(val).toFixed(2) + (row.unit === 'x' ? 'x' : ''));
            return `<td class="text-right font-mono">${fmt}</td>`;
          }).join('')}
          <td class="text-right font-mono" style="font-weight: 600; color: ${row.yoy_change > 0 ? 'var(--accent-emerald)' : (row.yoy_change < 0 ? 'var(--accent-rose)' : 'inherit')};">
            ${row.yoy_change !== null && row.yoy_change !== undefined ? (row.yoy_change > 0 ? '+' : '') + Number(row.yoy_change).toFixed(2) : '-'}
          </td>
        </tr>
      `;
    }).join('');

    // Render Sub-Tab 2: SA 520 Risk Alerts Table
    elements.sa520AlertsBody.innerHTML = years.map(y => {
      const sol = y.statuses.solvencyStatus;
      const col = y.statuses.collectionStatus;
      const cov = y.statuses.coverageStatus;

      const solBadge = sol.includes("ALERT")
        ? `<span class="badge badge-danger">${sol}</span>`
        : `<span class="badge badge-success">${sol}</span>`;

      const colBadge = col.includes("ALERT")
        ? `<span class="badge badge-danger">${col}</span>`
        : `<span class="badge badge-success">${col}</span>`;

      const covBadge = cov.includes("ALERT")
        ? `<span class="badge badge-danger">${cov}</span>`
        : `<span class="badge badge-success">${cov}</span>`;

      const overall = y.flags.length > 0
        ? `<span class="badge badge-danger">${y.flags.length} Red-Flags</span>`
        : `<span class="badge badge-success">Acceptable Risk</span>`;

      return `
        <tr>
          <td><strong>${y.financial_year}</strong></td>
          <td>${solBadge}</td>
          <td>${colBadge}</td>
          <td>${covBadge}</td>
          <td class="text-center">${overall}</td>
        </tr>
      `;
    }).join('');

    // Render Sub-Tab 3: Full Financial Statement Dataset Table
    elements.sa520FullHead.innerHTML = `
      <tr>
        <th>Financial Year</th>
        <th class="text-right">Sales / Revenue</th>
        <th class="text-right">COGS</th>
        <th class="text-right">EBIT</th>
        <th class="text-right">Interest</th>
        <th class="text-right">Net Profit</th>
        <th class="text-right">Total Debt</th>
        <th class="text-right">Shareholders' Equity</th>
        <th class="text-right">Current Assets</th>
        <th class="text-right">Current Liab</th>
        <th class="text-right">Receivables</th>
        <th class="text-right">Total Assets</th>
      </tr>
    `;

    elements.sa520FullBody.innerHTML = years.map(y => `
      <tr>
        <td><strong>${y.financial_year}</strong></td>
        <td class="text-right font-mono">${formatINR(y.raw.revenue)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.cogs)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.ebit)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.interest_expense)}</td>
        <td class="text-right font-mono" style="font-weight: 700; color: var(--accent-emerald);">${formatINR(y.raw.net_profit)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.total_debt)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.shareholders_equity)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.current_assets)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.current_liabilities)}</td>
        <td class="text-right font-mono">${formatINR(y.raw.trade_receivables)}</td>
        <td class="text-right font-mono font-bold">${formatINR(y.raw.total_assets)}</td>
      </tr>
    `).join('');
  }

  function renderTrendChart(years) {
    if (!elements.chartCanvas) return;

    const labels = years.map(y => y.financial_year);
    const roeData = years.map(y => y.ratios.dupontRoePct);
    const marginData = years.map(y => y.ratios.netMarginPct);

    if (state.chartDupont) {
      state.chartDupont.destroy();
    }

    const ctx = elements.chartCanvas.getContext('2d');
    state.chartDupont = new Chart(ctx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'DuPont ROE (%)',
            data: roeData,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.1)',
            borderWidth: 3,
            fill: true,
            tension: 0.3,
            pointBackgroundColor: '#10b981',
            pointRadius: 5
          },
          {
            label: 'Net Margin (%)',
            data: marginData,
            borderColor: '#38bdf8',
            backgroundColor: 'transparent',
            borderWidth: 2,
            borderDash: [5, 5],
            tension: 0.3,
            pointBackgroundColor: '#38bdf8',
            pointRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              color: state.theme === 'dark' ? '#f8fafc' : '#0f172a',
              font: { weight: 'bold' }
            }
          },
          tooltip: {
            callbacks: {
              label: function (context) {
                return `${context.dataset.label}: ${context.parsed.y.toFixed(2)}%`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: {
              color: state.theme === 'dark' ? '#23324d' : '#e2e8f0'
            },
            ticks: {
              color: state.theme === 'dark' ? '#94a3b8' : '#475569'
            }
          },
          y: {
            grid: {
              color: state.theme === 'dark' ? '#23324d' : '#e2e8f0'
            },
            ticks: {
              color: state.theme === 'dark' ? '#94a3b8' : '#475569',
              callback: function (val) {
                return val + '%';
              }
            }
          }
        }
      }
    });
  }

  function loadInfosysSample() {
    try {
      const records = SA520Engine.infosys10Year.map(item => ({
        financial_year: item.year,
        revenue: 1000000.0,
        cogs: 600000.0,
        ebit: (item.ebitMargin / 100) * 1000000.0,
        interest_expense: item.interestCoverage ? ((item.ebitMargin / 100) * 1000000.0) / item.interestCoverage : 0,
        net_profit: (item.netMargin / 100) * 1000000.0,
        current_assets: 500000.0,
        inventory: 0,
        trade_receivables: ((item.receivablesDays * 1000000.0) / 365),
        cash_bank: 200000.0,
        current_liabilities: 500000.0 / (item.currentRatio || 1.5),
        total_debt: (1000000.0 / (item.assetTurnover || 1.0)) * (item.debtToEquity || 0),
        shareholders_equity: (1000000.0 / (item.assetTurnover || 1.0)) / (item.equityMultiplier || 1.5),
        total_assets: 1000000.0 / (item.assetTurnover || 1.0)
      }));

      const res = SA520Engine.analyzeFinancialData(records);
      renderSA520Results(res);
      showToast("10-Year Infosys Real Screener Model loaded successfully!", "success");
    } catch (e) {
      showToast(`Error loading Infosys model: ${e.message}`, "error");
    }
  }

  function loadBaselineSample() {
    try {
      const res = SA520Engine.analyzeFinancialData(SA520Engine.baseline3Year);
      renderSA520Results(res);
      showToast("3-Year Baseline Sample loaded!", "success");
    } catch (e) {
      showToast(`Error loading baseline: ${e.message}`, "error");
    }
  }

  // =========================================================================
  // FILE UPLOAD HANDLERS
  // =========================================================================
  function handleAPFileUpload(file) {
    if (!file) return;
    const reader = new FileReader();

    if (file.name.endsWith('.csv')) {
      reader.onload = function (e) {
        try {
          const csvText = e.target.result;
          const workbook = XLSX.read(csvText, { type: 'string' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(firstSheet);
          const results = APAuditEngine.analyzeAPLedger(rawRows, "2026-03-31");
          renderAPAuditResults(results);
          showToast(`Uploaded ${file.name} analyzed (${rawRows.length} vouchers)!`, "success");
        } catch (err) {
          showToast(`Error parsing CSV: ${err.message}`, "error");
        }
      };
      reader.readAsText(file);
    } else {
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(firstSheet);
          const results = APAuditEngine.analyzeAPLedger(rawRows, "2026-03-31");
          renderAPAuditResults(results);
          showToast(`Uploaded ${file.name} analyzed (${rawRows.length} vouchers)!`, "success");
        } catch (err) {
          showToast(`Error parsing Excel: ${err.message}`, "error");
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function handleCutoffFileUpload(file) {
    if (!file) return;
    const reader = new FileReader();

    const processRows = (rawRows) => {
      state.rawCutoffRows = rawRows;
      const yearEnd = elements.cutoffDatePicker ? elements.cutoffDatePicker.value : "2026-03-31";
      const results = APAuditEngine.analyzeCutoffLedger(rawRows, yearEnd);
      renderCutoffResults(results);
      showToast(`Uploaded ${file.name} analyzed for Cut-Off boundary!`, "success");
    };

    if (file.name.endsWith('.csv')) {
      reader.onload = function (e) {
        try {
          const workbook = XLSX.read(e.target.result, { type: 'string' });
          const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          processRows(rawRows);
        } catch (err) {
          showToast(`Error parsing Cut-off CSV: ${err.message}`, "error");
        }
      };
      reader.readAsText(file);
    } else {
      reader.onload = function (e) {
        try {
          const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          processRows(rawRows);
        } catch (err) {
          showToast(`Error parsing Cut-off Excel: ${err.message}`, "error");
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function handleScreenerFileUpload(file) {
    if (!file) return;
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const records = SA520Engine.parseScreenerWorkbook(workbook);
        const results = SA520Engine.analyzeFinancialData(records);
        renderSA520Results(results);
        showToast(`Uploaded Screener.in workbook (${records.length} years) analyzed!`, "success");
      } catch (err) {
        showToast(`Error parsing Screener workbook: ${err.message}`, "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // =========================================================================
  // EVENT WIRING & INITIALIZATION
  // =========================================================================
  function initEvents() {
    // Theme toggle
    if (elements.themeToggleBtn) {
      elements.themeToggleBtn.addEventListener('click', () => {
        applyTheme(state.theme === 'dark' ? 'light' : 'dark');
        if (state.sa520Data) {
          renderTrendChart(state.sa520Data.years);
        }
      });
    }

    // Nav buttons
    elements.navBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-target');
        if (target) switchModule(target);
      });
    });

    setupSubtabs();

    // Tab 1 Events
    if (elements.btnTestAPSample) {
      elements.btnTestAPSample.addEventListener('click', loadAPMockData);
    }
    if (elements.btnUploadAP) {
      elements.btnUploadAP.addEventListener('click', () => elements.inputAPFile.click());
    }
    if (elements.inputAPFile) {
      elements.inputAPFile.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleAPFileUpload(e.target.files[0]);
      });
    }
    if (elements.btnExportAPExcel) {
      elements.btnExportAPExcel.addEventListener('click', () => {
        if (!state.apAuditData) {
          showToast("Please analyze AP data first before exporting.", "error");
          return;
        }
        ExportService.exportMasterAPAudit(state.apAuditData);
        showToast("Exporting Master AP Audit Report (5 Tabs)...", "success");
      });
    }

    // Ledger filter & pagination
    if (elements.ledgerSearchInput) {
      elements.ledgerSearchInput.addEventListener('input', (e) => {
        state.ledgerFilter = e.target.value;
        state.ledgerPage = 1;
        renderFullLedgerTable();
      });
    }
    if (elements.btnPrevPage) {
      elements.btnPrevPage.addEventListener('click', () => {
        if (state.ledgerPage > 1) {
          state.ledgerPage--;
          renderFullLedgerTable();
        }
      });
    }
    if (elements.btnNextPage) {
      elements.btnNextPage.addEventListener('click', () => {
        state.ledgerPage++;
        renderFullLedgerTable();
      });
    }

    // Tab 2 Events
    if (elements.btnTestCutoffSample) {
      elements.btnTestCutoffSample.addEventListener('click', loadCutoffSampleData);
    }
    if (elements.btnUploadCutoff) {
      elements.btnUploadCutoff.addEventListener('click', () => elements.inputCutoffFile.click());
    }
    if (elements.inputCutoffFile) {
      elements.inputCutoffFile.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleCutoffFileUpload(e.target.files[0]);
      });
    }
    if (elements.cutoffDatePicker) {
      elements.cutoffDatePicker.addEventListener('change', () => {
        if (state.rawCutoffRows) {
          const res = APAuditEngine.analyzeCutoffLedger(state.rawCutoffRows, elements.cutoffDatePicker.value);
          renderCutoffResults(res);
          showToast(`Recalculated cut-off boundary for ${elements.cutoffDatePicker.value}`, "info");
        }
      });
    }
    if (elements.btnExportCutoffExcel) {
      elements.btnExportCutoffExcel.addEventListener('click', () => {
        if (!state.cutoffData) {
          showToast("Please test or upload cut-off data first.", "error");
          return;
        }
        ExportService.exportCutoffAudit(state.cutoffData.exceptions, state.cutoffData.allVouchers);
        showToast("Exporting Cut-Off Audit Report (2 Tabs)...", "success");
      });
    }

    // Tab 3 Events
    if (elements.btnTestInfosysSample) {
      elements.btnTestInfosysSample.addEventListener('click', loadInfosysSample);
    }
    if (elements.btnLoad3YrBaseline) {
      elements.btnLoad3YrBaseline.addEventListener('click', loadBaselineSample);
    }
    if (elements.btnUploadScreener) {
      elements.btnUploadScreener.addEventListener('click', () => elements.inputScreenerFile.click());
    }
    if (elements.inputScreenerFile) {
      elements.inputScreenerFile.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleScreenerFileUpload(e.target.files[0]);
      });
    }
    if (elements.btnExportSA520Excel) {
      elements.btnExportSA520Excel.addEventListener('click', () => {
        if (!state.sa520Data) {
          showToast("Please test or upload Screener data first.", "error");
          return;
        }
        ExportService.exportSA520Audit(state.sa520Data);
        showToast("Exporting SA 520 Ratio Audit Report (3 Tabs)...", "success");
      });
    }
  }

  // Self-initialization
  document.addEventListener('DOMContentLoaded', () => {
    applyTheme(state.theme);
    initEvents();

    // Auto-populate all 3 modules on first open for an instantaneous "working app" experience!
    loadAPMockData();
    loadCutoffSampleData();
    loadInfosysSample();

    showToast("CA Forensic Audit Suite loaded & ready!", "success");
  });

})();
