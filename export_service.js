/**
 * Excel Export Service
 * Produces multi-tab workbooks identical in structure to openpyxl deliverables
 * in Google Colab Cells 1, 3, 4, and 9.
 */

(function (window) {
  'use strict';

  function downloadWorkbook(workbook, filename) {
    if (typeof XLSX === 'undefined') {
      alert("SheetJS library is loading or unavailable. Please check your connection.");
      return;
    }
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx', type: 'binary' });
  }

  /**
   * Generates Master 6-Tab AP Audit Report Excel workbook
   */
  function exportMasterAPAudit(auditResults, filename = "Master_6Tab_AP_Audit_Report.xlsx") {
    if (!auditResults) return;
    const { exceptions, fullLedger } = auditResults;

    const wb = XLSX.utils.book_new();

    // 1. Vendor Summary
    const vendorSummaryData = exceptions.vendorSummary.map(v => ({
      "vendor": v.vendor,
      "gstin": v.gstin,
      "voucher_count": v.voucher_count,
      "total_spend": v.total_spend,
      "total_gst_claimed": v.total_gst_claimed,
      "total_tds_shortfall": v.total_tds_shortfall,
      "total_disallowance_30pct": v.total_disallowance_30pct
    }));
    const wsVendor = XLSX.utils.json_to_sheet(vendorSummaryData);
    XLSX.utils.book_append_sheet(wb, wsVendor, "1_Vendor_Summary");

    // 2. Duplicate Invoices
    const duplicatesData = exceptions.duplicates.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "gstin": r.gstin,
      "invoice_no": r.invoice_no,
      "invoice_date": r.invoice_date,
      "grn_date": r.grn_date,
      "section": r.section,
      "amount": r.amount,
      "gst_claimed": r.gst_claimed,
      "tds_deducted": r.tds_deducted
    }));
    const wsDup = XLSX.utils.json_to_sheet(duplicatesData);
    XLSX.utils.book_append_sheet(wb, wsDup, "2_Duplicate_Invoices");

    // 3. TDS Shortfalls
    const tdsData = exceptions.tdsShortfalls.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "gstin": r.gstin,
      "invoice_no": r.invoice_no,
      "section": r.section,
      "amount": r.amount,
      "tds_rate": r.tds_rate,
      "expected_tds": r.expected_tds,
      "tds_deducted": r.tds_deducted,
      "tds_shortfall": r.tds_shortfall,
      "disallowance_30pct": r.disallowance_30pct
    }));
    const wsTds = XLSX.utils.json_to_sheet(tdsData);
    XLSX.utils.book_append_sheet(wb, wsTds, "3_TDS_Shortfalls");

    // 4. Invalid GSTINs
    const gstData = exceptions.invalidGstins.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "gstin": r.gstin,
      "invoice_no": r.invoice_no,
      "gst_claimed": r.gst_claimed,
      "amount": r.amount
    }));
    const wsGst = XLSX.utils.json_to_sheet(gstData);
    XLSX.utils.book_append_sheet(wb, wsGst, "4_Invalid_GSTINs");

    // 5. Cutoff Errors
    const cutoffData = exceptions.cutoffErrors.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "invoice_date": r.invoice_date,
      "grn_date": r.grn_date,
      "amount": r.amount,
      "cutoff_status": r.cutoff_status
    }));
    const wsCutoff = XLSX.utils.json_to_sheet(cutoffData);
    XLSX.utils.book_append_sheet(wb, wsCutoff, "5_Cutoff_Errors");

    // 6. Full Clean Ledger
    const fullData = fullLedger.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "gstin": r.gstin,
      "invoice_no": r.invoice_no,
      "invoice_date": r.invoice_date,
      "grn_date": r.grn_date,
      "section": r.section,
      "amount": r.amount,
      "gst_claimed": r.gst_claimed,
      "tds_deducted": r.tds_deducted,
      "expected_tds": r.expected_tds,
      "tds_shortfall": r.tds_shortfall,
      "disallowance_30pct": r.disallowance_30pct,
      "is_valid_gstin": r.is_valid_gstin,
      "cutoff_status": r.cutoff_status
    }));
    const wsFull = XLSX.utils.json_to_sheet(fullData);
    XLSX.utils.book_append_sheet(wb, wsFull, "6_Full_Clean_Ledger");

    downloadWorkbook(wb, filename);
  }

  /**
   * Generates SA 520 Financial Ratio Audit Deliverable Excel workbook
   */
  function exportSA520Audit(saResults, filename = "SA520_Financial_Ratio_Audit.xlsx") {
    if (!saResults) return;
    const { scorecard, years } = saResults;

    const wb = XLSX.utils.book_new();

    // 1. Ratio Scorecard YoY
    const wsScorecard = XLSX.utils.json_to_sheet(scorecard);
    XLSX.utils.book_append_sheet(wb, wsScorecard, "1_Ratio_Scorecard_YoY");

    // 2. Risk Flags
    const flagsData = years.map(y => ({
      "Financial_Year": y.financial_year,
      "Liquidity_Flag": y.statuses.liquidityStatus,
      "Solvency_Flag": y.statuses.solvencyStatus,
      "Collection_Flag": y.statuses.collectionStatus,
      "Coverage_Flag": y.statuses.coverageStatus
    }));
    const wsFlags = XLSX.utils.json_to_sheet(flagsData);
    XLSX.utils.book_append_sheet(wb, wsFlags, "2_SA520_Risk_Flags");

    // 3. Full Financials & Extracted Data
    const fullData = years.map(y => ({
      "Financial_Year": y.financial_year,
      "Revenue": y.raw.revenue,
      "COGS": y.raw.cogs,
      "EBIT": y.raw.ebit,
      "Interest_Expense": y.raw.interest_expense,
      "Net_Profit": y.raw.net_profit,
      "Shareholders_Equity": y.raw.shareholders_equity,
      "Total_Debt": y.raw.total_debt,
      "Current_Assets": y.raw.current_assets,
      "Current_Liabilities": y.raw.current_liabilities,
      "Inventory": y.raw.inventory,
      "Trade_Receivables": y.raw.trade_receivables,
      "Cash_Bank": y.raw.cash_bank,
      "Total_Assets": y.raw.total_assets,
      "EBIT_Margin_%": y.ratios.ebitMarginPct,
      "Net_Margin_%": y.ratios.netMarginPct,
      "DuPont_ROE_%": y.ratios.dupontRoePct,
      "Current_Ratio": y.ratios.currentRatio,
      "Debt_to_Equity": y.ratios.debtToEquity,
      "Interest_Coverage": y.ratios.interestCoverage,
      "Receivables_Days": y.ratios.receivablesDays,
      "Inventory_Days": y.ratios.inventoryDays
    }));
    const wsFull = XLSX.utils.json_to_sheet(fullData);
    XLSX.utils.book_append_sheet(wb, wsFull, "3_Full_Financials_&_Ratios");

    downloadWorkbook(wb, filename);
  }

  // Export to global scope
  window.ExportService = {
    exportMasterAPAudit,
    exportSA520Audit,
    downloadWorkbook
  };

})(window);
