/**
 * Excel Export Service
 * Produces multi-tab workbooks identical in structure to openpyxl deliverables
 * for AP Audit, Year-End Cut-Off, and SA 520 DuPont ROE analysis.
 */

(function (window) {
  'use strict';

  function downloadWorkbook(workbook, filename) {
    if (typeof XLSX === 'undefined') {
      alert("SheetJS library is loading or unavailable. Please check your internet connection.");
      return;
    }
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx', type: 'binary' });
  }

  /**
   * Generates Project 1: Master AP Audit Report (5 Tabs)
   */
  function exportMasterAPAudit(auditResults, filename = "Project1_AP_Audit_Report.xlsx") {
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

    // 5. Full Clean Ledger
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
    XLSX.utils.book_append_sheet(wb, wsFull, "5_Full_Clean_Ledger");

    downloadWorkbook(wb, filename);
  }

  /**
   * Generates Project 2: Year-End Cut-Off Audit Report (2 Tabs)
   */
  function exportCutoffAudit(cutoffExceptions, allVouchers, filename = "Project2_Cutoff_Audit_Report.xlsx") {
    const wb = XLSX.utils.book_new();

    const exceptionsData = cutoffExceptions.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "invoice_no": r.invoice_no || "",
      "invoice_date": r.invoice_date,
      "grn_date": r.grn_date,
      "amount": r.amount,
      "cutoff_status": r.cutoff_status
    }));
    const wsExceptions = XLSX.utils.json_to_sheet(exceptionsData);
    XLSX.utils.book_append_sheet(wb, wsExceptions, "1_Cutoff_Exceptions");

    const allData = allVouchers.map(r => ({
      "voucher_no": r.voucher_no,
      "vendor": r.vendor,
      "invoice_no": r.invoice_no || "",
      "invoice_date": r.invoice_date,
      "grn_date": r.grn_date,
      "amount": r.amount,
      "cutoff_status": r.cutoff_status
    }));
    const wsAll = XLSX.utils.json_to_sheet(allData);
    XLSX.utils.book_append_sheet(wb, wsAll, "2_All_Tested_Vouchers");

    downloadWorkbook(wb, filename);
  }

  /**
   * Generates Project 3: SA 520 Financial Ratio & DuPont ROE Audit Report (3 Tabs)
   */
  function exportSA520Audit(saResults, filename = "Project3_SA520_Ratio_Audit.xlsx") {
    if (!saResults) return;
    const { scorecard, years } = saResults;

    const wb = XLSX.utils.book_new();

    // 1. Ratio Scorecard
    const wsScorecard = XLSX.utils.json_to_sheet(scorecard);
    XLSX.utils.book_append_sheet(wb, wsScorecard, "1_10Yr_Ratio_Scorecard");

    // 2. Risk Alerts
    const flagsData = years.map(y => ({
      "Financial_Year": y.financial_year,
      "Solvency_Flag": y.statuses.solvencyStatus,
      "Collection_Flag": y.statuses.collectionStatus,
      "Coverage_Flag": y.statuses.coverageStatus
    }));
    const wsFlags = XLSX.utils.json_to_sheet(flagsData);
    XLSX.utils.book_append_sheet(wb, wsFlags, "2_SA520_Risk_Alerts");

    // 3. Full Extracted Data
    const fullData = years.map(y => ({
      "Financial_Year": y.financial_year,
      "Sales": y.raw.revenue,
      "COGS": y.raw.cogs,
      "EBIT": y.raw.ebit,
      "Interest": y.raw.interest_expense,
      "Net_Profit": y.raw.net_profit,
      "Shareholders_Equity": y.raw.shareholders_equity,
      "Borrowings": y.raw.total_debt,
      "Current_Liabilities": y.raw.current_liabilities,
      "Inventory": y.raw.inventory,
      "Receivables": y.raw.trade_receivables,
      "Cash_Bank": y.raw.cash_bank,
      "Total_Assets": y.raw.total_assets,
      "EBIT_Margin_%": y.ratios.ebitMarginPct,
      "Net_Margin_%": y.ratios.netMarginPct,
      "Asset_Turnover_x": y.ratios.assetTurnover,
      "Equity_Multiplier_x": y.ratios.equityMultiplier,
      "DuPont_ROE_%": y.ratios.dupontRoePct,
      "Debt_to_Equity": y.ratios.debtToEquity,
      "Interest_Coverage": y.ratios.interestCoverage,
      "Approx_Current_Ratio": y.ratios.currentRatio,
      "Receivables_Days": y.ratios.receivablesDays,
      "Inventory_Days": y.ratios.inventoryDays
    }));
    const wsFull = XLSX.utils.json_to_sheet(fullData);
    XLSX.utils.book_append_sheet(wb, wsFull, "3_Full_Extracted_Data");

    downloadWorkbook(wb, filename);
  }

  // Export to global scope
  window.ExportService = {
    exportMasterAPAudit,
    exportCutoffAudit,
    exportSA520Audit,
    downloadWorkbook
  };

})(window);
