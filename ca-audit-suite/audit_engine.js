/**
 * AP Ledger Forensic Audit & Statutory Compliance Engine
 * Replicates and enhances the logic from Google Colab Cells 0, 1, 2, and 3.
 */

(function (window) {
  'use strict';

  // Seeded Random Number Generator (Mulberry32) for deterministic mock data
  function Mulberry32(seed) {
    return function () {
      let t = (seed += 0x6d2b79f5);
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const vendors = [
    { name: "Apex Consulting LLP", gstin: "37AABCA1234D1Z5", sec: "194J" },
    { name: "Southern Logistics Ltd", gstin: "33AAACS5678E1Z2", sec: "194C" },
    { name: "Deccan Tech Solutions", gstin: "36AABCD9012F1Z8", sec: "194J" },
    { name: "Metro Facility Services", gstin: "29AAACM3456G1Z1", sec: "194C" },
    { name: "Coastal Engineering Co", gstin: "37AAACC7890H1Z9", sec: "194C" },
    { name: "Vanguard Legal Associates", gstin: "27AAACV2345J1Z4", sec: "194J" },
    { name: "QuickFix Maintenance", gstin: "33AAACQ6789K1Z7", sec: "194C" }
  ];

  const candidateAmounts = [18000, 25000, 45000, 75000, 120000, 185000, 250000];

  const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;

  /**
   * Generates a 500-row synthetic ERP Ledger mirroring Colab seed(42)
   */
  function generateMockLedger(nRows = 500, seed = 42) {
    const rng = Mulberry32(seed);
    const rows = [];

    for (let i = 1; i <= nRows; i++) {
      const vIdx = Math.floor(rng() * vendors.length);
      const v = vendors[vIdx];

      const aIdx = Math.floor(rng() * candidateAmounts.length);
      const amt = candidateAmounts[aIdx];
      const gst = Math.round(amt * 0.18 * 100) / 100;

      const rate = v.sec === "194J" ? 0.10 : 0.02;
      const tds = amt > 30000 ? Math.round(amt * rate * 100) / 100 : 0.0;

      const invDay = 1 + Math.floor(rng() * 24);
      const invDayStr = String(invDay).padStart(2, '0');
      const grnDayStr = String(invDay + 2).padStart(2, '0');

      rows.push({
        voucher_no: `V-${1000 + i}`,
        vendor: v.name,
        gstin: v.gstin,
        invoice_no: `INV-${50000 + i}`,
        invoice_date: `2026-03-${invDayStr}`,
        grn_date: `2026-03-${grnDayStr}`,
        section: v.sec,
        amount: amt.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        gst_claimed: gst.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        tds_deducted: tds.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
      });
    }

    // --- INJECT REALISTIC AUDIT ANOMALIES (Identical to Colab Notebook) ---
    // 1. Missing vendor names
    [14, 189, 412].forEach(idx => {
      if (rows[idx]) rows[idx].vendor = null;
    });

    // 2. Broken & missing GSTINs
    [22, 105, 310].forEach(idx => {
      if (rows[idx]) rows[idx].gstin = "33AAACS5678E1Z"; // 14 chars
    });
    [67, 245, 455].forEach(idx => {
      if (rows[idx]) rows[idx].gstin = "29INVALIDGST999"; // Corrupt PAN
    });
    [91, 378].forEach(idx => {
      if (rows[idx]) rows[idx].gstin = null; // Missing
    });

    // 3. Duplicate invoices (Double-payment risks)
    if (rows[341] && rows[18]) {
      rows[341].invoice_no = rows[18].invoice_no;
    }
    if (rows[488] && rows[112]) {
      rows[488].invoice_no = rows[112].invoice_no;
      rows[488].vendor = "Southern Logistics Ltd";
    }

    // 4. TDS Shortfalls
    [9, 74, 215, 399].forEach(idx => {
      if (rows[idx]) rows[idx].tds_deducted = "0.00";
    });
    [150, 290].forEach(idx => {
      if (rows[idx]) rows[idx].tds_deducted = "500.00";
    });

    // 5. Cut-off violations (March 31, 2026 boundary)
    if (rows[495]) {
      rows[495].invoice_date = "2026-03-30";
      rows[495].grn_date = "2026-04-03";
    }
    if (rows[496]) {
      rows[496].invoice_date = "2026-03-31";
      rows[496].grn_date = "2026-04-05";
    }
    if (rows[497]) {
      rows[497].invoice_date = "2026-04-02";
      rows[497].grn_date = "2026-03-28";
    }
    if (rows[498]) {
      rows[498].invoice_date = "2026-04-04";
      rows[498].grn_date = "2026-03-29";
    }

    return rows;
  }

  /**
   * Parses raw amounts, handles strings with commas, nulls, etc.
   */
  function parseMoney(val) {
    if (val === null || val === undefined || val === '') return 0.0;
    if (typeof val === 'number') return isNaN(val) ? 0.0 : val;
    const str = String(val).replace(/,/g, '').trim();
    const parsed = parseFloat(str);
    return isNaN(parsed) ? 0.0 : parsed;
  }

  /**
   * Runs the full forensic audit analysis on raw ledger rows
   */
  function analyzeAPLedger(rawRows, yearEndDateStr = "2026-03-31") {
    if (!Array.isArray(rawRows) || rawRows.length === 0) {
      throw new Error("No ledger data provided for audit analysis.");
    }

    const yearEnd = new Date(yearEndDateStr + "T23:59:59");

    // Pass 1: Sanitize and compute base statutory fields
    const cleanLedger = rawRows.map((r, i) => {
      const voucherNo = r.voucher_no || r.voucher || `V-${1000 + i + 1}`;
      const vendor = (r.vendor && String(r.vendor).trim()) ? String(r.vendor).trim() : "UNKNOWN_VENDOR";
      const rawGstin = r.gstin ? String(r.gstin).trim().toUpperCase() : "MISSING";
      const gstin = rawGstin === "" ? "MISSING" : rawGstin;
      const invoiceNo = r.invoice_no || r.invoice || `INV-${50000 + i + 1}`;
      const sec = (r.section || "194C").toUpperCase();

      const amount = parseMoney(r.amount);
      const gstClaimed = parseMoney(r.gst_claimed);
      const tdsDeducted = parseMoney(r.tds_deducted);

      // TDS Statutory Logic
      const rateMap = { "194J": 0.10, "194C": 0.02 };
      const tdsRate = rateMap[sec] || 0.0;
      const expectedTds = amount > 30000 ? Math.round(amount * tdsRate * 100) / 100 : 0.0;
      const tdsShortfall = Math.max(0, Math.round((expectedTds - tdsDeducted) * 100) / 100);
      const disallowance30pct = tdsShortfall > 0 ? Math.round(amount * 0.30 * 100) / 100 : 0.0;

      // GSTIN Statutory Check
      const isValidGstin = GSTIN_REGEX.test(gstin);

      // Cut-off Check
      const invDateStr = r.invoice_date || "";
      const grnDateStr = r.grn_date || "";
      let cutoffStatus = "COMPLIANT";
      let isCutoffError = false;

      if (invDateStr && grnDateStr) {
        const invDate = new Date(invDateStr + "T00:00:00");
        const grnDate = new Date(grnDateStr + "T00:00:00");

        if (!isNaN(invDate) && !isNaN(grnDate)) {
          const earlyBooking = (invDate <= yearEnd) && (grnDate > yearEnd);
          const lateBooking = (invDate > yearEnd) && (grnDate <= yearEnd);

          if (earlyBooking) {
            cutoffStatus = "ERROR: Early Booking (Goods arrived in April)";
            isCutoffError = true;
          } else if (lateBooking) {
            cutoffStatus = "ERROR: Late Booking (Goods arrived in March)";
            isCutoffError = true;
          }
        }
      }

      return {
        id: i + 1,
        voucher_no: voucherNo,
        vendor,
        gstin,
        invoice_no: invoiceNo,
        invoice_date: invDateStr,
        grn_date: grnDateStr,
        section: sec,
        amount,
        gst_claimed: gstClaimed,
        tds_deducted: tdsDeducted,
        tds_rate: tdsRate,
        expected_tds: expectedTds,
        tds_shortfall: tdsShortfall,
        disallowance_30pct: disallowance30pct,
        is_valid_gstin: isValidGstin,
        cutoff_status: cutoffStatus,
        is_cutoff_error: isCutoffError,
        is_duplicate: false // Will be set in Pass 2
      };
    });

    // Pass 2: Detect duplicate invoices
    const invoiceCounts = {};
    cleanLedger.forEach(row => {
      const inv = row.invoice_no;
      invoiceCounts[inv] = (invoiceCounts[inv] || 0) + 1;
    });

    cleanLedger.forEach(row => {
      if (invoiceCounts[row.invoice_no] > 1) {
        row.is_duplicate = true;
      }
    });

    // Exception Subsets
    const duplicates = cleanLedger.filter(r => r.is_duplicate).sort((a, b) => a.invoice_no.localeCompare(b.invoice_no));
    const tdsShortfalls = cleanLedger.filter(r => r.tds_shortfall > 0);
    const invalidGstins = cleanLedger.filter(r => !r.is_valid_gstin);
    const cutoffErrors = cleanLedger.filter(r => r.is_cutoff_error);

    // Pass 3: Vendor Risk Summary
    const vendorMap = {};
    cleanLedger.forEach(r => {
      const key = `${r.vendor}|||${r.gstin}`;
      if (!vendorMap[key]) {
        vendorMap[key] = {
          vendor: r.vendor,
          gstin: r.gstin,
          voucher_count: 0,
          total_spend: 0,
          total_gst_claimed: 0,
          total_tds_shortfall: 0,
          total_disallowance_30pct: 0,
          has_invalid_gst: !r.is_valid_gstin,
          has_missing_name: r.vendor === "UNKNOWN_VENDOR"
        };
      }
      const item = vendorMap[key];
      item.voucher_count += 1;
      item.total_spend += r.amount;
      item.total_gst_claimed += r.gst_claimed;
      item.total_tds_shortfall += r.tds_shortfall;
      item.total_disallowance_30pct += r.disallowance_30pct;
    });

    const vendorSummary = Object.values(vendorMap).map(v => ({
      ...v,
      total_spend: Math.round(v.total_spend * 100) / 100,
      total_gst_claimed: Math.round(v.total_gst_claimed * 100) / 100,
      total_tds_shortfall: Math.round(v.total_tds_shortfall * 100) / 100,
      total_disallowance_30pct: Math.round(v.total_disallowance_30pct * 100) / 100,
      risk_score: (v.total_tds_shortfall > 0 ? 30 : 0) + (v.has_invalid_gst ? 40 : 0) + (v.has_missing_name ? 30 : 0)
    })).sort((a, b) => b.total_spend - a.total_spend);

    // High Level KPIs
    const totalSpend = cleanLedger.reduce((sum, r) => sum + r.amount, 0);
    const totalGstClaimed = cleanLedger.reduce((sum, r) => sum + r.gst_claimed, 0);
    const totalTdsShortfall = tdsShortfalls.reduce((sum, r) => sum + r.tds_shortfall, 0);
    const totalDisallowance30 = cleanLedger.reduce((sum, r) => sum + r.disallowance_30pct, 0);
    const totalDuplicateExposure = duplicates.reduce((sum, r) => sum + r.amount, 0);
    const totalCutoffExposure = cutoffErrors.reduce((sum, r) => sum + r.amount, 0);
    const totalInvalidGstExposure = invalidGstins.reduce((sum, r) => sum + r.gst_claimed, 0);

    const kpisSummary = {
      totalVouchers: cleanLedger.length,
      totalSpend: Math.round(totalSpend * 100) / 100,
      totalGstClaimed: Math.round(totalGstClaimed * 100) / 100,
      totalTdsShortfall: Math.round(totalTdsShortfall * 100) / 100,
      totalDisallowance30: Math.round(totalDisallowance30 * 100) / 100,
      duplicateCount: duplicates.length,
      duplicateExposure: Math.round(totalDuplicateExposure * 100) / 100,
      tdsShortfallCount: tdsShortfalls.length,
      invalidGstinCount: invalidGstins.length,
      invalidGstExposure: Math.round(totalInvalidGstExposure * 100) / 100,
      cutoffErrorCount: cutoffErrors.length,
      cutoffExposure: Math.round(totalCutoffExposure * 100) / 100,
      vendorCount: vendorSummary.length
    };

    return {
      summary: kpisSummary,
      kpis: kpisSummary,
      exceptions: {
        duplicates,
        tdsShortfalls,
        invalidGstins,
        cutoffErrors,
        vendorSummary
      },
      fullLedger: cleanLedger
    };
  }

  /**
   * Dedicated Cut-off Boundary Audit Analyzer
   */
  function analyzeCutoffLedger(rawRows, yearEndDateStr = "2026-03-31") {
    if (!Array.isArray(rawRows) || rawRows.length === 0) {
      throw new Error("No ledger data provided for cut-off audit analysis.");
    }
    const yearEnd = new Date(yearEndDateStr + "T23:59:59");

    let totalErrors = 0;
    let totalErrorAmt = 0;
    let earlyCount = 0;
    let earlyAmt = 0;
    let lateCount = 0;
    let lateAmt = 0;

    const testedRows = rawRows.map((r, i) => {
      const voucherNo = r.voucher_no || r.voucher || `V-${1000 + i + 1}`;
      const vendor = (r.vendor && String(r.vendor).trim()) ? String(r.vendor).trim() : "UNKNOWN_VENDOR";
      const invoiceNo = r.invoice_no || r.invoice || `INV-${50000 + i + 1}`;
      const amount = parseMoney(r.amount);
      const invDateStr = r.invoice_date || "";
      const grnDateStr = r.grn_date || "";

      let cutoffStatus = "COMPLIANT";
      let isEarly = false;
      let isLate = false;

      if (invDateStr && grnDateStr) {
        const invDate = new Date(invDateStr + "T00:00:00");
        const grnDate = new Date(grnDateStr + "T00:00:00");

        if (!isNaN(invDate) && !isNaN(grnDate)) {
          if (invDate <= yearEnd && grnDate > yearEnd) {
            cutoffStatus = "ERROR: Early Booking (Goods arrived after year-end)";
            isEarly = true;
          } else if (invDate > yearEnd && grnDate <= yearEnd) {
            cutoffStatus = "ERROR: Late Booking (Goods arrived before year-end)";
            isLate = true;
          }
        }
      }

      if (isEarly) {
        earlyCount++;
        earlyAmt += amount;
        totalErrors++;
        totalErrorAmt += amount;
      } else if (isLate) {
        lateCount++;
        lateAmt += amount;
        totalErrors++;
        totalErrorAmt += amount;
      }

      return {
        voucher_no: voucherNo,
        vendor: vendor,
        invoice_no: invoiceNo,
        invoice_date: invDateStr,
        grn_date: grnDateStr,
        amount: amount,
        cutoff_status: cutoffStatus,
        is_error: isEarly || isLate
      };
    });

    const exceptions = testedRows.filter(r => r.is_error);

    const cutoffSummary = {
      testedCount: testedRows.length,
      totalErrors: totalErrors,
      totalErrorAmt: Math.round(totalErrorAmt * 100) / 100,
      earlyCount: earlyCount,
      earlyAmt: Math.round(earlyAmt * 100) / 100,
      lateCount: lateCount,
      lateAmt: Math.round(lateAmt * 100) / 100,
      yearEnd: yearEndDateStr
    };

    return {
      summary: cutoffSummary,
      kpis: cutoffSummary,
      exceptions: exceptions,
      allVouchers: testedRows
    };
  }

  // Export to global scope
  window.APAuditEngine = {
    generateMockLedger,
    analyzeAPLedger,
    analyzeCutoffLedger,
    parseMoney,
    vendors
  };

})(window);
