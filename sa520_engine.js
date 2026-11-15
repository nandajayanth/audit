/**
 * SA 520 Analytical Procedures & DuPont ROE Engine
 * Implements financial statement analysis, DuPont 3-stage decomposition,
 * automated statutory red-flagging, and Screener.in Excel extraction.
 * Replicates and extends Colab Cells 4, 5, 6, 7, 8, 9.
 */

(function (window) {
  'use strict';

  // 3-Year Baseline Dataset (from Colab Cell 4)
  const baseline3Year = [
    {
      financial_year: "FY 2023-24",
      revenue: 5000000.0,
      cogs: 3000000.0,
      ebit: 850000.0,
      interest_expense: 150000.0,
      net_profit: 520000.0,
      current_assets: 2200000.0,
      inventory: 600000.0,
      trade_receivables: 700000.0,
      cash_bank: 900000.0,
      current_liabilities: 1100000.0,
      total_debt: 1500000.0,
      shareholders_equity: 2500000.0,
      total_assets: 5100000.0
    },
    {
      financial_year: "FY 2024-25",
      revenue: 5600000.0,
      cogs: 3400000.0,
      ebit: 920000.0,
      interest_expense: 180000.0,
      net_profit: 550000.0,
      current_assets: 2100000.0,
      inventory: 750000.0,
      trade_receivables: 820000.0,
      cash_bank: 530000.0,
      current_liabilities: 1350000.0,
      total_debt: 1800000.0,
      shareholders_equity: 2700000.0,
      total_assets: 5850000.0
    },
    {
      financial_year: "FY 2025-26",
      revenue: 6800000.0,
      cogs: 3900000.0,
      ebit: 1350000.0,
      interest_expense: 260000.0,
      net_profit: 820000.0,
      current_assets: 1950000.0,
      inventory: 1100000.0,
      trade_receivables: 1250000.0,
      cash_bank: 400000.0,
      current_liabilities: 1650000.0,
      total_debt: 2600000.0,
      shareholders_equity: 2900000.0,
      total_assets: 7150000.0
    }
  ];

  // 10-Year Infosys Real Screener Model (Exact extracted numbers from Colab Cell 9)
  const infosys10Year = [
    { year: "Mar-2017", ebitMargin: 29.13, netMargin: 20.96, assetTurnover: 0.82, equityMultiplier: 1.21, dupontRoe: 20.81, debtToEquity: 0.00, interestCoverage: null, currentRatio: 2.47, receivablesDays: 65.7, inventoryDays: 0.0 },
    { year: "Mar-2018", ebitMargin: 28.74, netMargin: 22.73, assetTurnover: 0.89, equityMultiplier: 1.22, dupontRoe: 24.69, debtToEquity: 0.00, interestCoverage: null, currentRatio: 2.28, receivablesDays: 68.0, inventoryDays: 0.0 },
    { year: "Mar-2019", ebitMargin: 25.45, netMargin: 18.63, assetTurnover: 0.98, equityMultiplier: 1.29, dupontRoe: 23.72, debtToEquity: 0.00, interestCoverage: null, currentRatio: 1.80, receivablesDays: 65.5, inventoryDays: 0.0 },
    { year: "Mar-2020", ebitMargin: 24.43, netMargin: 18.28, assetTurnover: 0.99, equityMultiplier: 1.40, dupontRoe: 25.35, debtToEquity: 0.07, interestCoverage: 130.45, currentRatio: 1.71, receivablesDays: 74.3, inventoryDays: 0.0 },
    { year: "Mar-2021", ebitMargin: 26.70, netMargin: 19.26, assetTurnover: 0.93, equityMultiplier: 1.41, dupontRoe: 25.34, debtToEquity: 0.07, interestCoverage: 137.55, currentRatio: 1.70, receivablesDays: 70.1, inventoryDays: 0.0 },
    { year: "Mar-2022", ebitMargin: 24.92, netMargin: 18.18, assetTurnover: 1.04, equityMultiplier: 1.55, dupontRoe: 29.34, debtToEquity: 0.07, interestCoverage: 151.55, currentRatio: 1.12, receivablesDays: 68.1, inventoryDays: 0.0 },
    { year: "Mar-2023", ebitMargin: 22.90, netMargin: 16.42, assetTurnover: 1.18, equityMultiplier: 1.65, dupontRoe: 31.95, debtToEquity: 0.11, interestCoverage: 118.33, currentRatio: 0.92, receivablesDays: 63.2, inventoryDays: 0.0 },
    { year: "Mar-2024", ebitMargin: 23.72, netMargin: 17.07, assetTurnover: 1.13, equityMultiplier: 1.54, dupontRoe: 29.77, debtToEquity: 0.09, interestCoverage: 77.57, currentRatio: 1.14, receivablesDays: 71.7, inventoryDays: 0.0 },
    { year: "Mar-2025", ebitMargin: 23.33, netMargin: 16.39, assetTurnover: 1.10, equityMultiplier: 1.54, dupontRoe: 27.88, debtToEquity: 0.09, interestCoverage: 91.40, currentRatio: 1.27, receivablesDays: 69.8, inventoryDays: 0.0 },
    { year: "Mar-2026", ebitMargin: 22.62, netMargin: 16.48, assetTurnover: 1.16, equityMultiplier: 1.66, dupontRoe: 31.71, debtToEquity: 0.10, interestCoverage: 97.14, currentRatio: 1.10, receivablesDays: 72.0, inventoryDays: 0.0 }
  ];

  /**
   * Evaluates financial records and computes all ratios + red flags
   */
  function analyzeFinancialData(records) {
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error("No financial records provided for SA 520 analysis.");
    }

    const calculatedYears = records.map(r => {
      const year = r.financial_year || r.year || "FY";
      const rev = parseFloat(r.revenue || r.sales || 0);
      const cogs = parseFloat(r.cogs || 0);
      const ebit = parseFloat(r.ebit !== undefined ? r.ebit : (parseFloat(r.profit_before_tax || 0) + parseFloat(r.interest || r.interest_expense || 0)));
      const interest = parseFloat(r.interest_expense || r.interest || 0);
      const netProfit = parseFloat(r.net_profit || 0);
      const equity = parseFloat(r.shareholders_equity || r.equity || 0);
      const debt = parseFloat(r.total_debt || r.debt || r.borrowings || 0);
      const currLiab = parseFloat(r.current_liabilities || r.other_liabilities || 0);
      const inv = parseFloat(r.inventory || 0);
      const rec = parseFloat(r.trade_receivables || r.receivables || 0);
      const cash = parseFloat(r.cash_bank || r.cash || 0);

      // Current assets fallback
      const currAssets = parseFloat(r.current_assets || (inv + rec + cash) || 0);
      // Total assets fallback
      const totAssets = parseFloat(r.total_assets || (currAssets + equity + debt - currLiab) || (equity + debt) || 0);

      // 1. Profitability & DuPont
      const grossMarginPct = rev > 0 && cogs > 0 ? Math.round(((rev - cogs) / rev) * 10000) / 100 : null;
      const ebitMarginPct = rev > 0 ? Math.round((ebit / rev) * 10000) / 100 : 0.0;
      const netMarginPct = rev > 0 ? Math.round((netProfit / rev) * 10000) / 100 : 0.0;
      const assetTurnover = totAssets > 0 ? Math.round((rev / totAssets) * 100) / 100 : 0.0;
      const equityMultiplier = equity > 0 ? Math.round((totAssets / equity) * 100) / 100 : 1.0;
      const dupontRoePct = Math.round((netMarginPct * assetTurnover * equityMultiplier) * 100) / 100;
      const directRoePct = equity > 0 ? Math.round(((netProfit / equity) * 100) * 100) / 100 : 0.0;

      // 2. Liquidity
      const currentRatio = currLiab > 0 ? Math.round((currAssets / currLiab) * 100) / 100 : null;
      const quickRatio = currLiab > 0 ? Math.round(((currAssets - inv) / currLiab) * 100) / 100 : null;

      // 3. Solvency
      const debtToEquity = equity > 0 ? Math.round((debt / equity) * 100) / 100 : 0.0;
      const interestCoverage = interest > 0 ? Math.round((ebit / interest) * 100) / 100 : null;

      // 4. Working Capital Days
      const receivablesDays = rev > 0 ? Math.round(((rec / rev) * 365) * 10) / 10 : 0.0;
      const inventoryBasis = cogs > 0 ? cogs : rev;
      const inventoryDays = inventoryBasis > 0 ? Math.round(((inv / inventoryBasis) * 365) * 10) / 10 : 0.0;
      const inventoryTurnover = inv > 0 && cogs > 0 ? Math.round((cogs / inv) * 100) / 100 : null;

      // 5. SA 520 Automated Red-Flag Checks
      const flags = [];
      let liquidityStatus = "OK";
      let solvencyStatus = "OK";
      let collectionStatus = "OK";
      let coverageStatus = "OK";

      // Liquidity Flag (< 1.33x standard Indian banking MPBF limit)
      if (currentRatio !== null && currentRatio < 1.33) {
        liquidityStatus = "HIGH RISK: Liquidity Crunch (< 1.33x)";
        flags.push({ type: "liquidity", severity: "high", title: "Liquidity Crunch", detail: `Current Ratio is ${currentRatio}x, below the 1.33x statutory banking benchmark.` });
      }

      // Solvency Flag (> 0.85x or > 1.0x)
      if (debtToEquity > 0.85) {
        solvencyStatus = debtToEquity > 1.0 ? "ALERT: High Leverage (> 1.0x)" : "HIGH RISK: Heavy Debt (> 0.85x)";
        flags.push({ type: "solvency", severity: debtToEquity > 1.0 ? "danger" : "warning", title: "Heavy Debt Burden", detail: `Debt-to-Equity is ${debtToEquity}x, exceeding risk threshold.` });
      }

      // Collection / Fictitious Sales Flag (> 60 or > 75 days)
      if (receivablesDays > 60) {
        collectionStatus = receivablesDays > 75 ? "ALERT: Slow Collections / Fictitious Sales Risk (> 75d)" : "AUDIT ALERT: Slow Collections (> 60d)";
        flags.push({ type: "collection", severity: "warning", title: "Elevated Receivables Days", detail: `Debtor days spiked to ${receivablesDays} days. Potential revenue cut-off or unbilled inflation risk.` });
      }

      // Coverage Flag (< 2.5x)
      if (interestCoverage !== null && interestCoverage < 2.5) {
        coverageStatus = "ALERT: Weak Interest Cover (< 2.5x)";
        flags.push({ type: "coverage", severity: "danger", title: "Debt Service Strain", detail: `Interest coverage ratio is only ${interestCoverage}x, indicating difficulty meeting debt obligations.` });
      }

      return {
        financial_year: year,
        raw: {
          revenue: rev,
          cogs,
          ebit,
          interest_expense: interest,
          net_profit: netProfit,
          shareholders_equity: equity,
          total_debt: debt,
          current_assets: currAssets,
          current_liabilities: currLiab,
          inventory: inv,
          trade_receivables: rec,
          cash_bank: cash,
          total_assets: totAssets
        },
        ratios: {
          grossMarginPct,
          ebitMarginPct,
          netMarginPct,
          assetTurnover,
          equityMultiplier,
          dupontRoePct: dupontRoePct || directRoePct,
          currentRatio,
          quickRatio,
          debtToEquity,
          interestCoverage,
          receivablesDays,
          inventoryDays,
          inventoryTurnover
        },
        flags,
        statuses: {
          liquidityStatus,
          solvencyStatus,
          collectionStatus,
          coverageStatus
        }
      };
    });

    // Compute YoY Variance table (Transposed: Ratios as rows, Years as columns)
    const ratioKeys = [
      { key: "ebitMarginPct", label: "EBIT Margin %", unit: "%" },
      { key: "netMarginPct", label: "Net Margin %", unit: "%" },
      { key: "assetTurnover", label: "Asset Turnover", unit: "x" },
      { key: "equityMultiplier", label: "Equity Multiplier", unit: "x" },
      { key: "dupontRoePct", label: "DuPont ROE %", unit: "%" },
      { key: "debtToEquity", label: "Debt-to-Equity", unit: "x" },
      { key: "interestCoverage", label: "Interest Coverage", unit: "x" },
      { key: "currentRatio", label: "Current Ratio", unit: "x" },
      { key: "quickRatio", label: "Quick Ratio", unit: "x" },
      { key: "receivablesDays", label: "Receivables Days", unit: "d" },
      { key: "inventoryDays", label: "Inventory Days", unit: "d" }
    ];

    const scorecardRows = ratioKeys.map(rk => {
      const row = { ratio: rk.label, unit: rk.unit };
      calculatedYears.forEach(y => {
        row[y.financial_year] = y.ratios[rk.key];
      });

      // YoY variance between last two years if available
      if (calculatedYears.length >= 2) {
        const curr = calculatedYears[calculatedYears.length - 1].ratios[rk.key];
        const prev = calculatedYears[calculatedYears.length - 2].ratios[rk.key];
        if (curr !== null && prev !== null && curr !== undefined && prev !== undefined) {
          row.yoy_change = Math.round((curr - prev) * 100) / 100;
        } else {
          row.yoy_change = null;
        }
      }
      return row;
    });

    return {
      years: calculatedYears,
      scorecard: scorecardRows,
      summary: {
        totalYears: calculatedYears.length,
        latestYear: calculatedYears[calculatedYears.length - 1],
        totalAlerts: calculatedYears.reduce((sum, y) => sum + y.flags.length, 0)
      }
    };
  }

  /**
   * Flexible Screener.in Excel Data Sheet Extractor
   * Replicates Python functions from Colab Cells 5 - 9
   */
  function parseScreenerWorkbook(workbook) {
    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error("Invalid Excel workbook.");
    }

    // Locate target sheet
    let targetSheetName = "Data Sheet";
    if (!workbook.SheetNames.includes(targetSheetName)) {
      targetSheetName = workbook.SheetNames[workbook.SheetNames.length - 1];
    }

    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet) {
      throw new Error(`Target sheet '${targetSheetName}' not found.`);
    }

    // Convert worksheet to 2D array of rows
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
    if (!rawData || rawData.length === 0) {
      throw new Error("Sheet contains no data.");
    }

    function extractBlock(startMarker, stopMarkers) {
      let startIdx = -1;
      for (let i = 0; i < rawData.length; i++) {
        const cell = String(rawData[i][0] || '').trim().toUpperCase();
        if (cell === startMarker) {
          startIdx = i + 1;
          break;
        }
      }

      if (startIdx === -1) {
        return null;
      }

      let endIdx = rawData.length;
      for (let i = startIdx + 1; i < rawData.length; i++) {
        const cell = String(rawData[i][0] || '').trim().toUpperCase();
        if (stopMarkers.includes(cell)) {
          endIdx = i;
          break;
        }
      }

      // First row of block is date headers
      const headerRow = rawData[startIdx] || [];
      const dateHeaders = [];
      for (let j = 1; j < headerRow.length; j++) {
        const val = String(headerRow[j] || '').trim();
        if (val && (val.includes("Mar") || val.includes("Jun") || val.includes("Sep") || val.includes("Dec") || val.match(/20\d\d/))) {
          dateHeaders.push({ colIndex: j, name: val });
        }
      }

      const rowsByItem = {};
      for (let i = startIdx + 1; i < endIdx; i++) {
        const row = rawData[i];
        if (!row || !row[0]) continue;
        const lineItem = String(row[0]).trim().toLowerCase();
        if (!rowsByItem[lineItem]) {
          rowsByItem[lineItem] = {};
          dateHeaders.forEach(dh => {
            const numStr = String(row[dh.colIndex] || '').replace(/,/g, '').trim();
            rowsByItem[lineItem][dh.name] = parseFloat(numStr) || 0.0;
          });
        }
      }

      return { dateHeaders, rowsByItem };
    }

    const pnl = extractBlock("PROFIT & LOSS", ["QUARTERS", "BALANCE SHEET", "CASH FLOW"]);
    const bs = extractBlock("BALANCE SHEET", ["CASH FLOW", "RATIOS", "DERIVED"]);

    if (!pnl && !bs) {
      throw new Error("Could not detect standard Screener.in financial statement blocks.");
    }

    // Merge date columns
    const allDates = (pnl ? pnl.dateHeaders : bs.dateHeaders).map(d => d.name);
    const helperGet = (block, ...keys) => {
      if (!block) return {};
      for (const k of keys) {
        if (block.rowsByItem[k.toLowerCase()]) {
          return block.rowsByItem[k.toLowerCase()];
        }
      }
      return {};
    };

    const salesMap = helperGet(pnl, "sales", "revenue");
    const netProfMap = helperGet(pnl, "net profit");
    const opProfitMap = helperGet(pnl, "operating profit");
    const pbtMap = helperGet(pnl, "profit before tax");
    const interestMap = helperGet(pnl, "interest");

    const equityCapMap = helperGet(bs, "equity share capital", "share capital");
    const reservesMap = helperGet(bs, "reserves");
    const borrowingsMap = helperGet(bs, "borrowings", "total debt");
    const otherLiabMap = helperGet(bs, "other liabilities");
    const inventoryMap = helperGet(bs, "inventory");
    const tradeRecMap = helperGet(bs, "trade receivables", "receivables");
    const cashMap = helperGet(bs, "cash & bank", "cash");
    const totalAssetsMap = helperGet(bs, "total", "total assets");

    const mergedRecords = allDates.map(dateName => {
      const sales = salesMap[dateName] || 0.0;
      const netProfit = netProfMap[dateName] || 0.0;
      const pbt = pbtMap[dateName] || 0.0;
      const interest = interestMap[dateName] || 0.0;
      const ebit = (pbt + interest) || opProfitMap[dateName] || 0.0;

      const equity = (equityCapMap[dateName] || 0.0) + (reservesMap[dateName] || 0.0);
      const debt = borrowingsMap[dateName] || 0.0;
      const inv = inventoryMap[dateName] || 0.0;
      const rec = tradeRecMap[dateName] || 0.0;
      const cash = cashMap[dateName] || 0.0;
      const currLiab = otherLiabMap[dateName] || 0.0;
      const totAssets = totalAssetsMap[dateName] || (equity + debt);

      return {
        financial_year: dateName,
        revenue: sales,
        cogs: 0.0,
        ebit,
        interest_expense: interest,
        net_profit: netProfit,
        shareholders_equity: equity,
        total_debt: debt,
        current_assets: inv + rec + cash,
        current_liabilities: currLiab,
        inventory: inv,
        trade_receivables: rec,
        cash_bank: cash,
        total_assets: totAssets
      };
    });

    return analyzeFinancialData(mergedRecords);
  }

  // Export to global scope
  window.SA520Engine = {
    baseline3Year,
    infosys10Year,
    analyzeFinancialData,
    parseScreenerWorkbook
  };

})(window);
