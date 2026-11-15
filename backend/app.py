"""
Forensic Audit & SA 520 Backend API Server
Implements the exact logic from Google Colab notebook:
- AP Ledger Forensic Anomaly Engine (TDS, GSTIN, Duplicates, Cut-off)
- SA 520 Analytical Procedures & DuPont ROE Decomposition
- Screener.in Excel Data Sheet Extraction
- Formatted Multi-Tab Excel Export (openpyxl)
"""

import io
import os
import re
import numpy as np
import pandas as pd
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel
from openpyxl.styles import Font, PatternFill

app = FastAPI(
    title="Forensic Audit & SA 520 Intelligence API",
    description="Backend API for Accounts Payable forensic audit, statutory tax compliance, and SA 520 Analytical Procedures.",
    version="1.0.0"
)

# Enable CORS for local or hosted frontends
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

VENDORS = [
    ("Apex Consulting LLP", "37AABCA1234D1Z5", "194J"),
    ("Southern Logistics Ltd", "33AAACS5678E1Z2", "194C"),
    ("Deccan Tech Solutions", "36AABCD9012F1Z8", "194J"),
    ("Metro Facility Services", "29AAACM3456G1Z1", "194C"),
    ("Coastal Engineering Co", "37AAACC7890H1Z9", "194C"),
    ("Vanguard Legal Associates", "27AAACV2345J1Z4", "194J"),
    ("QuickFix Maintenance", "33AAACQ6789K1Z7", "194C")
]
GST_REGEX = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"

class FinancialYearRecord(BaseModel):
    financial_year: str
    revenue: float
    cogs: Optional[float] = 0.0
    ebit: Optional[float] = 0.0
    interest_expense: Optional[float] = 0.0
    net_profit: float
    shareholders_equity: float
    total_debt: Optional[float] = 0.0
    current_assets: Optional[float] = 0.0
    current_liabilities: Optional[float] = 0.0
    inventory: Optional[float] = 0.0
    trade_receivables: Optional[float] = 0.0
    cash_bank: Optional[float] = 0.0
    total_assets: Optional[float] = None

@app.get("/")
def health_check():
    return {
        "status": "online",
        "service": "Forensic Audit & SA 520 Analytical Suite API",
        "endpoints": [
            "/api/ap-audit/generate-mock",
            "/api/ap-audit/analyze",
            "/api/ap-audit/upload",
            "/api/ap-audit/export-excel",
            "/api/sa520/baseline",
            "/api/sa520/analyze",
            "/api/sa520/upload-screener",
            "/api/sa520/export-excel"
        ]
    }

# ==============================================================
# AP LEDGER FORENSIC AUDIT ENDPOINTS (COLAB CELLS 0 - 3)
# ==============================================================

@app.get("/api/ap-audit/generate-mock")
def generate_mock_ledger(rows_count: int = 500, seed: int = 42):
    """Generates synthetic 500-row ERP Ledger with anomalies matching Colab seed(42)."""
    np.random.seed(seed)
    rows = []
    candidate_amts = [18000, 25000, 45000, 75000, 120000, 185000, 250000]

    for i in range(1, rows_count + 1):
        v_name, v_gstin, sec = VENDORS[np.random.randint(0, len(VENDORS))]
        amt = float(np.random.choice(candidate_amts))
        gst = round(amt * 0.18, 2)
        rate = 0.10 if sec == "194J" else 0.02
        tds = round(amt * rate, 2) if amt > 30000 else 0.0

        inv_day = np.random.randint(1, 25)
        inv_date = f"2026-03-{inv_day:02d}"
        grn_date = f"2026-03-{inv_day + 2:02d}"

        rows.append({
            "voucher_no": f"V-{1000 + i}",
            "vendor": v_name,
            "gstin": v_gstin,
            "invoice_no": f"INV-{50000 + i}",
            "invoice_date": inv_date,
            "grn_date": grn_date,
            "section": sec,
            "amount": f"{amt:,.2f}",
            "gst_claimed": f"{gst:,.2f}",
            "tds_deducted": f"{tds:,.2f}"
        })

    df = pd.DataFrame(rows)

    # Injected Anomalies
    if len(df) >= 499:
        df.loc[[14, 189, 412], "vendor"] = None
        df.loc[[22, 105, 310], "gstin"] = "33AAACS5678E1Z"
        df.loc[[67, 245, 455], "gstin"] = "29INVALIDGST999"
        df.loc[[91, 378], "gstin"] = None
        df.loc[341, "invoice_no"] = df.loc[18, "invoice_no"]
        df.loc[488, "invoice_no"] = df.loc[112, "invoice_no"]
        df.loc[488, "vendor"] = "Southern Logistics Ltd"
        df.loc[[9, 74, 215, 399], "tds_deducted"] = "0.00"
        df.loc[[150, 290], "tds_deducted"] = "500.00"

        df.loc[495, ["invoice_date", "grn_date"]] = ["2026-03-30", "2026-04-03"]
        df.loc[496, ["invoice_date", "grn_date"]] = ["2026-03-31", "2026-04-05"]
        df.loc[497, ["invoice_date", "grn_date"]] = ["2026-04-02", "2026-03-28"]
        df.loc[498, ["invoice_date", "grn_date"]] = ["2026-04-04", "2026-03-29"]

    # Fill NaNs for JSON response
    return {"status": "success", "rows": df.where(pd.notnull(df), None).to_dict(orient="records")}

def run_ap_audit_pipeline(df: pd.DataFrame, year_end_str: str = "2026-03-31"):
    """Core audit execution logic matching Colab Step 2 - Step 5."""
    df.columns = df.columns.str.strip().str.lower()

    for col in ["amount", "gst_claimed", "tds_deducted"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "").astype(float)
        else:
            df[col] = 0.0

    df["vendor"] = df["vendor"].fillna("UNKNOWN_VENDOR")
    if "gstin" not in df.columns:
        df["gstin"] = "MISSING"
    else:
        df["gstin"] = df["gstin"].fillna("MISSING").astype(str).str.strip().str.upper()

    # Dates
    if "invoice_date" in df.columns:
        df["invoice_date_dt"] = pd.to_datetime(df["invoice_date"], errors="coerce")
    else:
        df["invoice_date_dt"] = pd.NaT

    if "grn_date" in df.columns:
        df["grn_date_dt"] = pd.to_datetime(df["grn_date"], errors="coerce")
    else:
        df["grn_date_dt"] = pd.NaT

    # TDS Compliance
    rate_map = {"194J": 0.10, "194C": 0.02}
    df["tds_rate"] = df["section"].str.upper().map(rate_map).fillna(0.0)
    df["expected_tds"] = (df["amount"] * df["tds_rate"]).where(df["amount"] > 30000, 0.0)
    df["tds_shortfall"] = (df["expected_tds"] - df["tds_deducted"]).clip(lower=0.0)
    df["disallowance_30pct"] = (df["amount"] * 0.30).where(df["tds_shortfall"] > 0, 0.0)

    # GSTIN check
    df["is_valid_gstin"] = df["gstin"].str.match(GST_REGEX).fillna(False)

    # Cut-Off Check
    year_end = pd.Timestamp(year_end_str)
    early_booking = (df["invoice_date_dt"] <= year_end) & (df["grn_date_dt"] > year_end)
    late_booking = (df["invoice_date_dt"] > year_end) & (df["grn_date_dt"] <= year_end)

    df["cutoff_status"] = "COMPLIANT"
    df.loc[early_booking, "cutoff_status"] = "ERROR: Early Booking (Goods arrived in April)"
    df.loc[late_booking, "cutoff_status"] = "ERROR: Late Booking (Goods arrived in March)"
    df["is_cutoff_error"] = df["cutoff_status"] != "COMPLIANT"

    # Duplicates
    df["is_duplicate"] = df.duplicated(subset=["invoice_no"], keep=False)

    duplicates_df = df[df["is_duplicate"]].sort_values(by="invoice_no")
    tds_shortfall_df = df[df["tds_shortfall"] > 0]
    invalid_gst_df = df[~df["is_valid_gstin"]]
    cutoff_error_df = df[df["is_cutoff_error"]]

    # Vendor summary
    vendor_summary = df.groupby(["vendor", "gstin"]).agg(
        voucher_count=("voucher_no", "count"),
        total_spend=("amount", "sum"),
        total_gst_claimed=("gst_claimed", "sum"),
        total_tds_shortfall=("tds_shortfall", "sum"),
        total_disallowance_30pct=("disallowance_30pct", "sum")
    ).reset_index()

    # Drop temporary datetime objects
    df_clean = df.drop(columns=["invoice_date_dt", "grn_date_dt"], errors="ignore")

    return {
        "kpis": {
            "totalVouchers": int(len(df_clean)),
            "totalSpend": round(float(df_clean["amount"].sum()), 2),
            "totalGstClaimed": round(float(df_clean["gst_claimed"].sum()), 2),
            "totalTdsShortfall": round(float(tds_shortfall_df["tds_shortfall"].sum()), 2),
            "totalDisallowance30": round(float(df_clean["disallowance_30pct"].sum()), 2),
            "duplicateCount": int(len(duplicates_df)),
            "duplicateExposure": round(float(duplicates_df["amount"].sum()), 2),
            "tdsShortfallCount": int(len(tds_shortfall_df)),
            "invalidGstinCount": int(len(invalid_gst_df)),
            "invalidGstExposure": round(float(invalid_gst_df["gst_claimed"].sum()), 2),
            "cutoffErrorCount": int(len(cutoff_error_df)),
            "cutoffExposure": round(float(cutoff_error_df["amount"].sum()), 2),
            "vendorCount": int(len(vendor_summary))
        },
        "exceptions": {
            "duplicates": duplicates_df.to_dict(orient="records"),
            "tdsShortfalls": tds_shortfall_df.to_dict(orient="records"),
            "invalidGstins": invalid_gst_df.to_dict(orient="records"),
            "cutoffErrors": cutoff_error_df.to_dict(orient="records"),
            "vendorSummary": vendor_summary.to_dict(orient="records")
        },
        "fullLedger": df_clean.to_dict(orient="records")
    }

@app.post("/api/ap-audit/analyze")
def analyze_ap_ledger(payload: dict):
    rows = payload.get("rows", [])
    year_end = payload.get("year_end", "2026-03-31")
    if not rows:
        raise HTTPException(status_code=400, detail="Empty rows provided.")
    df = pd.DataFrame(rows)
    return run_ap_audit_pipeline(df, year_end)

@app.post("/api/ap-audit/upload")
async def upload_ap_ledger(file: UploadFile = File(...), year_end: str = Form("2026-03-31")):
    contents = await file.read()
    filename = file.filename.lower()
    if filename.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(contents))
    elif filename.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(contents))
    else:
        raise HTTPException(status_code=400, detail="Unsupported file format. Please upload CSV or Excel.")
    return run_ap_audit_pipeline(df, year_end)

@app.post("/api/ap-audit/export-excel")
def export_ap_excel(payload: dict):
    """Exports formatted 6-tab openpyxl workbook matching Colab Cell 3."""
    rows = payload.get("rows", [])
    year_end = payload.get("year_end", "2026-03-31")
    df = pd.DataFrame(rows)
    result = run_ap_audit_pipeline(df, year_end)

    v_df = pd.DataFrame(result["exceptions"]["vendorSummary"])
    dup_df = pd.DataFrame(result["exceptions"]["duplicates"])
    tds_df = pd.DataFrame(result["exceptions"]["tdsShortfalls"])
    gst_df = pd.DataFrame(result["exceptions"]["invalidGstins"])
    cut_df = pd.DataFrame(result["exceptions"]["cutoffErrors"])
    full_df = pd.DataFrame(result["fullLedger"])

    output = io.BytesIO()
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        v_df.to_excel(writer, sheet_name="1_Vendor_Summary", index=False)
        dup_df.to_excel(writer, sheet_name="2_Duplicate_Invoices", index=False)
        tds_df.to_excel(writer, sheet_name="3_TDS_Shortfalls", index=False)
        gst_df.to_excel(writer, sheet_name="4_Invalid_GSTINs", index=False)
        cut_df.to_excel(writer, sheet_name="5_Cutoff_Errors", index=False)
        full_df.to_excel(writer, sheet_name="6_Full_Clean_Ledger", index=False)

        for sheet_name, ws in writer.sheets.items():
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 5, 12)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = "#,##0.00"

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Master_6Tab_AP_Audit_Report.xlsx"}
    )

# ==============================================================
# SA 520 ANALYTICAL PROCEDURES & DUPONT ROE (COLAB CELLS 4 - 9)
# ==============================================================

@app.get("/api/sa520/baseline")
def get_sa520_baseline():
    """Returns 3-year baseline model and 10-year Infosys dataset from Colab."""
    baseline = [
        {"financial_year": "FY 2023-24", "revenue": 5000000.0, "cogs": 3000000.0, "ebit": 850000.0, "interest_expense": 150000.0, "net_profit": 520000.0, "current_assets": 2200000.0, "inventory": 600000.0, "trade_receivables": 700000.0, "current_liabilities": 1100000.0, "total_debt": 1500000.0, "shareholders_equity": 2500000.0},
        {"financial_year": "FY 2024-25", "revenue": 5600000.0, "cogs": 3400000.0, "ebit": 920000.0, "interest_expense": 180000.0, "net_profit": 550000.0, "current_assets": 2100000.0, "inventory": 750000.0, "trade_receivables": 820000.0, "current_liabilities": 1350000.0, "total_debt": 1800000.0, "shareholders_equity": 2700000.0},
        {"financial_year": "FY 2025-26", "revenue": 6800000.0, "cogs": 3900000.0, "ebit": 1350000.0, "interest_expense": 260000.0, "net_profit": 820000.0, "current_assets": 1950000.0, "inventory": 1100000.0, "trade_receivables": 1250000.0, "current_liabilities": 1650000.0, "total_debt": 2600000.0, "shareholders_equity": 2900000.0}
    ]
    return {"status": "success", "baseline": baseline}

def compute_sa520_ratios(df: pd.DataFrame):
    """Computes DuPont ROE, liquidity, solvency, and automated red-flags."""
    # Ratios
    df["Current_Ratio"] = (df["current_assets"] / df["current_liabilities"]).round(2)
    df["Quick_Ratio"] = ((df["current_assets"] - df["inventory"]) / df["current_liabilities"]).round(2)
    df["Debt_to_Equity"] = (df["total_debt"] / df["shareholders_equity"]).round(2)
    df["Interest_Coverage"] = (df["ebit"] / df["interest_expense"]).round(2)
    df["Gross_Margin_%"] = (((df["revenue"] - df["cogs"]) / df["revenue"]) * 100).round(2)
    df["Net_Margin_%"] = ((df["net_profit"] / df["revenue"]) * 100).round(2)
    df["EBIT_Margin_%"] = ((df["ebit"] / df["revenue"]) * 100).round(2)

    tot_assets = df.get("total_assets", df["current_assets"] + df["shareholders_equity"] + df["total_debt"] - df["current_liabilities"])
    df["Asset_Turnover_x"] = (df["revenue"] / tot_assets).round(2)
    df["Equity_Multiplier_x"] = (tot_assets / df["shareholders_equity"]).round(2)
    df["DuPont_ROE_%"] = (df["Net_Margin_%"] * df["Asset_Turnover_x"] * df["Equity_Multiplier_x"]).round(2)

    df["Inventory_Turnover"] = (df["cogs"] / df["inventory"]).round(2)
    df["Receivables_Days"] = ((df["trade_receivables"] / df["revenue"]) * 365).round(1)
    df["Inventory_Days"] = ((df["inventory"] / df["revenue"]) * 365).round(1)

    # SA 520 Red Flags
    df["Liquidity_Flag"] = np.where(df["Current_Ratio"] < 1.33, "HIGH RISK: Liquidity Crunch", "OK")
    df["Solvency_Flag"] = np.where(df["Debt_to_Equity"] > 0.85, "HIGH RISK: Heavy Debt", "OK")
    df["Collection_Flag"] = np.where(df["Receivables_Days"] > 60, "AUDIT ALERT: Slow Collections / Fictitious Sales Risk", "OK")
    df["Coverage_Flag"] = np.where(df["Interest_Coverage"] < 2.5, "ALERT: Weak Interest Cover", "OK")

    ratio_cols = [
        "EBIT_Margin_%", "Net_Margin_%", "Asset_Turnover_x", "Equity_Multiplier_x",
        "DuPont_ROE_%", "Debt_to_Equity", "Interest_Coverage", "Current_Ratio",
        "Quick_Ratio", "Receivables_Days", "Inventory_Days"
    ]

    scorecard = df.set_index("financial_year")[ratio_cols].T
    if len(df) >= 2:
        last_col = scorecard.columns[-1]
        prev_col = scorecard.columns[-2]
        scorecard["YoY_Change"] = (scorecard[last_col] - scorecard[prev_col]).round(2)
    scorecard = scorecard.reset_index().rename(columns={"index": "Audit_Ratio"})

    return {
        "scorecard": scorecard.to_dict(orient="records"),
        "full_data": df.to_dict(orient="records"),
        "flags": df[["financial_year", "Liquidity_Flag", "Solvency_Flag", "Collection_Flag", "Coverage_Flag"]].to_dict(orient="records")
    }

@app.post("/api/sa520/analyze")
def analyze_sa520(records: List[FinancialYearRecord]):
    df = pd.DataFrame([r.dict() for r in records])
    return compute_sa520_ratios(df)

@app.post("/api/sa520/upload-screener")
async def upload_screener_file(file: UploadFile = File(...)):
    """Parses real Screener.in Excel files matching Colab Cells 5 - 9."""
    contents = await file.read()
    excel_file = pd.ExcelFile(io.BytesIO(contents))
    target_sheet = "Data Sheet" if "Data Sheet" in excel_file.sheet_names else excel_file.sheet_names[-1]
    raw_sheet = pd.read_excel(excel_file, sheet_name=target_sheet, header=None)

    def extract_block(df_raw, start_marker, stop_markers):
        col0 = df_raw[0].astype(str).str.strip().str.upper()
        matches = df_raw[col0 == start_marker].index
        if len(matches) == 0:
            return None
        start_idx = matches[0] + 1
        end_idx = len(df_raw)
        for idx in range(start_idx + 1, len(df_raw)):
            if col0.iloc[idx] in stop_markers:
                end_idx = idx
                break
        block = df_raw.iloc[start_idx:end_idx].dropna(how="all").copy()
        raw_headers = block.iloc[0]
        clean_headers = ["Line_Item"]
        for h in raw_headers.iloc[1:]:
            parsed_dt = pd.to_datetime(h, errors="coerce")
            clean_headers.append(parsed_dt.strftime("%b-%Y") if pd.notna(parsed_dt) else str(h).strip())
        block.columns = clean_headers
        block = block.iloc[1:].copy()
        block["Line_Item"] = block["Line_Item"].astype(str).str.strip().str.lower()
        block = block.drop_duplicates(subset=["Line_Item"], keep="first").set_index("Line_Item")
        valid_cols = [c for c in block.columns if any(y in c for y in ["201", "202"])]
        block = block[valid_cols]
        transposed = block.T.apply(pd.to_numeric, errors="coerce").dropna(how="all")
        transposed.index.name = "financial_year"
        return transposed.reset_index()

    pnl_df = extract_block(raw_sheet, "PROFIT & LOSS", ["QUARTERS", "BALANCE SHEET", "CASH FLOW"])
    bs_df = extract_block(raw_sheet, "BALANCE SHEET", ["CASH FLOW", "RATIOS", "DERIVED"])

    if pnl_df is None or bs_df is None:
        raise HTTPException(status_code=400, detail="Could not parse Screener P&L and Balance Sheet blocks.")

    df = pd.merge(pnl_df, bs_df, on="financial_year", suffixes=("_pnl", "_bs"))

    def get_c(d, *keys, default=0.0):
        for k in keys:
            if k.lower() in d.columns:
                series = d[k.lower()]
                if isinstance(series, pd.DataFrame):
                    series = series.iloc[:, 0]
                return series.fillna(default)
        return pd.Series(default, index=d.index)

    sales = get_c(df, "sales", "revenue").replace(0, np.nan)
    net_prof = get_c(df, "net profit")
    pbt = get_c(df, "profit before tax")
    interest = get_c(df, "interest")
    ebit = pbt + interest

    equity = get_c(df, "equity share capital") + get_c(df, "reserves")
    debt = get_c(df, "borrowings")
    inv = get_c(df, "inventory")
    rec = get_c(df, "receivables")
    cash = get_c(df, "cash & bank")
    curr_liab = get_c(df, "other liabilities", default=np.nan).replace(0, np.nan)
    tot_assets = get_c(df, "total", "total assets")

    parsed_records = []
    for i in range(len(df)):
        parsed_records.append({
            "financial_year": str(df["financial_year"].iloc[i]),
            "revenue": float(sales.iloc[i]) if pd.notna(sales.iloc[i]) else 0.0,
            "cogs": 0.0,
            "ebit": float(ebit.iloc[i]) if pd.notna(ebit.iloc[i]) else 0.0,
            "interest_expense": float(interest.iloc[i]) if pd.notna(interest.iloc[i]) else 0.0,
            "net_profit": float(net_prof.iloc[i]) if pd.notna(net_prof.iloc[i]) else 0.0,
            "shareholders_equity": float(equity.iloc[i]) if pd.notna(equity.iloc[i]) else 0.0,
            "total_debt": float(debt.iloc[i]) if pd.notna(debt.iloc[i]) else 0.0,
            "current_assets": float(inv.iloc[i] + rec.iloc[i] + cash.iloc[i]),
            "current_liabilities": float(curr_liab.iloc[i]) if pd.notna(curr_liab.iloc[i]) else 0.0,
            "inventory": float(inv.iloc[i]),
            "trade_receivables": float(rec.iloc[i]),
            "cash_bank": float(cash.iloc[i]),
            "total_assets": float(tot_assets.iloc[i]) if pd.notna(tot_assets.iloc[i]) else None
        })

    return compute_sa520_ratios(pd.DataFrame(parsed_records))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
