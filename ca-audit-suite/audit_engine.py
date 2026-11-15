"""
Audit Engine Backend Module
Extracted directly from audit_backend.ipynb without altering any accounting or audit formulas.
"""

import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill


def run_ap_audit(input_file_path: str, output_excel_path: str = "Project1_AP_Audit_Report.xlsx") -> dict:
    """
    Backend function for Project 1: Master Accounts Payable Audit.
    Accepts a path to an .xlsx or .csv AP ledger and generates a 5-tab formatted Excel report.
    """
    # 1. Load input file
    if input_file_path.lower().endswith(".csv"):
        df = pd.read_csv(input_file_path)
    else:
        df = pd.read_excel(input_file_path)

    # 2. Sanitize numbers and text fields
    for col in ["amount", "gst_claimed", "tds_deducted"]:
        df[col] = df[col].astype(str).str.replace(",", "").astype(float)

    df["vendor"] = df["vendor"].fillna("UNKNOWN_VENDOR")
    df["gstin"] = df["gstin"].fillna("MISSING").astype(str).str.strip().str.upper()

    # 3. Statutory TDS Rules (Sections 194C & 194J, > ₹30,000 threshold)
    rate_map = {"194J": 0.10, "194C": 0.02}
    df["tds_rate"] = df["section"].map(rate_map).fillna(0.0)
    df["expected_tds"] = (df["amount"] * df["tds_rate"]).where(df["amount"] > 30000, 0.0)
    df["tds_shortfall"] = df["expected_tds"] - df["tds_deducted"]
    df["disallowance_30pct"] = (df["amount"] * 0.30).where(df["tds_shortfall"] > 0, 0.0)

    # 4. 15-Character Indian GSTIN Regex Validation
    gst_pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    df["is_valid_gstin"] = df["gstin"].str.match(gst_pattern)

    # 5. Isolate Exception Tables
    duplicates_df    = df[df.duplicated(subset=["invoice_no"], keep=False)].sort_values(by="invoice_no")
    tds_shortfall_df = df[df["tds_shortfall"] > 0]
    invalid_gst_df   = df[~df["is_valid_gstin"]]

    # 6. Vendor-Level Risk Summary
    vendor_summary = df.groupby(["vendor", "gstin"]).agg(
        voucher_count=("voucher_no", "count"),
        total_spend=("amount", "sum"),
        total_gst_claimed=("gst_claimed", "sum"),
        total_tds_shortfall=("tds_shortfall", "sum"),
        total_disallowance_30pct=("disallowance_30pct", "sum")
    ).reset_index()

    # 7. Export Formatted Excel Report
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        vendor_summary.to_excel(writer, sheet_name="1_Vendor_Summary", index=False)
        duplicates_df.to_excel(writer, sheet_name="2_Duplicate_Invoices", index=False)
        tds_shortfall_df.to_excel(writer, sheet_name="3_TDS_Shortfalls", index=False)
        invalid_gst_df.to_excel(writer, sheet_name="4_Invalid_GSTINs", index=False)
        df.to_excel(writer, sheet_name="5_Full_Clean_Ledger", index=False)

        for _, ws in writer.sheets.items():
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 5
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = "#,##0.00"

    return {
        "summary": {
            "total_vouchers": int(len(df)),
            "duplicate_rows": int(len(duplicates_df)),
            "tds_shortfall_count": int(len(tds_shortfall_df)),
            "total_tds_shortfall_inr": float(tds_shortfall_df["tds_shortfall"].sum()),
            "total_disallowance_30pct_inr": float(tds_shortfall_df["disallowance_30pct"].sum()),
            "invalid_gstin_count": int(len(invalid_gst_df)),
        },
        "tables": {
            "vendor_summary": vendor_summary,
            "duplicates": duplicates_df,
            "tds_shortfalls": tds_shortfall_df,
            "invalid_gstins": invalid_gst_df,
            "full_ledger": df,
        },
        "output_file": output_excel_path,
    }


def run_cutoff_audit(
    input_file_path: str,
    year_end_date: str = "2026-03-31",
    output_excel_path: str = "Project2_Cutoff_Audit_Report.xlsx"
) -> dict:
    """
    Backend function for Project 2: Year-End Cut-Off Boundary Testing.
    Compares 'invoice_date' and 'grn_date' against 'year_end_date'.
    """
    # 1. Load input file
    if input_file_path.lower().endswith(".csv"):
        df = pd.read_csv(input_file_path)
    else:
        df = pd.read_excel(input_file_path)

    if "amount" in df.columns:
        df["amount"] = df["amount"].astype(str).str.replace(",", "").astype(float)

    # 2. Parse calendar dates and boundary timestamp
    df["invoice_date"] = pd.to_datetime(df["invoice_date"])
    df["grn_date"]     = pd.to_datetime(df["grn_date"])
    boundary           = pd.Timestamp(year_end_date)

    # 3. Cut-Off Boundary Logic
    early_booking = (df["invoice_date"] <= boundary) & (df["grn_date"] > boundary)
    late_booking  = (df["invoice_date"] > boundary) & (df["grn_date"] <= boundary)

    df["cutoff_status"] = "COMPLIANT"
    df.loc[early_booking, "cutoff_status"] = "ERROR: Early Booking (Goods arrived after year-end)"
    df.loc[late_booking,  "cutoff_status"] = "ERROR: Late Booking (Goods arrived before year-end)"

    # 4. Isolate Cut-Off Exceptions
    cutoff_exceptions = df[df["cutoff_status"] != "COMPLIANT"].copy()

    # 5. Export Formatted Excel Report
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    alert_fill  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    alert_font  = Font(color="9C0006", bold=True)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        cutoff_exceptions.to_excel(writer, sheet_name="1_Cutoff_Exceptions", index=False)
        df.to_excel(writer, sheet_name="2_All_Tested_Vouchers", index=False)

        for _, ws in writer.sheets.items():
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 5
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = "#,##0.00"
                    elif isinstance(cell.value, str) and "ERROR:" in cell.value:
                        cell.fill = alert_fill
                        cell.font = alert_font

    total_exception_amt = float(cutoff_exceptions["amount"].sum()) if "amount" in cutoff_exceptions.columns else 0.0

    return {
        "summary": {
            "year_end_boundary": year_end_date,
            "total_vouchers_tested": int(len(df)),
            "cutoff_error_count": int(len(cutoff_exceptions)),
            "total_exception_amount_inr": total_exception_amt,
        },
        "tables": {
            "cutoff_exceptions": cutoff_exceptions,
            "all_vouchers": df,
        },
        "output_file": output_excel_path,
    }


def run_screener_ratio_audit(
    screener_excel_path: str,
    output_excel_path: str = "Project3_SA520_Ratio_Audit.xlsx"
) -> dict:
    """
    Backend function for Project 3: 10-Year Financial Statement & SA 520 Ratio Audit.
    Parses a Screener.in 'Export to Excel' file and generates a 3-tab Excel scorecard.
    """
    excel_file = pd.ExcelFile(screener_excel_path)
    target_sheet = "Data Sheet" if "Data Sheet" in excel_file.sheet_names else excel_file.sheet_names[-1]
    raw_sheet = pd.read_excel(excel_file, sheet_name=target_sheet, header=None)

    def extract_screener_block(df_raw: pd.DataFrame, start_marker: str, stop_markers: list) -> pd.DataFrame:
        col0 = df_raw[0].astype(str).str.strip().str.upper()
        matches = df_raw[col0 == start_marker].index
        if len(matches) == 0:
            raise ValueError(f"Could not find '{start_marker}' in sheet '{target_sheet}'.")

        start_idx = matches[0] + 1
        end_idx = len(df_raw)
        for idx in range(start_idx + 1, len(df_raw)):
            if col0.iloc[idx] in stop_markers:
                end_idx = idx
                break

        block = df_raw.iloc[start_idx:end_idx].dropna(how="all").copy()

        # Format Excel timestamps cleanly into 'Mar-2026'
        raw_headers = block.iloc[0]
        clean_headers = ["Line_Item"]
        for h in raw_headers.iloc[1:]:
            parsed_dt = pd.to_datetime(h, errors="coerce")
            if pd.notna(parsed_dt):
                clean_headers.append(parsed_dt.strftime("%b-%Y"))
            else:
                clean_headers.append(str(h).strip())

        block.columns = clean_headers
        block = block.iloc[1:].copy()
        block["Line_Item"] = block["Line_Item"].astype(str).str.strip().str.lower()

        # Deduplicate 'Total' row so Balance Sheet Total Liabilities doesn't collide
        block = block.drop_duplicates(subset=["Line_Item"], keep="first").set_index("Line_Item")

        valid_cols = [c for c in block.columns if any(y in c for y in ["201", "202"])]
        block = block[valid_cols]

        transposed = block.T.apply(pd.to_numeric, errors="coerce").dropna(how="all")
        transposed.index.name = "Financial_Year"
        return transposed.reset_index()

    pnl_df = extract_screener_block(raw_sheet, "PROFIT & LOSS", ["QUARTERS", "BALANCE SHEET", "CASH FLOW"])
    bs_df  = extract_screener_block(raw_sheet, "BALANCE SHEET", ["CASH FLOW", "RATIOS", "DERIVED"])

    df = pd.merge(pnl_df, bs_df, on="Financial_Year", suffixes=("_pnl", "_bs"))

    def get_col(dataframe: pd.DataFrame, *possible_names: str, default: float = 0.0) -> pd.Series:
        for name in possible_names:
            if name.lower() in dataframe.columns:
                col_data = dataframe[name.lower()]
                if isinstance(col_data, pd.DataFrame):
                    col_data = col_data.iloc[:, 0]
                return col_data.fillna(default)
        return pd.Series(default, index=dataframe.index)

    # Extract core ledger lines and derive EBIT = Profit Before Tax + Interest
    sales         = get_col(df, "sales", "revenue").replace(0, np.nan)
    net_prof      = get_col(df, "net profit")
    pbt           = get_col(df, "profit before tax")
    interest      = get_col(df, "interest")
    ebit          = pbt + interest
    interest_safe = interest.replace(0, np.nan)

    equity     = (get_col(df, "equity share capital") + get_col(df, "reserves")).replace(0, np.nan)
    debt       = get_col(df, "borrowings")
    inv        = get_col(df, "inventory")
    rec        = get_col(df, "receivables")
    cash       = get_col(df, "cash & bank")
    curr_liab  = get_col(df, "other liabilities", default=np.nan).replace(0, np.nan)
    tot_assets = get_col(df, "total", "total assets").replace(0, np.nan)

    # Compute 10-Year Audit Ratios + 3-Step DuPont ROE Decomposition
    df["EBIT_Margin_%"]        = ((ebit / sales) * 100).round(2)
    df["Net_Margin_%"]         = ((net_prof / sales) * 100).round(2)
    df["Asset_Turnover_x"]     = (sales / tot_assets).round(2)
    df["Equity_Multiplier_x"]  = (tot_assets / equity).round(2)
    df["DuPont_ROE_%"]         = ((net_prof / sales) * (sales / tot_assets) * (tot_assets / equity) * 100).round(2)
    df["Debt_to_Equity"]       = (debt / equity).round(2)
    df["Interest_Coverage"]    = (ebit / interest_safe).round(2)
    df["Approx_Current_Ratio"] = ((inv + rec + cash) / curr_liab).round(2)
    df["Receivables_Days"]     = ((rec / sales) * 365).round(1)
    df["Inventory_Days"]       = ((inv / sales) * 365).round(1)

    # SA 520 Automated Red-Flag Alerts
    df["Solvency_Flag"]   = np.where(df["Debt_to_Equity"] > 1.0, "ALERT: High Leverage (>1.0x)", "OK")
    df["Collection_Flag"] = np.where(df["Receivables_Days"] > 75, "ALERT: Slow Collections (>75 Days)", "OK")
    df["Coverage_Flag"]   = np.where(df["Interest_Coverage"] < 2.5, "ALERT: Weak Interest Cover", "OK")

    ratio_cols = [
        "EBIT_Margin_%", "Net_Margin_%", "Asset_Turnover_x",
        "Equity_Multiplier_x", "DuPont_ROE_%", "Debt_to_Equity",
        "Interest_Coverage", "Approx_Current_Ratio",
        "Receivables_Days", "Inventory_Days"
    ]

    scorecard = df.set_index("Financial_Year")[ratio_cols].T.reset_index().rename(columns={"index": "Audit_Ratio"})
    flags_df  = df[["Financial_Year", "Solvency_Flag", "Collection_Flag", "Coverage_Flag"]]

    # Export Formatted Excel Report
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    alert_fill  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    alert_font  = Font(color="9C0006", bold=True)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        scorecard.to_excel(writer, sheet_name="1_10Yr_Ratio_Scorecard", index=False)
        flags_df.to_excel(writer, sheet_name="2_SA520_Risk_Alerts", index=False)
        df.to_excel(writer, sheet_name="3_Full_Extracted_Data", index=False)

        for _, ws in writer.sheets.items():
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 5
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = "#,##0.00"
                    elif isinstance(cell.value, str) and "ALERT" in cell.value:
                        cell.fill = alert_fill
                        cell.font = alert_font

    return {
        "summary": {
            "years_analyzed": list(df["Financial_Year"]),
            "latest_year": str(df["Financial_Year"].iloc[-1]),
            "latest_dupont_roe_pct": float(df["DuPont_ROE_%"].iloc[-1]),
            "latest_debt_to_equity": float(df["Debt_to_Equity"].iloc[-1]),
        },
        "tables": {
            "scorecard": scorecard,
            "risk_alerts": flags_df,
            "full_data": df,
        },
        "output_file": output_excel_path,
    }
