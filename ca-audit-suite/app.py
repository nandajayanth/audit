"""
CA Audit Suite - Executive Forensic & Statutory Audit Intelligence Platform
Dual-mode: Streamlit Interactive UI + Vercel Services ASGI Application
"""

import os
import io
import sys
import tempfile
import datetime
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st
import requests
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

# Bound internal service URL (injected by Vercel Services via service bindings)
BACKEND_URL = os.environ.get("BACKEND_URL", "").strip().rstrip("/")

# Import the core audit engine functions
from audit_engine import run_ap_audit, run_cutoff_audit, run_screener_ratio_audit

# =============================================================================
# VERCEL SERVICES ASGI APPLICATION (entrypoint: app:app)
# =============================================================================
app = FastAPI(
    title="CA Audit Suite Service",
    description="Executive Forensic Audit and Financial Intelligence Platform",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/health")
def health():
    return {
        "service": "ca-audit-suite",
        "status": "online",
        "backend_binding": BACKEND_URL or "not_configured"
    }

@app.get("/", response_class=HTMLResponse)
def index():
    return """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>CA Forensic Audit Suite</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0F172A; color: #F8FAFC; margin: 0; padding: 2rem; }
    .card { background: #1E293B; border-radius: 12px; padding: 2rem; max-width: 800px; margin: 2rem auto; border-top: 4px solid #1F4E78; box-shadow: 0 10px 25px rgba(0,0,0,0.3); }
    h1 { color: #FFFFFF; font-size: 1.8rem; margin-top: 0; }
    p { color: #94A3B8; line-height: 1.6; }
    .badge { display: inline-block; padding: 4px 12px; border-radius: 999px; background: #1F4E78; color: #fff; font-size: 0.8rem; font-weight: 600; }
    ul { line-height: 1.8; color: #E2E8F0; }
    code { background: #0F172A; padding: 2px 6px; border-radius: 4px; color: #38BDF8; font-size: 0.9em; }
  </style>
</head>
<body>
  <div class="card">
    <span class="badge">Vercel Service Active</span>
    <h1>CA Forensic Audit and Financial Intelligence Suite</h1>
    <p>Automated Accounts Payable Forensic Audit, Statutory Tax Compliance, Year-End Cut-Off Testing and SA 520 DuPont ROE Analytical Procedures.</p>
    <hr style="border: 0; border-top: 1px solid #334155; margin: 1.5rem 0;" />
    <h3>Available Service Endpoints:</h3>
    <ul>
      <li><code>GET /api/health</code> - Service Health and Internal Backend Binding Status</li>
      <li><code>POST /api/audit/ap</code> - Run AP Forensic Audit Pipeline</li>
      <li><code>POST /api/audit/cutoff</code> - Run Year-End Cut-Off Boundary Audit</li>
      <li><code>POST /api/audit/screener</code> - Run 10-Year SA 520 DuPont ROE Audit</li>
    </ul>
  </div>
</body>
</html>"""

@app.post("/api/audit/ap")
async def audit_ap(file: UploadFile = File(...)):
    suffix = ".csv" if (file.filename and file.filename.endswith(".csv")) else ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_in = tmp.name
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        tmp_out_path = tmp_out.name
    try:
        res = run_ap_audit(tmp_in, tmp_out_path)
        with open(tmp_out_path, "rb") as f:
            excel_bytes = f.read()
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=AP_Audit_Report.xlsx"}
        )
    finally:
        for p in [tmp_in, tmp_out_path]:
            if os.path.exists(p):
                os.remove(p)

@app.post("/api/audit/cutoff")
async def audit_cutoff(file: UploadFile = File(...), year_end: str = Form("2026-03-31")):
    suffix = ".csv" if (file.filename and file.filename.endswith(".csv")) else ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_in = tmp.name
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        tmp_out_path = tmp_out.name
    try:
        res = run_cutoff_audit(tmp_in, year_end_date=year_end, output_excel_path=tmp_out_path)
        with open(tmp_out_path, "rb") as f:
            excel_bytes = f.read()
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Cutoff_Audit_Report.xlsx"}
        )
    finally:
        for p in [tmp_in, tmp_out_path]:
            if os.path.exists(p):
                os.remove(p)

@app.post("/api/audit/screener")
async def audit_screener(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        tmp_in = tmp.name
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        tmp_out_path = tmp_out.name
    try:
        res = run_screener_ratio_audit(tmp_in, output_excel_path=tmp_out_path)
        with open(tmp_out_path, "rb") as f:
            excel_bytes = f.read()
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=SA520_Audit_Report.xlsx"}
        )
    finally:
        for p in [tmp_in, tmp_out_path]:
            if os.path.exists(p):
                os.remove(p)

# =============================================================================
# STREAMLIT APPLICATION RUNNER
# =============================================================================
def run_streamlit_app():
    st.set_page_config(
        page_title="CA Forensic Audit & Financial Intelligence Suite",
        page_icon="⚖️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # -----------------------------------------------------------------------------
    # EXECUTIVE THEME CUSTOM CSS (Navy Blue #1F4E78 Theme)
    # -----------------------------------------------------------------------------
    st.markdown("""
    <style>
        /* Global & Typography */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        /* Executive Top Banner */
        .executive-header {
            background: linear-gradient(135deg, #1F4E78 0%, #112d46 100%);
            padding: 1.5rem 2rem;
            border-radius: 10px;
            color: #FFFFFF;
            margin-bottom: 1.5rem;
            box-shadow: 0 4px 15px rgba(31, 78, 120, 0.2);
        }
        .executive-header h1 {
            color: #FFFFFF !important;
            font-size: 1.9rem;
            font-weight: 700;
            margin: 0 0 0.4rem 0;
            letter-spacing: -0.5px;
        }
        .executive-header p {
            color: #D2E3F4 !important;
            font-size: 0.95rem;
            margin: 0;
            font-weight: 400;
        }
        .badge-tag {
            display: inline-block;
            background-color: rgba(255, 255, 255, 0.15);
            color: #FFFFFF;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            margin-left: 10px;
            letter-spacing: 0.5px;
        }

        /* Metric Cards */
        .metric-card {
            background: #FFFFFF;
            border: 1px solid #E1E8EE;
            border-top: 4px solid #1F4E78;
            border-radius: 8px;
            padding: 1rem 1.2rem;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
            margin-bottom: 1rem;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 12px rgba(31, 78, 120, 0.12);
        }
        .metric-card.alert-card {
            border-top-color: #D9534F;
        }
        .metric-card.warning-card {
            border-top-color: #F0AD4E;
        }
        .metric-card.success-card {
            border-top-color: #2E7D32;
        }
        .metric-label {
            font-size: 0.8rem;
            font-weight: 600;
            text-transform: uppercase;
            color: #5A6A7E;
            letter-spacing: 0.5px;
            margin-bottom: 0.3rem;
        }
        .metric-value {
            font-size: 1.6rem;
            font-weight: 700;
            color: #1F4E78;
            line-height: 1.2;
        }
        .metric-value.danger {
            color: #C0392B;
        }
        .metric-value.warning {
            color: #D35400;
        }
        .metric-value.success {
            color: #2E7D32;
        }
        .metric-sub {
            font-size: 0.75rem;
            color: #7F8C8D;
            margin-top: 0.2rem;
        }

        /* Section Subheadings */
        .section-title {
            color: #1F4E78;
            font-size: 1.2rem;
            font-weight: 700;
            margin-top: 1.2rem;
            margin-bottom: 0.8rem;
            padding-bottom: 0.3rem;
            border-bottom: 2px solid #EBF1F5;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        /* Sidebar Styling */
        .sidebar-brand {
            padding: 10px 0;
            border-bottom: 1px solid #E1E8EE;
            margin-bottom: 15px;
        }
        .sidebar-brand h3 {
            color: #1F4E78;
            font-size: 1.15rem;
            font-weight: 700;
            margin: 0;
        }
        .sidebar-brand p {
            color: #6C7A89;
            font-size: 0.8rem;
            margin: 0;
        }

        /* Streamlit Tab Customization */
        button[data-baseweb="tab"] {
            font-weight: 600;
            font-size: 0.95rem;
            color: #4A5568;
            padding: 10px 20px;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            color: #1F4E78 !important;
            border-bottom-color: #1F4E78 !important;
        }

        /* Primary buttons */
        .stDownloadButton > button {
            background-color: #1F4E78 !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            padding: 0.5rem 1.2rem !important;
            transition: background-color 0.2s ease;
        }
        .stDownloadButton > button:hover {
            background-color: #163857 !important;
            color: #FFFFFF !important;
        }
    </style>
    """, unsafe_allow_html=True)

    # -----------------------------------------------------------------------------
    # SAMPLE DATA DIRECTORY & PREPARATION
    # -----------------------------------------------------------------------------
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")

    def get_sample_file_bytes(filename: str) -> bytes:
        filepath = os.path.join(SAMPLE_DIR, filename)
        if os.path.exists(filepath):
            with open(filepath, "rb") as f:
                return f.read()
        return b""

    # -----------------------------------------------------------------------------
    # SIDEBAR
    # -----------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-brand">
            <h3>⚖️ CA Audit Suite</h3>
            <p>Statutory & Forensic Audit Suite v2.0</p>
        </div>
        """, unsafe_allow_html=True)

        st.info("💡 **Executive Audit Workbench**\n\nImplements real-time forensic ledger anomaly testing, Section 194C/J TDS audits, GSTIN verification, Year-End cut-off boundary checks, and 10-year SA 520 DuPont ROE analytics.")

        if BACKEND_URL:
            try:
                resp = requests.get(f"{BACKEND_URL}/", timeout=2)
                if resp.status_code == 200:
                    backend_info = resp.json()
                    st.success(f"🔗 **Internal Service Connected**\n\n`backend` service bound at `{BACKEND_URL}` ({backend_info.get('status', 'online')})")
                else:
                    st.info(f"🔗 **Internal Service Bound**: `{BACKEND_URL}`")
            except Exception:
                st.info(f"🔗 **Internal Service Bound**: `{BACKEND_URL}`")

        st.markdown("### 📥 Sample Test Files")
        st.markdown("Download realistic sample files to test all audit modules:")

        sample_ap_bytes = get_sample_file_bytes("sample_ap_ledger.csv")
        if sample_ap_bytes:
            st.download_button(
                label="📄 Sample AP Ledger (CSV)",
                data=sample_ap_bytes,
                file_name="Sample_AP_500_Rows_Ledger.csv",
                mime="text/csv",
                help="500-row ERP Accounts Payable ledger with synthetic anomalies: duplicate invoices, broken GSTINs, TDS shortfalls, and missing vendors."
            )

        sample_cutoff_bytes = get_sample_file_bytes("sample_cutoff_ledger.csv")
        if sample_cutoff_bytes:
            st.download_button(
                label="📅 Sample Cut-Off Ledger (CSV)",
                data=sample_cutoff_bytes,
                file_name="Sample_Cutoff_Ledger.csv",
                mime="text/csv",
                help="Sample vouchers around March 31 with early & late booking delivery exceptions."
            )

        sample_screener_bytes = get_sample_file_bytes("sample_screener_infosys.xlsx")
        if sample_screener_bytes:
            st.download_button(
                label="📊 Sample Screener.in (XLSX)",
                data=sample_screener_bytes,
                file_name="Sample_Screener_Infosys.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="10-year audited financial statement data from Screener.in for SA 520 ratio analysis."
            )

        st.markdown("---")
        st.markdown("### 📋 Regulatory References")
        st.markdown("""
        - **Income Tax Act, 1961**:
          - Section 194C (Contractors: 2%)
          - Section 194J (Professionals: 10%)
          - Section 40(a)(ia) (30% Disallowance)
        - **CGST Act, 2017**: Section 16(2) ITC
        - **ICAI Standards**:
          - SA 520: Analytical Procedures
          - SA 500: Audit Evidence
        """)
        st.markdown("<p style='font-size:0.75rem; color:#888; text-align:center;'>Prepared for Chartered Accountants & Audit Teams</p>", unsafe_allow_html=True)

    # -----------------------------------------------------------------------------
    # MAIN HEADER BANNER
    # -----------------------------------------------------------------------------
    st.markdown("""
    <div class="executive-header">
        <h1>CA Audit Suite <span class="badge-tag">Executive Edition</span></h1>
        <p>Automated Accounts Payable Forensic Audit, Statutory Tax Compliance, Cut-Off Testing & SA 520 Analytical Procedures</p>
    </div>
    """, unsafe_allow_html=True)

    # -----------------------------------------------------------------------------
    # 3-TAB INTERFACE
    # -----------------------------------------------------------------------------
    tab1, tab2, tab3 = st.tabs([
        "📑 Tab 1: Master AP Audit",
        "⏳ Tab 2: Year-End Cut-Off Testing",
        "📈 Tab 3: 10-Year SA 520 & DuPont ROE"
    ])

    # =============================================================================
    # TAB 1: ACCOUNTS PAYABLE FORENSIC AUDIT
    # =============================================================================
    with tab1:
        st.markdown("""
        <div class="section-title">
            <span>📑</span> Accounts Payable Ledger Anomaly & Statutory Compliance Engine
        </div>
        <p style='color:#5A6A7E; font-size:0.9rem;'>
            Upload an ERP Accounts Payable ledger (.xlsx or .csv) to execute automated forensic checks: 
            Duplicate Invoices, Statutory TDS Shortfalls (Sec 194C & 194J with 30% Disallowance u/s 40(a)(ia)), and 15-character GSTIN regex validation.
        </p>
        """, unsafe_allow_html=True)

        col_upload, col_quick = st.columns([3, 1])
        with col_upload:
            ap_file = st.file_uploader(
                "Upload AP Ledger (.xlsx or .csv)",
                type=["xlsx", "csv"],
                key="ap_uploader",
                help="Upload ledger containing voucher_no, vendor, gstin, invoice_no, section, amount, gst_claimed, tds_deducted"
            )
        with col_quick:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            use_sample_ap = st.button("⚡ Test with Sample AP Ledger", key="btn_sample_ap")

        # Determine data source
        active_ap_source = None
        if ap_file is not None:
            active_ap_source = ("uploaded", ap_file)
        elif use_sample_ap:
            active_ap_source = ("sample", os.path.join(SAMPLE_DIR, "sample_ap_ledger.csv"))

        if active_ap_source:
            with st.spinner("Executing forensic audit analysis and generating 5-tab Excel report..."):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    # Save input file to disk
                    if active_ap_source[0] == "uploaded":
                        suffix = ".csv" if active_ap_source[1].name.lower().endswith(".csv") else ".xlsx"
                        input_path = os.path.join(tmp_dir, f"input_ledger{suffix}")
                        with open(input_path, "wb") as f:
                            f.write(active_ap_source[1].getvalue())
                    else:
                        input_path = active_ap_source[1]

                    output_excel_path = os.path.join(tmp_dir, "Project1_AP_Audit_Report.xlsx")

                    try:
                        result = run_ap_audit(input_path, output_excel_path)
                        summary = result["summary"]
                        tables = result["tables"]

                        # Read generated excel report for download
                        with open(output_excel_path, "rb") as f:
                            report_bytes = f.read()

                        st.success("✅ AP Audit completed successfully!")

                        # --- SUMMARY KPIS ---
                        k1, k2, k3, k4, k5, k6 = st.columns(6)
                        with k1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Total Vouchers</div>
                                <div class="metric-value">{summary['total_vouchers']:,}</div>
                                <div class="metric-sub">Ledger Population</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with k2:
                            st.markdown(f"""
                            <div class="metric-card alert-card">
                                <div class="metric-label">Duplicate Rows</div>
                                <div class="metric-value danger">{summary['duplicate_rows']}</div>
                                <div class="metric-sub">Potential Double Bills</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with k3:
                            st.markdown(f"""
                            <div class="metric-card alert-card">
                                <div class="metric-label">TDS Shortfalls</div>
                                <div class="metric-value danger">{summary['tds_shortfall_count']}</div>
                                <div class="metric-sub">Under/Zero Deducted</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with k4:
                            st.markdown(f"""
                            <div class="metric-card warning-card">
                                <div class="metric-label">TDS Shortfall (₹)</div>
                                <div class="metric-value warning">₹{summary['total_tds_shortfall_inr']:,.2f}</div>
                                <div class="metric-sub">Statutory Tax Deficit</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with k5:
                            st.markdown(f"""
                            <div class="metric-card alert-card">
                                <div class="metric-label">30% Disallowance</div>
                                <div class="metric-value danger">₹{summary['total_disallowance_30pct_inr']:,.2f}</div>
                                <div class="metric-sub">u/s 40(a)(ia) Risk</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with k6:
                            st.markdown(f"""
                            <div class="metric-card alert-card">
                                <div class="metric-label">Invalid GSTINs</div>
                                <div class="metric-value danger">{summary['invalid_gstin_count']}</div>
                                <div class="metric-sub">ITC Ineligibility</div>
                            </div>
                            """, unsafe_allow_html=True)

                        # --- DOWNLOAD EXCEL REPORT BUTTON ---
                        st.download_button(
                            label="📥 Download Complete AP Audit Excel Report (5 Tabs)",
                            data=report_bytes,
                            file_name="Project1_AP_Audit_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_ap_report"
                        )

                        st.markdown("<div style='margin-top:1.5rem;'></div>", unsafe_allow_html=True)

                        # --- EXCEPTION & ANALYSIS SUB-TABS ---
                        ap_sub1, ap_sub2, ap_sub3, ap_sub4, ap_sub5 = st.tabs([
                            f"📊 Vendor Summary ({len(tables['vendor_summary'])})",
                            f"⚠️ Duplicate Invoices ({len(tables['duplicates'])})",
                            f"🚨 TDS Shortfalls ({len(tables['tds_shortfalls'])})",
                            f"❌ Invalid GSTINs ({len(tables['invalid_gstins'])})",
                            f"📋 Full Clean Ledger ({len(tables['full_ledger'])})"
                        ])

                        with ap_sub1:
                            st.markdown("**Vendor-Level Risk Summary** (Aggregated Spends, GST Claimed, TDS Shortfalls, and Tax Disallowances)")
                            st.dataframe(
                                tables['vendor_summary'].style.format({
                                    "total_spend": "₹{:,.2f}",
                                    "total_gst_claimed": "₹{:,.2f}",
                                    "total_tds_shortfall": "₹{:,.2f}",
                                    "total_disallowance_30pct": "₹{:,.2f}"
                                }),
                                use_container_width=True
                            )

                        with ap_sub2:
                            st.markdown("**Duplicate Invoices Detected** (Potential Double-Payment Risks matching on Invoice Number)")
                            if len(tables['duplicates']) > 0:
                                st.dataframe(
                                    tables['duplicates'].style.format({
                                        "amount": "₹{:,.2f}",
                                        "gst_claimed": "₹{:,.2f}",
                                        "tds_deducted": "₹{:,.2f}"
                                    }),
                                    use_container_width=True
                                )
                            else:
                                st.info("No duplicate invoice numbers found in this ledger.")

                        with ap_sub3:
                            st.markdown("**TDS Statutory Shortfall Register** (Bills > ₹30,000 evaluated under Sections 194C [2%] and 194J [10%])")
                            if len(tables['tds_shortfalls']) > 0:
                                st.dataframe(
                                    tables['tds_shortfalls'].style.format({
                                        "amount": "₹{:,.2f}",
                                        "tds_rate": "{:.0%}",
                                        "expected_tds": "₹{:,.2f}",
                                        "tds_deducted": "₹{:,.2f}",
                                        "tds_shortfall": "₹{:,.2f}",
                                        "disallowance_30pct": "₹{:,.2f}"
                                    }),
                                    use_container_width=True
                                )
                            else:
                                st.info("No TDS shortfalls identified. All tested vouchers are fully compliant!")

                        with ap_sub4:
                            st.markdown("**Invalid or Corrupted GSTINs** (Failed 15-character statutory GST regex `^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$`)")
                            if len(tables['invalid_gstins']) > 0:
                                st.dataframe(
                                    tables['invalid_gstins'].style.format({
                                        "amount": "₹{:,.2f}",
                                        "gst_claimed": "₹{:,.2f}"
                                    }),
                                    use_container_width=True
                                )
                            else:
                                st.info("All supplier GSTINs strictly conform to the 15-character GST checksum format.")

                        with ap_sub5:
                            st.markdown("**Full Clean Ledger** (with appended audit check flags and calculations)")
                            st.dataframe(tables['full_ledger'], use_container_width=True)

                    except Exception as e:
                        st.error(f"Error during AP Audit analysis: {str(e)}")
                        st.exception(e)
        else:
            st.info("👆 Please upload your AP ledger (.xlsx or .csv) or click **'Test with Sample AP Ledger'** above to view the analysis.")


    # =============================================================================
    # TAB 2: YEAR-END CUT-OFF TESTING
    # =============================================================================
    with tab2:
        st.markdown("""
        <div class="section-title">
            <span>⏳</span> Year-End Cut-Off Boundary Testing (Invoicing vs Physical GRN Receipt)
        </div>
        <p style='color:#5A6A7E; font-size:0.9rem;'>
            Tests expense completeness and period boundary recognition by comparing <b>invoice_date</b> and <b>grn_date</b> 
            against the statutory year-end boundary date (default: 31-Mar-2026). Detects Early Bookings (goods received post year-end) 
            and Late Bookings (goods received pre year-end but invoiced later).
        </p>
        """, unsafe_allow_html=True)

        col_cut_upload, col_cut_date, col_cut_quick = st.columns([2, 1, 1])
        with col_cut_upload:
            cutoff_file = st.file_uploader(
                "Upload Cut-Off Ledger (.xlsx or .csv)",
                type=["xlsx", "csv"],
                key="cutoff_uploader",
                help="Upload file containing invoice_date, grn_date, voucher_no, vendor, amount"
            )
        with col_cut_date:
            selected_boundary_date = st.date_input(
                "Year-End Boundary Date",
                value=datetime.date(2026, 3, 31),
                help="Statutory financial year-end boundary date (default: 2026-03-31)"
            )
        with col_cut_quick:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            use_sample_cutoff = st.button("⚡ Test with Sample Cut-Off", key="btn_sample_cutoff")

        active_cutoff_source = None
        if cutoff_file is not None:
            active_cutoff_source = ("uploaded", cutoff_file)
        elif use_sample_cutoff:
            active_cutoff_source = ("sample", os.path.join(SAMPLE_DIR, "sample_cutoff_ledger.csv"))

        if active_cutoff_source:
            year_end_str = selected_boundary_date.strftime("%Y-%m-%d")
            with st.spinner("Executing cut-off boundary evaluation and generating formatted report..."):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    if active_cutoff_source[0] == "uploaded":
                        suffix = ".csv" if active_cutoff_source[1].name.lower().endswith(".csv") else ".xlsx"
                        input_path = os.path.join(tmp_dir, f"cutoff_input{suffix}")
                        with open(input_path, "wb") as f:
                            f.write(active_cutoff_source[1].getvalue())
                    else:
                        input_path = active_cutoff_source[1]

                    output_excel_path = os.path.join(tmp_dir, "Project2_Cutoff_Audit_Report.xlsx")

                    try:
                        result = run_cutoff_audit(input_path, year_end_date=year_end_str, output_excel_path=output_excel_path)
                        summary = result["summary"]
                        tables = result["tables"]

                        with open(output_excel_path, "rb") as f:
                            report_bytes = f.read()

                        st.success("✅ Cut-Off Audit completed successfully!")

                        # --- SUMMARY KPIS ---
                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Year-End Boundary</div>
                                <div class="metric-value">{summary['year_end_boundary']}</div>
                                <div class="metric-sub">Closing Brick Wall</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Total Vouchers Tested</div>
                                <div class="metric-value">{summary['total_vouchers_tested']:,}</div>
                                <div class="metric-sub">Tested Sample</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c3:
                            st.markdown(f"""
                            <div class="metric-card alert-card">
                                <div class="metric-label">Cut-Off Errors</div>
                                <div class="metric-value danger">{summary['cutoff_error_count']}</div>
                                <div class="metric-sub">Boundary Deviations</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with c4:
                            st.markdown(f"""
                            <div class="metric-card warning-card">
                                <div class="metric-label">Exception Amount (₹)</div>
                                <div class="metric-value warning">₹{summary['total_exception_amount_inr']:,.2f}</div>
                                <div class="metric-sub">Misstated Period Spend</div>
                            </div>
                            """, unsafe_allow_html=True)

                        # --- DOWNLOAD EXCEL REPORT BUTTON ---
                        st.download_button(
                            label="📥 Download Cut-Off Audit Excel Report (2 Tabs)",
                            data=report_bytes,
                            file_name="Project2_Cutoff_Audit_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_cutoff_report"
                        )

                        st.markdown("<div style='margin-top:1.5rem;'></div>", unsafe_allow_html=True)

                        # --- TABLES ---
                        cut_sub1, cut_sub2 = st.tabs([
                            f"🚨 Cut-Off Exceptions ({len(tables['cutoff_exceptions'])})",
                            f"📋 All Tested Vouchers ({len(tables['all_vouchers'])})"
                        ])

                        with cut_sub1:
                            st.markdown("**Cut-Off Misstatement Exceptions** (Early Booking & Late Booking errors)")
                            if len(tables['cutoff_exceptions']) > 0:
                                exceptions_display = tables['cutoff_exceptions'].copy()
                                # Format date columns cleanly
                                for d_col in ["invoice_date", "grn_date"]:
                                    if d_col in exceptions_display.columns:
                                        exceptions_display[d_col] = pd.to_datetime(exceptions_display[d_col]).dt.strftime("%d-%b-%Y")

                                st.dataframe(
                                    exceptions_display.style.format({
                                        "amount": "₹{:,.2f}"
                                    } if "amount" in exceptions_display.columns else {}),
                                    use_container_width=True
                                )
                            else:
                                st.info("Zero cut-off exceptions! All invoice bookings correctly match physical delivery dates across the year-end boundary.")

                        with cut_sub2:
                            st.markdown("**All Tested Vouchers** (Population with cut-off compliance status)")
                            all_display = tables['all_vouchers'].copy()
                            for d_col in ["invoice_date", "grn_date"]:
                                if d_col in all_display.columns:
                                    all_display[d_col] = pd.to_datetime(all_display[d_col]).dt.strftime("%d-%b-%Y")
                            st.dataframe(all_display, use_container_width=True)

                    except Exception as e:
                        st.error(f"Error during Cut-Off analysis: {str(e)}")
                        st.exception(e)
        else:
            st.info("👆 Please upload your cut-off ledger or click **'Test with Sample Cut-Off'** above to view the analysis.")


    # =============================================================================
    # TAB 3: 10-YEAR SA 520 ANALYTICAL PROCEDURES & DUPONT ROE
    # =============================================================================
    with tab3:
        st.markdown("""
        <div class="section-title">
            <span>📈</span> 10-Year Financial Statement Analysis & SA 520 DuPont ROE Audit
        </div>
        <p style='color:#5A6A7E; font-size:0.9rem;'>
            Upload an audited financial statement workbook exported from <b>Screener.in</b> ('Export to Excel' file with raw <code>Data Sheet</code>). 
            The engine extracts 10 years of P&L and Balance Sheet figures, decomposes Return on Equity via the 3-step <b>DuPont Formula</b> 
            (<i>Net Margin × Asset Turnover × Equity Multiplier</i>), computes audit ratios, and evaluates automated <b>SA 520 red flags</b>.
        </p>
        """, unsafe_allow_html=True)

        col_scr_upload, col_scr_quick = st.columns([3, 1])
        with col_scr_upload:
            screener_file = st.file_uploader(
                "Upload Screener.in Excel File (.xlsx)",
                type=["xlsx"],
                key="screener_uploader",
                help="Upload standard 'Export to Excel' file from Screener.in containing the 'Data Sheet'"
            )
        with col_scr_quick:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            use_sample_screener = st.button("⚡ Test with Sample Screener (Infosys)", key="btn_sample_screener")

        active_screener_source = None
        if screener_file is not None:
            active_screener_source = ("uploaded", screener_file)
        elif use_sample_screener:
            active_screener_source = ("sample", os.path.join(SAMPLE_DIR, "sample_screener_infosys.xlsx"))

        if active_screener_source:
            with st.spinner("Extracting Screener financial blocks, calculating DuPont decomposition and SA 520 ratios..."):
                with tempfile.TemporaryDirectory() as tmp_dir:
                    if active_screener_source[0] == "uploaded":
                        input_path = os.path.join(tmp_dir, "screener_input.xlsx")
                        with open(input_path, "wb") as f:
                            f.write(active_screener_source[1].getvalue())
                    else:
                        input_path = active_screener_source[1]

                    output_excel_path = os.path.join(tmp_dir, "Project3_SA520_Ratio_Audit.xlsx")

                    try:
                        result = run_screener_ratio_audit(input_path, output_excel_path=output_excel_path)
                        summary = result["summary"]
                        tables = result["tables"]

                        with open(output_excel_path, "rb") as f:
                            report_bytes = f.read()

                        st.success("✅ SA 520 Ratio Audit & DuPont ROE Scorecard generated successfully!")

                        # Fetch latest figures for KPIs
                        full_df = tables["full_data"]
                        latest_row = full_df.iloc[-1]
                        latest_net_margin = float(latest_row.get("Net_Margin_%", 0.0))
                        latest_asset_to = float(latest_row.get("Asset_Turnover_x", 0.0))
                        latest_eq_mult = float(latest_row.get("Equity_Multiplier_x", 0.0))

                        # --- SUMMARY KPIS ---
                        s1, s2, s3, s4, s5 = st.columns(5)
                        with s1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Latest Fiscal Year</div>
                                <div class="metric-value">{summary['latest_year']}</div>
                                <div class="metric-sub">{len(summary['years_analyzed'])} Years Trended</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with s2:
                            st.markdown(f"""
                            <div class="metric-card success-card">
                                <div class="metric-label">Latest DuPont ROE %</div>
                                <div class="metric-value success">{summary['latest_dupont_roe_pct']:.2f}%</div>
                                <div class="metric-sub">Net Profit / Shareholders' Equity</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with s3:
                            st.markdown(f"""
                            <div class="metric-card {'alert-card' if summary['latest_debt_to_equity'] > 1.0 else 'success-card'}">
                                <div class="metric-label">Latest Debt / Equity</div>
                                <div class="metric-value {'danger' if summary['latest_debt_to_equity'] > 1.0 else ''}">{summary['latest_debt_to_equity']:.2f}x</div>
                                <div class="metric-sub">{'⚠️ High Leverage' if summary['latest_debt_to_equity'] > 1.0 else '✅ Conservative Leverage'}</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with s4:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Latest Net Margin %</div>
                                <div class="metric-value">{latest_net_margin:.2f}%</div>
                                <div class="metric-sub">Operating Profitability</div>
                            </div>
                            """, unsafe_allow_html=True)
                        with s5:
                            st.markdown(f"""
                            <div class="metric-card">
                                <div class="metric-label">Asset Turnover</div>
                                <div class="metric-value">{latest_asset_to:.2f}x</div>
                                <div class="metric-sub">Leverage: {latest_eq_mult:.2f}x</div>
                            </div>
                            """, unsafe_allow_html=True)

                        # --- DOWNLOAD EXCEL REPORT BUTTON ---
                        st.download_button(
                            label="📥 Download SA 520 Ratio Audit Excel Report (3 Tabs)",
                            data=report_bytes,
                            file_name="Project3_SA520_Ratio_Audit.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_screener_report"
                        )

                        st.markdown("<div style='margin-top:1.5rem;'></div>", unsafe_allow_html=True)

                        # --- DUPONT ROE & NET MARGIN TREND CHART ---
                        st.markdown("#### 📊 10-Year DuPont ROE % and Net Margin % Trajectory")

                        years = full_df["Financial_Year"].tolist()
                        roe_vals = full_df["DuPont_ROE_%"].tolist()
                        margin_vals = full_df["Net_Margin_%"].tolist()

                        fig = go.Figure()

                        fig.add_trace(go.Scatter(
                            x=years,
                            y=roe_vals,
                            mode="lines+markers",
                            name="DuPont ROE %",
                            line=dict(color="#1F4E78", width=3),
                            marker=dict(size=8, symbol="circle", color="#1F4E78"),
                            hovertemplate="<b>%{x}</b><br>DuPont ROE: %{y:.2f}%<extra></extra>"
                        ))

                        fig.add_trace(go.Scatter(
                            x=years,
                            y=margin_vals,
                            mode="lines+markers",
                            name="Net Margin %",
                            line=dict(color="#0D8050", width=3, dash="solid"),
                            marker=dict(size=8, symbol="diamond", color="#0D8050"),
                            hovertemplate="<b>%{x}</b><br>Net Margin: %{y:.2f}%<extra></extra>"
                        ))

                        fig.update_layout(
                            title=dict(
                                text="DuPont Return on Equity vs. Net Profit Margin (10-Year Multi-Period Comparison)",
                                font=dict(color="#1F4E78", size=15)
                            ),
                            xaxis=dict(title="Financial Year", gridcolor="#EBF1F5"),
                            yaxis=dict(title="Percentage (%)", gridcolor="#EBF1F5", zeroline=True),
                            plot_bgcolor="#FFFFFF",
                            paper_bgcolor="#FFFFFF",
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1
                            ),
                            margin=dict(l=40, r=40, t=50, b=40),
                            hovermode="x unified"
                        )

                        st.plotly_chart(fig, use_container_width=True)

                        # --- DETAILED TABLES ---
                        scr_sub1, scr_sub2, scr_sub3 = st.tabs([
                            "📋 10-Year Ratio Scorecard",
                            "🚨 SA 520 Risk Alerts",
                            "📑 Full Extracted Data"
                        ])

                        with scr_sub1:
                            st.markdown("**10-Year Audit Ratio & DuPont ROE Decomposition Scorecard**")
                            st.dataframe(tables['scorecard'], use_container_width=True)

                        with scr_sub2:
                            st.markdown("**SA 520 Automated Red-Flag Exception Register**")
                            st.markdown("""
                            - **Solvency Flag**: ALERT if Debt-to-Equity > 1.0x (over-leveraged balance sheet)
                            - **Collection Flag**: ALERT if Trade Receivables Days > 75 Days (working capital lock-in / revenue recognition risk)
                            - **Coverage Flag**: ALERT if Interest Coverage Ratio < 2.5x (debt service distress risk)
                            """)

                            alerts_display = tables['risk_alerts'].copy()

                            # Apply conditional formatting compatible with all Pandas versions
                            def alert_styler(val):
                                if isinstance(val, str) and "ALERT" in val:
                                    return "background-color: #FFC7CE; color: #9C0006; font-weight: bold;"
                                return ""

                            styled_alerts = alerts_display.style
                            if hasattr(styled_alerts, "map"):
                                styled_alerts = styled_alerts.map(alert_styler)
                            else:
                                styled_alerts = styled_alerts.applymap(alert_styler)

                            st.dataframe(styled_alerts, use_container_width=True)

                        with scr_sub3:
                            st.markdown("**Full Extracted and Transposed Financial Statement Dataset**")
                            st.dataframe(tables['full_data'], use_container_width=True)

                    except Exception as e:
                        st.error(f"Error during SA 520 Screener analysis: {str(e)}")
                        st.exception(e)
        else:
            st.info("👆 Please upload your Screener.in workbook (.xlsx) or click **'Test with Sample Screener (Infosys)'** above to view the analysis.")

    # -----------------------------------------------------------------------------
    # FOOTER
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #7F8C8D; font-size: 0.85rem; padding: 1rem 0;'>
        CA Forensic Audit & Financial Intelligence Suite &bull; Built with Python & Streamlit &bull; 
        Standard on Auditing (SA) 520 & Statutory Tax Compliance Engine &bull; Executive Theme (#1F4E78)
    </div>
    """, unsafe_allow_html=True)

# Run Streamlit UI when launched directly or via `streamlit run`
if __name__ == "__main__" or any("streamlit" in arg.lower() for arg in sys.argv):
    run_streamlit_app()
