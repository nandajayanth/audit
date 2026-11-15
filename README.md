# AuditGenix &mdash; Forensic Audit & SA 520 Intelligence Suite

A responsive, high-performance web application designed for forensic financial auditing, statutory tax compliance surveillance (TDS Section 194J/194C & Section 40(a)(ia)), duplicate payment mitigation, year-end cut-off testing, and **SA 520 Analytical Procedures with 3-Stage DuPont ROE Decomposition**.

This project connects to and replicates the backend logic from the Google Colab notebook:
🔗 **[Google Colab Notebook Reference](https://colab.research.google.com/drive/1tZJyT0vdKfw9iVeZv4YmgYV9gTHFWi5S?usp=sharing)**

---

## 🌟 Key Capabilities & Features

### 1. Accounts Payable (AP) Ledger Forensic Audit (Cells 0 &ndash; 3)
* **Synthetic ERP Ledger Generator**: Reproduces the exact 500-row messy ERP dataset from Colab (`seed=42`) with realistic injected anomalies.
* **Statutory TDS Surveillance & Sec 40(a)(ia) Disallowance**:
  * Evaluates **Section 194J (10%)** and **Section 194C (2%)** on payments exceeding ₹30,000.
  * Calculates exact `tds_shortfall` (`expected_tds - tds_deducted`).
  * Computes statutory **30% Tax Disallowance** under Section 40(a)(ia) of the Indian Income Tax Act.
* **15-Character GSTIN Integrity Audit**:
  * Regex structural audit: `^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$`
  * Catches missing 14-character checksum digits, corrupted PAN formats, and unrecorded vendor GSTINs risking Input Tax Credit (ITC) disallowance.
* **Duplicate Invoice Detection**:
  * Scans across all vouchers for duplicate invoice numbers and cross-vendor billing collisions to prevent double payments.
* **Year-End Cut-Off Testing (March 31 Boundary)**:
  * Detects **Early Booking** (`invoice_date <= year_end` & `grn_date > year_end`), which overstates expenses in the closing financial year.
  * Detects **Late Booking** (`invoice_date > year_end` & `grn_date <= year_end`), which unrecords closing liabilities.
* **Vendor Risk Matrix**: Aggregated exposure ranking by Vendor & GSTIN with total spend, voucher counts, TDS shortfalls, and disallowances.
* **Master 6-Tab Excel Workpaper Export**: Generates `.xlsx` workbooks identical to `Master_6Tab_AP_Audit_Report.xlsx` with corporate blue headers and number formatting.

---

### 2. SA 520 Analytical Procedures & DuPont ROE (Cells 4 &ndash; 9)
* **DuPont 3-Way ROE Decomposition**:
  $$\text{DuPont ROE} = \text{Net Profit Margin} \times \text{Asset Turnover} \times \text{Equity Multiplier}$$
  Decomposes operational efficiency, asset utilization, and balance sheet leverage.
* **Comprehensive Financial Ratios**:
  * **Margins**: Gross Margin %, EBIT Margin %, Net Profit Margin %.
  * **Liquidity**: Current Ratio (vs 1.33x banking MPBF benchmark), Quick Ratio.
  * **Solvency**: Debt-to-Equity (vs 0.85x / 1.0x limits), Interest Coverage Ratio (EBIT / Interest vs 2.5x threshold).
  * **Efficiency**: Receivables Days / DSO (vs 60 / 75-day fictitious sales alert), Inventory Days / DIO.
* **Automated Red-Flag Diagnostics**: Color-coded severity badges for liquidity crunches, excessive leverage, revenue cut-off risks, and debt service strains.
* **Multi-Year Working Paper Scorecard**: Transposed layout (years as columns, ratios as rows) with YoY variance calculations.
* **Screener.in Excel Integration**: Drag-and-drop parsing of raw 10-year Excel exports from Screener.in ("Data Sheet" Profit & Loss and Balance Sheet blocks).

---

## 🚀 How to Run the Application

### Option A: Standalone Browser Mode (Zero Installation, 100% Offline)
1. Simply double-click **`index.html`** or open it in any web browser (Google Chrome, Microsoft Edge, Mozilla Firefox, Safari).
2. The built-in client engine (`audit_engine.js` + `sa520_engine.js`) will instantly compute all forensic anomalies, generate mock datasets, parse uploaded Excel files, and export formatted `.xlsx` workbooks right inside your browser!

### Option B: Local HTTP Server
If you prefer running via a local web server:
```bash
# Using Python
python -m http.server 3000

# Using Node.js
npx serve .
```
Then navigate to `http://localhost:3000`.

### Option C: Streamlit Executive Audit Suite (Python)
Launch the executive 3-tab Streamlit web application:
```bash
cd ca-audit-suite
pip install -r requirements.txt
streamlit run app.py
```
Then navigate to `http://localhost:8501`.

---

## 🔗 Connecting to Google Colab Backend

You can connect this frontend directly to your active Google Colab notebook instance:

1. Open your Google Colab notebook at:
   `https://colab.research.google.com/drive/1tZJyT0vdKfw9iVeZv4YmgYV9gTHFWi5S?usp=sharing`
2. Add a new code cell and paste the contents of **`backend/colab_server.py`**:
   ```python
   !pip install fastapi uvicorn pyngrok nest-asyncio python-multipart openpyxl -q

   import nest_asyncio, uvicorn, threading
   from pyngrok import ngrok
   from backend.app import app

   nest_asyncio.apply()
   threading.Thread(target=lambda: uvicorn.run(app, host="127.0.0.1", port=8000), daemon=True).start()
   public_url = ngrok.connect(8000)
   print("Colab Backend LIVE at:", public_url)
   ```
3. Run the cell. Colab will display your public URL (e.g. `https://xxxx-xx-xx.ngrok-free.app`).
4. In the web application:
   - Click on **"Colab Backend Bridge"** in the top navigation bar.
   - Select **"Google Colab REST API Backend"**.
   - Paste your ngrok URL and click **"Ping API"**.

---

## 📁 Project Structure

```text
├── ca-audit-suite/          # Streamlit Executive Forensic & Statutory Audit Suite
│   ├── app.py               # Modern 3-tab executive UI (Navy #1F4E78 theme)
│   ├── audit_engine.py      # Core audit logic extracted from audit_backend.ipynb
│   ├── requirements.txt     # Python dependencies (streamlit, pandas, plotly, etc.)
│   └── sample_data/         # Sample ERP AP ledger, Cut-off CSV, and Screener XLSX
├── index.html               # Responsive Single-Page Application (HTML5)
├── styles.css               # Corporate Dark/Light Theme & Responsive Design (CSS3)
├── app.js                   # Application State, Event Handlers & Chart.js Controllers
├── audit_engine.js          # AP Forensic Engine (TDS, GSTIN, Duplicates, Cut-Off)
├── sa520_engine.js          # SA 520 Analytical Procedures & DuPont ROE Engine
├── export_service.js        # Multi-Tab Excel Workbook Generator (SheetJS / openpyxl style)
├── notebook.json            # Downloaded Colab notebook source
├── backend/
│   ├── app.py               # Complete FastAPI REST API Server
│   ├── colab_server.py      # Ready-to-run Colab code cell with ngrok
│   └── requirements.txt     # Python dependencies
└── README.md                # Documentation & Architecture Overview
```

---

## 🛡️ Statutory References
* **Section 194J & 194C, Indian Income Tax Act**: TDS on professional/technical services and contractor payments.
* **Section 40(a)(ia), Indian Income Tax Act**: 30% expenditure disallowance for failure to deduct/deposit TDS.
* **Standard on Auditing (SA) 520**: Analytical Procedures in financial audit planning and substantive testing.
* **Central Goods and Services Tax (CGST) Act**: Valid 15-character GSTIN structure and ITC eligibility criteria.
