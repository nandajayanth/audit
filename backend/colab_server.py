# ==============================================================
# COPY & PASTE THIS INTO A NEW CELL IN YOUR GOOGLE COLAB NOTEBOOK
# ==============================================================
"""
This script launches the Forensic Audit API directly inside your Google Colab instance,
and creates a secure public tunnel via ngrok so your HTML/CSS/JS frontend can connect to it!
"""

!pip install fastapi uvicorn pyngrok nest-asyncio python-multipart openpyxl -q

import nest_asyncio
import uvicorn
from pyngrok import ngrok
import threading

# 1. Paste your ngrok auth token here (Sign up free at https://dashboard.ngrok.com/get-started/your-authtoken)
NGROK_AUTH_TOKEN = "" # @param {type:"string"}

if NGROK_AUTH_TOKEN:
    ngrok.set_auth_token(NGROK_AUTH_TOKEN)

# 2. Write the backend app.py file directly inside Colab
colab_code = '''
# Imports and FastAPI app
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import io, re, numpy as np, pandas as pd
from openpyxl.styles import Font, PatternFill
from fastapi.responses import StreamingResponse

app = FastAPI(title="Colab Forensic Audit Backend")

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

@app.get("/")
def home():
    return {"status": "online", "message": "Google Colab Forensic Audit Backend is running!"}

@app.get("/api/ap-audit/generate-mock")
def generate_mock(rows_count: int = 500, seed: int = 42):
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
        rows.append({
            "voucher_no": f"V-{1000 + i}",
            "vendor": v_name,
            "gstin": v_gstin,
            "invoice_no": f"INV-{50000 + i}",
            "invoice_date": f"2026-03-{inv_day:02d}",
            "grn_date": f"2026-03-{inv_day + 2:02d}",
            "section": sec,
            "amount": f"{amt:,.2f}",
            "gst_claimed": f"{gst:,.2f}",
            "tds_deducted": f"{tds:,.2f}"
        })

    df = pd.DataFrame(rows)
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

    return {"status": "success", "rows": df.where(pd.notnull(df), None).to_dict(orient="records")}
'''

with open("colab_app.py", "w") as f:
    f.write(colab_code)

# 3. Apply nest_asyncio and start uvicorn in a background thread
nest_asyncio.apply()

def run_uvicorn():
    uvicorn.run("colab_app:app", host="127.0.0.1", port=8000, log_level="warning")

server_thread = threading.Thread(target=run_uvicorn, daemon=True)
server_thread.start()

# 4. Open public ngrok tunnel
try:
    public_url = ngrok.connect(8000)
    print("\n" + "="*60)
    print("SUCCESS! Your Colab Backend is LIVE at:")
    print(public_url)
    print("="*60)
    print("Copy the URL above and paste it into your Frontend 'API Settings' modal!")
except Exception as e:
    print("Could not start ngrok automatically:", e)
    print("You can also run locally or use Cloudflare tunnel / localtunnel.")
