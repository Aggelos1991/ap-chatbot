import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Aggregator — Update J & K", layout="wide")
st.title("📊 Update Sheet1 J & K from Sheet2 Aggregation")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

# === ZERO ACCOUNTS ===
ZERO_ACCOUNTS = {
    "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
    "50.01.00.0000","50.01.01.0000","50.05.00.0000"
}

# === MAPPING ===
D2_TO_TITLE = {
    "--": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "01 - OpEx Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "03 - Other Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "100 - General B2B Invoices – Payments": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "110 - B2B Aging collections": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "2200 - Development Capex": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "300 - Financing Cashflows": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "02 - CapEx Payables": "Προμηθευτές πιστωτικά υπόλοιπα τέλους περιόδου",
    "04 - OpEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
    "05 - CapEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Προκαταβολές για αγορές Παγίων",
    "06 - Other Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες"
}
TITLE_TO_D2 = {v.strip(): k for k, v in D2_TO_TITLE.items()}

def find_col(ws, name):
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v and name in str(v):
            return c
    return None

def autofit_columns(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]   # Sheet1 (target)
        ws2 = wb.worksheets[1]   # Sheet2 (source)

        # --- Sheet2 Aggregation ---
        aggK = {}
        aggL = {}
        for r in range(2, ws2.max_row + 1):
            d2 = str(ws2.cell(r, 2).value or "").strip()
            if not d2:
                continue
            try:
                k_val = float(ws2.cell(r, 11).value or 0)
            except:
                k_val = 0.0
            try:
                l_val = float(ws2.cell(r, 12).value or 0)
            except:
                l_val = 0.0
            aggK[d2] = aggK.get(d2, 0.0) + k_val
            aggL[d2] = aggL.get(d2, 0.0) + l_val

        # --- Find key columns in Sheet1 ---
        acct_col = find_col(ws1, "Λογαριασμός") or 2
        title_col = find_col(ws1, "Τίτλος")
        debit_col = find_col(ws1, "Χρεωστικό Υπόλοιπο")  # J
        credit_col = find_col(ws1, "Πιστωτικό Υπόλοιπο") # K
        src_col = find_col(ws1, "Διάσταση 2 (Source)")
        if not src_col:
            src_col = credit_col + 1
            ws1.cell(1, src_col, "Διάσταση 2 (Source)")

        # --- Update Sheet1 ---
        for r in range(2, ws1.max_row + 1):
            acct = str(ws1.cell(r, acct_col).value or "").strip()
            if acct in ZERO_ACCOUNTS:
                ws1.cell(r, debit_col, 0)
                ws1.cell(r, credit_col, 0)
                continue

            d2_key = ""
            if title_col:
                title = str(ws1.cell(r, title_col).value or "").strip()
                d2_key = TITLE_TO_D2.get(title, "")

            if d2_key and (d2_key in aggK or d2_key in aggL):
                ws1.cell(r, debit_col, aggL.get(d2_key, 0.0))   # J ← aggregated L
                ws1.cell(r, credit_col, aggK.get(d2_key, 0.0))  # K ← aggregated K
                ws1.cell(r, src_col, d2_key)

        # --- Auto fit columns for all sheets ---
        for ws in wb.worksheets:
            autofit_columns(ws)

        # --- Save ---
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Aggregation complete — J & K replaced, Διάσταση 2 mapped, widths auto-adjusted.")
        st.download_button("⬇️ Download Updated Excel",
                           data=out,
                           file_name="Updated_" + uploaded.name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error: {e}")
