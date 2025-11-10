import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Διάσταση 2 Aggregator — Add Mode", layout="wide")
st.title("📊 Aggregate Sheet2 ➜ Add to Sheet1 (K & L)")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

# === ZERO ACCOUNTS (unchanged) ===
ZERO_ACCOUNTS = {
    "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
    "50.01.00.0000","50.01.01.0000","50.05.00.0000"
}

# === Mapping Διάσταση2 ➜ Greek column F titles ===
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
    """Locate a column by partial header name."""
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val and name in str(val):
            return c
    return None

def autofit(ws):
    """Auto-adjust all column widths."""
    for col in ws.columns:
        max_len = 0
        letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = max_len + 2

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]     # Sheet 1 — target
        ws2 = wb.worksheets[1]     # Sheet 2 — source

        # ==== Aggregate Sheet 2 (Διάσταση 2 → sum of K & L) ====
        aggK, aggL = {}, {}
        for r in range(2, ws2.max_row + 1):
            d2 = str(ws2.cell(r, 2).value or "").strip()
            if not d2:
                continue
            try: k_val = float(ws2.cell(r, 11).value or 0)
            except: k_val = 0.0
            try: l_val = float(ws2.cell(r, 12).value or 0)
            except: l_val = 0.0
            aggK[d2] = aggK.get(d2, 0.0) + k_val
            aggL[d2] = aggL.get(d2, 0.0) + l_val

        # ==== Identify Sheet 1 columns ====
        acct_col   = find_col(ws1, "Λογαριασμός") or 2
        title_col  = 6   # explicit column F mapping
        debit_col  = find_col(ws1, "Χρεωστικό Υπόλοιπο")  # J
        credit_col = find_col(ws1, "Πιστωτικό Υπόλοιπο")  # K
        if not debit_col or not credit_col:
            raise ValueError("Columns J/K not found in Sheet1.")

        # ==== Update Sheet 1 ====
        for r in range(2, ws1.max_row + 1):
            acct = str(ws1.cell(r, acct_col).value or "").strip()
            if acct in ZERO_ACCOUNTS:
                ws1.cell(r, debit_col, 0)
                ws1.cell(r, credit_col, 0)
                continue

            title = str(ws1.cell(r, title_col).value or "").strip()
            d2_key = TITLE_TO_D2.get(title, "")

            if d2_key and (d2_key in aggK or d2_key in aggL):
                # ADD the aggregated values to existing ones
                try:
                    oldK = float(ws1.cell(r, credit_col).value or 0)
                    oldL = float(ws1.cell(r, debit_col).value or 0)
                except:
                    oldK, oldL = 0.0, 0.0

                newK = oldK + aggK.get(d2_key, 0.0)
                newL = oldL + aggL.get(d2_key, 0.0)

                ws1.cell(r, credit_col, newK)
                ws1.cell(r, debit_col, newL)

        # ==== Auto-fit all sheets ====
        for ws in wb.worksheets:
            autofit(ws)

        # ==== Save result ====
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Aggregated totals added to Sheet 1 (J & K) successfully. Formatting preserved.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
