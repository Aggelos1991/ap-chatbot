import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Διάσταση 2 Aggregator", layout="wide")
st.title("📘 Διάσταση 2 → Τίτλος Mapping & Aggregation")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

# Zero accounts start with 50.*
def is_zero_account(val):
    try:
        return str(val).strip().startswith("50")
    except:
        return False

# Mapping (Διάσταση 2 → Τίτλος)
mapping = {
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
reverse_mapping = {v: k for k, v in mapping.items()}

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]  # Sheet1
        ws2 = wb.worksheets[1]  # Sheet2

        # --- Delete column E (Λογαριασμός λογιστικής)
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Λογαριασμός λογιστικής":
                ws1.delete_cols(col)
                break

        # --- Locate key columns ---
        def find_col(ws, keyword):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=c).value
                if val and keyword in str(val):
                    return c
            return None

        col_d2 = find_col(ws2, "Διάσταση 2")
        col_K = find_col(ws2, "Χρεωστικό Υπόλοιπο - Σύνολα")
        col_L = find_col(ws2, "Πιστωτικό Υπόλοιπο - Σύνολα")
        col_titlos = find_col(ws1, "Τίτλος")
        col_credit = find_col(ws1, "Πιστωτικό Υπόλοιπο")

        if not all([col_d2, col_K, col_L, col_titlos, col_credit]):
            st.error("❌ Missing one of required columns (Διάσταση 2, Κ, L, Τίτλος, Πιστωτικό Υπόλοιπο).")
            st.stop()

        # --- Aggregate K+L totals from Sheet2 ---
        aggregates = {}
        for r in range(2, ws2.max_row + 1):
            d2 = str(ws2.cell(r, col_d2).value).strip() if ws2.cell(r, col_d2).value else ""
            if not d2:
                continue
            k_val = float(ws2.cell(r, col_K).value or 0)
            l_val = float(ws2.cell(r, col_L).value or 0)
            aggregates[d2] = aggregates.get(d2, 0) + k_val + l_val

        # --- Insert new column after Πιστωτικό Υπόλοιπο ---
        insert_col = col_credit + 1
        ws1.insert_cols(insert_col)
        ws1.cell(1, insert_col, "Διάσταση 2 (Source)")

        # --- Update Sheet1 ---
        for r in range(2, ws1.max_row + 1):
            acc = ws1.cell(r, 4).value
            titlos = str(ws1.cell(r, col_titlos).value or "").strip()

            if is_zero_account(acc):
                # Zeroed accounts
                ws1.cell(r, col_K, 0)
                ws1.cell(r, col_L, 0)
                ws1.cell(r, insert_col, "")
                continue

            d2_key = reverse_mapping.get(titlos)
            if d2_key and d2_key in aggregates:
                ws1.cell(r, col_K, aggregates[d2_key])
                ws1.cell(r, col_L, aggregates[d2_key])
                ws1.cell(r, insert_col, d2_key)
            else:
                ws1.cell(r, insert_col, "")

        # --- Save back ---
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Aggregation complete. Διάσταση 2 and K/L updated in Sheet1.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
