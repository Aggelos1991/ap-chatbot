import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO

st.set_page_config(page_title="Excel Aggregation — Color Groups", layout="wide")
st.title("📊 Sheet2 → Sheet1 Aggregation (Color-Group Logic, Add Mode)")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# ---------- FIXED CONFIG ----------
# Sheet positions & column indices (1-based)
SHEET1_COL_TITLE = 6     # F (Greek title used for mapping)
SHEET1_COL_J = 10        # J (Χρεωστικό Υπόλοιπο)
SHEET1_COL_K = 11        # K (Πιστωτικό Υπόλοιπο)
SHEET1_COL_ACCT = 4      # D (account code like 50.00.00.0000)

SHEET2_COL_D2 = 2        # B (Διάσταση 2)
SHEET2_COL_K = 11        # K
SHEET2_COL_L = 12        # L

# Explicit zero accounts (only these are zeroed)
ZERO_ACCOUNTS = {
    "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
    "50.01.00.0000","50.01.01.0000","50.05.00.0000"
}

# Διάσταση 2 → Greek title (Sheet1 col F)
D2_TO_TITLE = {
    "--": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "01 - OpEx Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "02 - CapEx Payables": "Προμηθευτές πιστωτικά υπόλοιπα τέλους περιόδου",
    "03 - Other Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "04 - OpEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
    "05 - CapEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Προκαταβολές για αγορές Παγίων",
    "06 - Other Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
    "100 - General B2B Invoices – Payments": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "110 - B2B Aging collections": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "2200 - Development Capex": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "300 - Financing Cashflows": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
}
TITLE_TO_D2 = {v: k for k, v in D2_TO_TITLE.items()}

# Simulated “color” groups (your permanent rule)
COLOR_GROUPS = {
    "GROUP_A": ["01 - OpEx Payables", "03 - Other Payables", "--"],
    "GROUP_B": ["02 - CapEx Payables"],
    "GROUP_C": ["04 - OpEx Advances", "05 - CapEx Advances", "06 - Other Advances"],
    "GROUP_D": ["100 - General B2B Invoices – Payments",
                "110 - B2B Aging collections",
                "2200 - Development Capex",
                "300 - Financing Cashflows"],
}

def find_group_for_title(title: str):
    """Get group name for a Greek title from Sheet1 column F."""
    d2 = TITLE_TO_D2.get(title.strip(), None)
    if not d2:
        return None
    for g, members in COLOR_GROUPS.items():
        if d2 in members:
            return g
    return None

def autofit(ws):
    """Lightweight auto-fit without touching styles."""
    for col in ws.columns:
        max_len = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            v = cell.value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max_len + 2, 45)

if uploaded:
    try:
        # Load workbook from uploaded bytes (preserves formatting)
        data = uploaded.read()
        wb = load_workbook(BytesIO(data))
        if len(wb.worksheets) < 2:
            st.error("The file must have at least two sheets (Sheet1 target, Sheet2 source).")
            st.stop()

        ws1, ws2 = wb.worksheets[0], wb.worksheets[1]

        # --- Step 1: Aggregate Sheet2 by COLOR_GROUPS using Διάσταση 2 (B) ---
        group_sums = {g: {"K": 0.0, "L": 0.0} for g in COLOR_GROUPS}
        for r in range(2, ws2.max_row + 1):
            d2 = ws2.cell(r, SHEET2_COL_D2).value
            if not d2:
                continue
            d2 = str(d2).strip()
            try:
                k_val = float(ws2.cell(r, SHEET2_COL_K).value or 0)
            except:
                k_val = 0.0
            try:
                l_val = float(ws2.cell(r, SHEET2_COL_L).value or 0)
            except:
                l_val = 0.0

            for g, members in COLOR_GROUPS.items():
                if d2 in members:
                    group_sums[g]["K"] += k_val
                    group_sums[g]["L"] += l_val
                    break

        # --- Step 2: Push into Sheet1 (ADD mode) on J & K, by column F mapping ---
        updated_rows = 0
        zeroed_rows = 0
        for r in range(2, ws1.max_row + 1):
            acct = str(ws1.cell(r, SHEET1_COL_ACCT).value or "").strip()
            title = str(ws1.cell(r, SHEET1_COL_TITLE).value or "").strip()
            if not title:
                continue

            # Explicit zero accounts
            if acct in ZERO_ACCOUNTS:
                ws1.cell(r, SHEET1_COL_J, 0)
                ws1.cell(r, SHEET1_COL_K, 0)
                zeroed_rows += 1
                continue

            group = find_group_for_title(title)
            if not group:
                continue

            curJ = ws1.cell(r, SHEET1_COL_J).value or 0
            curK = ws1.cell(r, SHEET1_COL_K).value or 0
            try:
                curJ = float(curJ)
            except:
                curJ = 0.0
            try:
                curK = float(curK)
            except:
                curK = 0.0

            ws1.cell(r, SHEET1_COL_J, curJ + group_sums[group]["K"])  # J gets aggregated K
            ws1.cell(r, SHEET1_COL_K, curK + group_sums[group]["L"])  # K gets aggregated L
            updated_rows += 1

        # --- Step 3: Auto-fit widths (non-destructive) ---
        for w in wb.worksheets:
            autofit(w)

        # Return file to user
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success(f"✅ Done. Updated rows: {updated_rows} • Zeroed rows: {zeroed_rows}")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name=f"Updated_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Optional: show the group totals used
        with st.expander("Show group totals used (K→J, L→K)"):
            st.write(group_sums)

    except Exception as e:
        st.error(f"❌ Error: {e}")
