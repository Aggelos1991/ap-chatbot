import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Aggregator & Mapper", layout="wide")
st.title("📘 Excel Aggregator & Mapper — Formatting Preserved")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]  # sheet1 target
        ws2 = wb.worksheets[1]  # sheet2 source

        # Zero accounts
        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002",
            "50.00.00.0003","50.01.00.0000","50.01.01.0000","50.05.00.0000"
        ]

        # Mapping dictionary
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
            "06 - Other Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
            "05 - CapEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Προκαταβολές για αγορές Παγίων"
        }

        # --- STEP 1 — Remove column E completely ---
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Λογαριασμός λογιστικής":
                ws1.delete_cols(col)
                break

        # --- STEP 2 — Find "Πιστωτικό Υπόλοιπο" and insert the new column after it ---
        target_col = None
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Πιστωτικό Υπόλοιπο":
                target_col = col
                break

        if not target_col:
            st.error("❌ Column 'Πιστωτικό Υπόλοιπο' not found.")
            st.stop()

        insert_pos = target_col + 1
        ws1.insert_cols(insert_pos)
        ws1.cell(row=1, column=insert_pos, value="Διάσταση 2 (Source)")

        # --- STEP 3 — Aggregate from Sheet 2 ---
        df2 = pd.DataFrame(ws2.values)
        df2.columns = df2.iloc[0]
        df2 = df2.drop(0)

        # Identify relevant columns dynamically
        col_B = [c for c in df2.columns if "Διάσταση" in str(c)][0]
        col_K = df2.columns[10]
        col_L = df2.columns[11]

        df2[col_K] = pd.to_numeric(df2[col_K], errors="coerce").fillna(0)
        df2[col_L] = pd.to_numeric(df2[col_L], errors="coerce").fillna(0)

        agg = df2.groupby(col_B)[[col_K, col_L]].sum().reset_index()
        agg["Διάσταση 2 (Source)"] = agg[col_B].map(mapping)

        # --- STEP 4 — Update Sheet 1 values ---
        for row in range(2, ws1.max_row + 1):
            acc = str(ws1.cell(row=row, column=4).value).strip() if ws1.cell(row=row, column=4).value else ""
            k_cell = ws1.cell(row=row, column=11)
            l_cell = ws1.cell(row=row, column=12)

            if acc in zero_accounts:
                k_cell.value = 0
                l_cell.value = 0
                ws1.cell(row=row, column=insert_pos, value="")
            else:
                match = agg.loc[agg[col_B] == acc, "Διάσταση 2 (Source)"]
                ws1.cell(row=row, column=insert_pos, value=match.iloc[0] if not match.empty else "")

        # --- STEP 5 — Auto-fit widths ---
        for col in ws1.columns:
            maxlen = max((len(str(c.value)) for c in col if c.value), default=0)
            ws1.column_dimensions[col[0].column_letter].width = maxlen + 2

        # --- Save to memory ---
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Excel updated — formatting and structure preserved.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
