import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Aggregator & Mapper", layout="wide")
st.title("📘 Excel Aggregator & Mapper — Sheet2 Διάσταση 2 → Sheet1 Τίτλος")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]   # Sheet 1 target
        ws2 = wb.worksheets[1]   # Sheet 2 source

        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002",
            "50.00.00.0003","50.01.00.0000","50.01.01.0000","50.05.00.0000"
        ]

        # --- Delete column E (Λογαριασμός λογιστικής) ---
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Λογαριασμός λογιστικής":
                ws1.delete_cols(col)
                break

        # --- Find Πιστωτικό Υπόλοιπο column ---
        target_col = None
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Πιστωτικό Υπόλοιπο":
                target_col = col
                break
        if not target_col:
            st.error("❌ Column 'Πιστωτικό Υπόλοιπο' not found in Sheet 1.")
            st.stop()

        insert_pos = target_col + 1
        ws1.insert_cols(insert_pos)
        ws1.cell(row=1, column=insert_pos, value="Διάσταση 2 (Source)")

        # --- Read Sheet 2 to DataFrame ---
        df2 = pd.DataFrame(ws2.values)
        df2.columns = df2.iloc[0]
        df2 = df2.drop(0)

        col_dim2 = next((c for c in df2.columns if "Διάσταση" in str(c)), None)
        col_K = df2.columns[10] if len(df2.columns) > 10 else df2.columns[-2]
        col_L = df2.columns[11] if len(df2.columns) > 11 else df2.columns[-1]

        if not col_dim2:
            st.error("❌ Could not find 'Διάσταση 2' column in Sheet 2.")
            st.stop()

        df2[col_K] = pd.to_numeric(df2[col_K], errors="coerce").fillna(0)
        df2[col_L] = pd.to_numeric(df2[col_L], errors="coerce").fillna(0)

        # --- Aggregate ---
        agg = df2.groupby(col_dim2)[[col_K, col_L]].sum().reset_index()

        # --- Read Sheet 1 → DataFrame to access Τίτλος values ---
        df1 = pd.DataFrame(ws1.values)
        df1.columns = df1.iloc[0]
        df1 = df1.drop(0)

        if "Τίτλος" not in df1.columns:
            st.error("❌ Column 'Τίτλος' not found in Sheet 1.")
            st.stop()

        # Build mapping dictionary directly from Sheet 2 (Διάσταση 2 → aggregated sum)
        mapping = dict(zip(agg[col_dim2], agg[col_K] + agg[col_L]))

        # --- Fill Διάσταση 2 (Source) in Sheet 1 ---
        for row in range(2, ws1.max_row + 1):
            title = str(ws1.cell(row=row, column=list(df1.columns).index("Τίτλος")+1).value).strip() if "Τίτλος" in df1.columns else ""
            k_cell = ws1.cell(row=row, column=11)
            l_cell = ws1.cell(row=row, column=12)
            acc = str(ws1.cell(row=row, column=4).value).strip() if ws1.cell(row=row, column=4).value else ""

            if acc in zero_accounts:
                k_cell.value = 0
                l_cell.value = 0
                ws1.cell(row=row, column=insert_pos, value="")
            else:
                matched_value = mapping.get(title, "")
                ws1.cell(row=row, column=insert_pos, value=matched_value)

        # --- Auto-fit columns ---
        for col in ws1.columns:
            maxlen = max((len(str(c.value)) for c in col if c.value), default=0)
            ws1.column_dimensions[col[0].column_letter].width = maxlen + 2

        # --- Save ---
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Mapping (Διάσταση 2 → Τίτλος) completed successfully!")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
