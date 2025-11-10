import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Aggregator & Mapper", layout="wide")
st.title("📘 Excel Aggregator & Mapper — Διάσταση 2 → Τίτλος Mapping")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]  # Sheet 1 (target)
        ws2 = wb.worksheets[1]  # Sheet 2 (source)

        # Zero accounts
        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002",
            "50.00.00.0003","50.01.00.0000","50.01.01.0000","50.05.00.0000"
        ]

        # Delete column E
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Λογαριασμός λογιστικής":
                ws1.delete_cols(col)
                break

        # Find Πιστωτικό Υπόλοιπο
        target_col = None
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Πιστωτικό Υπόλοιπο":
                target_col = col
                break
        if not target_col:
            st.error("❌ Column 'Πιστωτικό Υπόλοιπο' not found in Sheet1.")
            st.stop()

        insert_pos = target_col + 1
        ws1.insert_cols(insert_pos)
        ws1.cell(row=1, column=insert_pos, value="Διάσταση 2 (Source)")

        # Convert Sheet2 to DataFrame
        df2 = pd.DataFrame(ws2.values)
        df2.columns = df2.iloc[0]
        df2 = df2.drop(0)

        # Identify correct columns dynamically
        col_dim2 = [c for c in df2.columns if "Διάσταση" in str(c)][0]
        col_titlos = [c for c in df2.columns if "Τίτλος" in str(c)][0]
        col_K = df2.columns[10]
        col_L = df2.columns[11]

        df2[col_K] = pd.to_numeric(df2[col_K], errors="coerce").fillna(0)
        df2[col_L] = pd.to_numeric(df2[col_L], errors="coerce").fillna(0)

        # Aggregate K + L by Διάσταση 2
        agg = df2.groupby(col_dim2)[[col_K, col_L]].sum().reset_index()

        # Build mapping dict directly from Sheet2
        map_from_sheet2 = df2.set_index(col_dim2)[col_titlos].to_dict()

        # Apply updates to Sheet1
        for row in range(2, ws1.max_row + 1):
            acc = str(ws1.cell(row=row, column=4).value).strip() if ws1.cell(row=row, column=4).value else ""
            k_cell = ws1.cell(row=row, column=11)
            l_cell = ws1.cell(row=row, column=12)

            if acc in zero_accounts:
                k_cell.value = 0
                l_cell.value = 0
                ws1.cell(row=row, column=insert_pos, value="")
            else:
                mapped_value = map_from_sheet2.get(acc, "")
                ws1.cell(row=row, column=insert_pos, value=mapped_value)

        # Auto-fit columns
        for col in ws1.columns:
            maxlen = max((len(str(c.value)) for c in col if c.value), default=0)
            ws1.column_dimensions[col[0].column_letter].width = maxlen + 2

        # Save updated file
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Mapping and aggregation completed successfully!")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
