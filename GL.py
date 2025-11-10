import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Aggregator & Mapper", layout="wide")
st.title("📘 Excel Aggregator & Mapper — Robust Διάσταση 2 → Τίτλος Mapping")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]  # Sheet 1
        ws2 = wb.worksheets[1]  # Sheet 2

        # ======= CONFIG =======
        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002",
            "50.00.00.0003","50.01.00.0000","50.01.01.0000","50.05.00.0000"
        ]

        # ======= STEP 1 — DELETE COLUMN E =======
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Λογαριασμός λογιστικής":
                ws1.delete_cols(col)
                break

        # ======= STEP 2 — FIND "Πιστωτικό Υπόλοιπο" =======
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

        # ======= STEP 3 — READ SHEET2 INTO DATAFRAME =======
        df2 = pd.DataFrame(ws2.values)
        df2.columns = df2.iloc[0]
        df2 = df2.drop(0)

        st.write("📄 Detected columns in Sheet2:", list(df2.columns))

        # Try to locate “Διάσταση 2” and “Τίτλος” dynamically
        col_dim2 = next((c for c in df2.columns if "Διάσταση" in str(c)), None)
        col_titlos = next((c for c in df2.columns if "Τίτλος" in str(c)), None)

        if not col_dim2 or not col_titlos:
            st.error(f"❌ Could not find both 'Διάσταση 2' and 'Τίτλος' columns.\nFound Διάσταση2={col_dim2}, Τίτλος={col_titlos}")
            st.stop()

        # Find K and L dynamically by position or partial match
        col_K = df2.columns[10] if len(df2.columns) > 10 else df2.columns[-2]
        col_L = df2.columns[11] if len(df2.columns) > 11 else df2.columns[-1]

        # Convert to numeric
        df2[col_K] = pd.to_numeric(df2[col_K], errors="coerce").fillna(0)
        df2[col_L] = pd.to_numeric(df2[col_L], errors="coerce").fillna(0)

        # ======= STEP 4 — AGGREGATE =======
        agg = df2.groupby(col_dim2)[[col_K, col_L]].sum().reset_index()
        map_from_sheet2 = df2.set_index(col_dim2)[col_titlos].to_dict()

        # ======= STEP 5 — UPDATE SHEET1 =======
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

        # ======= STEP 6 — AUTO WIDTHS =======
        for col in ws1.columns:
            maxlen = max((len(str(c.value)) for c in col if c.value), default=0)
            ws1.column_dimensions[col[0].column_letter].width = maxlen + 2

        # ======= SAVE =======
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        st.success("✅ Mapping and aggregation successfully applied!")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=out,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
