import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Excel Mapping Updater", layout="wide")
st.title("📘 Excel Mapping Updater — Formatting Preserved")

st.write("""
Upload your Excel file (.xlsx) below.  
This app:
- Keeps all formatting intact  
- Removes nothing except what you specify  
- Inserts **Διάσταση 2 (Source)** right after the column **Πιστωτικό Υπόλοιπο** in Sheet1  
- Fills it using the mapping based on Διάσταση 2 values from Sheet2  
- Zeros out K & L for zero accounts (keeps Source blank)
""")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        # Load workbook while preserving formatting
        wb = load_workbook(uploaded)
        ws1 = wb.worksheets[0]  # Sheet1 (target)
        ws2 = wb.worksheets[1]  # Sheet2 (source)

        # === Zero accounts ===
        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
            "50.01.00.0000","50.01.01.0000","50.05.00.0000"
        ]

        # === Mapping ===
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

        # === Find the "Πιστωτικό Υπόλοιπο" column ===
        target_col = None
        for col in range(1, ws1.max_column + 1):
            if str(ws1.cell(row=1, column=col).value).strip() == "Πιστωτικό Υπόλοιπο":
                target_col = col
                break

        if not target_col:
            st.error("❌ Column 'Πιστωτικό Υπόλοιπο' not found in Sheet1.")
            st.stop()

        # Insert new column right after "Πιστωτικό Υπόλοιπο"
        insert_pos = target_col + 1
        ws1.insert_cols(insert_pos)
        ws1.cell(row=1, column=insert_pos, value="Διάσταση 2 (Source)")

        # === Extract Διάσταση 2 data from Sheet2 ===
        dim2_values = []
        for row in range(2, ws2.max_row + 1):
            value = ws2.cell(row=row, column=2).value  # column B in Sheet2
            if value:
                dim2_values.append(str(value).strip())

        dim2_values = list(dict.fromkeys(dim2_values))  # unique

        # === Update Sheet1 ===
        for row in range(2, ws1.max_row + 1):
            acc = str(ws1.cell(row=row, column=5).value).strip() if ws1.cell(row=row, column=5).value else ""
            k_cell = ws1.cell(row=row, column=11)
            l_cell = ws1.cell(row=row, column=12)

            if acc in zero_accounts:
                k_cell.value = 0
                l_cell.value = 0
                ws1.cell(row=row, column=insert_pos, value="")
            else:
                # If any Διάσταση 2 key from Sheet2 matches mapping
                found_key = next((v for v in dim2_values if v in mapping), None)
                mapped_val = mapping.get(found_key, "")
                ws1.cell(row=row, column=insert_pos, value=mapped_val)

        # === Preserve formatting + Auto column width ===
        for col in ws1.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws1.column_dimensions[col_letter].width = max_len + 2

        # === Save in memory ===
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ File updated — formatting preserved.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=output,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
