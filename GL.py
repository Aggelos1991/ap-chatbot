import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="Excel Formatter Preserver", layout="wide")
st.title("📘 Excel In-Place Processor (Formatting Preserved)")

st.write("""
Upload your Excel file (.xlsx).  
The app will:
- Preserve **all original formatting, fonts, and colors**
- **Remove column N** completely  
- Add **‘Διάσταση 2 (Source)’** right after column **L**
- Zero out **K + L** for zero accounts  
- Leave everything else exactly as in your file  
""")

uploaded = st.file_uploader("📁 Upload Excel", type=["xlsx"])

if uploaded:
    try:
        # Load workbook preserving formatting
        wb = load_workbook(uploaded)
        ws = wb.worksheets[0]  # First sheet (active one)
        
        # Define zero accounts
        zero_accounts = [
            "50.00.00.0000", "50.00.00.0001", "50.00.00.0002", "50.00.00.0003",
            "50.01.00.0000", "50.01.01.0000", "50.05.00.0000"
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

        # Step 1. Remove column N (14th column)
        if ws.max_column >= 14:
            ws.delete_cols(14)

        # Step 2. Insert “Διάσταση 2 (Source)” column after L (now column 12)
        insert_position = 13
        ws.insert_cols(insert_position)
        ws.cell(row=1, column=insert_position, value="Διάσταση 2 (Source)")

        # Step 3. Process rows
        for row in range(2, ws.max_row + 1):
            account = str(ws.cell(row=row, column=5).value).strip()  # Column E
            col_K, col_L = ws.cell(row=row, column=11), ws.cell(row=row, column=12)

            if account in zero_accounts:
                col_K.value = 0
                col_L.value = 0
                ws.cell(row=row, column=insert_position, value="")
            else:
                ws.cell(row=row, column=insert_position, value=mapping.get(account, ""))

        # Step 4. Auto column width
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        # Step 5. Save file to memory
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Excel updated — formatting preserved.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=output,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
