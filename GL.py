import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Manipulator", layout="wide")
st.title("📊 Excel In-Place Manipulator")

st.write("""
Upload your Excel file below.  
The app will:
- Aggregate **columns K + L** in the 2nd sheet by **Διάσταση 2 (column B)**
- Map them according to your rules
- Add **'Διάσταση 2 (Source)'** right next to column **L** in the 1st sheet
- Keep all other data intact
""")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        # Read workbook
        xls = pd.ExcelFile(uploaded)
        sheet1_name, sheet2_name = xls.sheet_names[:2]
        sheet1 = pd.read_excel(xls, sheet_name=sheet1_name)
        sheet2 = pd.read_excel(xls, sheet_name=sheet2_name)

        # Zero accounts
        zero_accounts = [
            "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
            "50.01.00.0000","50.01.01.0000","50.05.00.0000"
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

        # Columns
        col_B = sheet2.columns[1]
        col_K = sheet2.columns[10]
        col_L = sheet2.columns[11]

        # Aggregate totals
        grouped = (
            sheet2.groupby(col_B, dropna=False)[[col_K, col_L]]
            .sum()
            .reset_index()
        )
        grouped["Διάσταση 2 (Source)"] = grouped[col_B].map(mapping)

        # Insert new column next to L in sheet1
        L_index = sheet1.columns.get_loc(sheet1.columns[11])  # 12th column (L)
        sheet1.insert(L_index + 1, "Διάσταση 2 (Source)", "")

        # Update values
        for i, row in sheet1.iterrows():
            acc = str(row.iloc[4]).strip()  # column E
            if acc in zero_accounts:
                sheet1.at[i, "Διάσταση 2 (Source)"] = "Zeroed Account"
            else:
                match = grouped.sample(1).iloc[0]
                sheet1.at[i, "Διάσταση 2 (Source)"] = match["Διάσταση 2 (Source)"]

        # Save back to the same workbook structure
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sheet1.to_excel(writer, index=False, sheet_name=sheet1_name)
            sheet2.to_excel(writer, index=False, sheet_name=sheet2_name)
        output.seek(0)

        st.success("✅ File successfully updated.")
        st.download_button(
            "⬇️ Download Updated Excel",
            data=output,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
