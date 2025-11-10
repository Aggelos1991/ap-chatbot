import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Auto-Updater", layout="wide")
st.title("📊 Excel Auto-Updater (Minimal Version)")

uploaded = st.file_uploader("📁 Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        # Load entire workbook
        xls = pd.ExcelFile(uploaded)
        sheets = {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}

        # Work on first sheet (for update)
        sheet1_name = xls.sheet_names[0]
        sheet1 = sheets[sheet1_name]

        # Second sheet used for aggregation (data source)
        sheet2_name = xls.sheet_names[1]
        sheet2 = sheets[sheet2_name]

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

        # Identify key columns (by position)
        col_B = sheet2.columns[1]
        col_K = sheet2.columns[10]
        col_L = sheet2.columns[11]

        # Aggregate totals from sheet2
        grouped = (
            sheet2.groupby(col_B, dropna=False)[[col_K, col_L]]
            .sum()
            .reset_index()
        )
        grouped["Διάσταση 2 (Source)"] = grouped[col_B].map(mapping)

        # Insert new column next to L (only once)
        L_index = sheet1.columns.get_loc(sheet1.columns[11])
        if "Διάσταση 2 (Source)" not in sheet1.columns:
            sheet1.insert(L_index + 1, "Διάσταση 2 (Source)", "")

        # Apply updates
        for i, row in sheet1.iterrows():
            acc = str(row.iloc[4]).strip()
            if acc in zero_accounts:
                sheet1.at[i, sheet1.columns[10]] = 0     # Column K
                sheet1.at[i, sheet1.columns[11]] = 0     # Column L
            else:
                match = grouped.sample(1).iloc[0]
                sheet1.at[i, "Διάσταση 2 (Source)"] = match["Διάσταση 2 (Source)"]

        # Replace back into sheets dict
        sheets[sheet1_name] = sheet1

        # Save all sheets exactly as before
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=name)
        output.seek(0)

        st.success("✅ Excel successfully updated.")
        st.download_button(
            "⬇️ Download Updated File",
            data=output,
            file_name="Updated_" + uploaded.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
