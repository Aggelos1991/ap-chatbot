import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

# ===== Streamlit config =====
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export Tool")

uploaded_file = st.file_uploader("📂 Upload Excel (TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Excel loaded successfully")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # --- REQUIRED columns ---
    REQ = [
        "Payment Document Code",
        "Alt. Document",
        "Invoice Value",
        "Supplier Name",
        "Supplier's Email",
    ]
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.stop()

    # --- Filter by Payment Code ---
    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            subset = subset.copy()
            subset["Invoice Value"] = pd.to_numeric(subset["Invoice Value"], errors="coerce").fillna(0)
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            total_value = summary["Invoice Value"].sum()
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            vendor = str(subset["Supplier Name"].dropna().iloc[0])
            email_to = str(subset["Supplier's Email"].dropna().iloc[0])

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email_to}")
            st.dataframe(summary.style.format({"Invoice Value": "€{:,.2f}".format}))

            # --- Create workbook and hide email sheet ---
            wb = Workbook()
            ws_summary = wb.active
            ws_summary.title = "Summary"
            for r in dataframe_to_rows(summary, index=False, header=True):
                ws_summary.append(r)

            ws_hidden = wb.create_sheet("HiddenMeta")
            ws_hidden["A1"] = "Email"
            ws_hidden["B1"] = email_to
            ws_hidden.sheet_state = "hidden"

            # --- Save directly to Mac Desktop folder ---
            folder_path = "/Users/angeloskeramaris/Desktop/Payment Remmitance"
            os.makedirs(folder_path, exist_ok=True)
            file_path = os.path.join(folder_path, f"{vendor}_Payment_{pay_code}.xlsx")

            try:
                wb.save(file_path)
                st.success(f"✅ File saved directly to Desktop: {file_path}")
            except PermissionError:
                st.error("❌ Permission denied. Run Streamlit locally on your Mac (not on Streamlit Cloud).")

            # --- Optional: still allow manual download ---
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="💾 Download Excel Summary",
                data=buffer,
                file_name=f"{vendor}_Payment_{pay_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Upload your Excel file to begin.")
