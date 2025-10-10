import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

# ===== Helpers =====
def parse_amount(value):
    """Convert European/US numeric formats to float."""
    if pd.isna(value):
        return None
    s = str(value).strip()
    s = re.sub(r'[^\d,.\-]', '', s)
    if s.count(',') == 1 and s.count('.') == 1:
        if s.find(',') > s.find('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif s.count(',') == 1:
        s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return None


def find_col(df, variants):
    for c in df.columns:
        if any(v.lower().replace(' ', '') in c.lower().replace(' ', '') for v in variants):
            return c
    return None


# ===== Streamlit Config =====
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Auto CN Deduction Tool")

# ===== File Uploads =====
uploaded_file = st.file_uploader("📂 Upload Payment Excel (TEST.xlsx)", type=["xlsx"])
credit_file = st.file_uploader("📂 Upload Credit Notes Excel", type=["xlsx"])

if uploaded_file and credit_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment Excel loaded successfully")

        credit_df = pd.read_excel(credit_file)
        credit_df.columns = [c.strip() for c in credit_df.columns]
        credit_df = credit_df.loc[:, ~credit_df.columns.duplicated()]
        st.success("✅ Credit Notes Excel loaded successfully")

    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # --- Required columns ---
    REQ = ["Payment Document Code", "Alt. Document", "Invoice Value", "Supplier Name", "Supplier's Email"]
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            subset = subset.copy()
            subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount).fillna(0.0)

            vendor = subset["Supplier Name"].iloc[0]
            email_to = subset["Supplier's Email"].iloc[0]

            # --- Credit Notes columns ---
            alt_col = find_col(credit_df, ["Alt.Document", "Alt. Document", "Document"])
            val_col = find_col(credit_df, ["Invoice Value", "Amount", "Value"])

            credit_df[val_col] = credit_df[val_col].apply(parse_amount).fillna(0.0)

            # --- Summary creation ---
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            cn_rows = []

            # Loop through each Alt. Document in main file
            for alt_doc in summary["Alt. Document"]:
                cn_match = credit_df[credit_df[alt_col].astype(str) == str(alt_doc)]
                if not cn_match.empty:
                    # If multiple CNs for same doc, take last one
                    last_cn = cn_match.iloc[-1]
                    cn_val = -abs(last_cn[val_col])
                    cn_rows.append({"Alt. Document": f"{alt_doc} (CN)", "Invoice Value": cn_val})

            # Append CNs to summary
            if cn_rows:
                summary = pd.concat([summary, pd.DataFrame(cn_rows)], ignore_index=True)

            # --- Final total ---
            total_value = summary["Invoice Value"].sum()
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            # --- Display ---
            summary["Invoice Value"] = summary["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Vendor Email (from Excel):** {email_to}")
            st.dataframe(summary)

            # --- Excel Export ---
            wb = Workbook()
            ws = wb.active
            ws.title = "Summary"
            for r in dataframe_to_rows(summary, index=False, header=True):
                ws.append(r)

            ws_hidden = wb.create_sheet("HiddenMeta")
            ws_hidden["A1"], ws_hidden["B1"] = "Vendor", vendor
            ws_hidden["A2"], ws_hidden["B2"] = "Vendor Email", email_to
            ws_hidden["A3"], ws_hidden["B3"] = "Payment Code", pay_code
            ws_hidden["A4"], ws_hidden["B4"] = "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws_hidden.sheet_state = "hidden"

            folder_path = os.path.join(os.getcwd(), "exports")
            os.makedirs(folder_path, exist_ok=True)
            file_path = os.path.join(folder_path, f"{vendor}_Payment_{pay_code}.xlsx")
            wb.save(file_path)

            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                "💾 Download Excel Summary",
                buffer,
                file_name=f"{vendor}_Payment_{pay_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Please upload both Payment and Credit Notes Excel files to begin.")
