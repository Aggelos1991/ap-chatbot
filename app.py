import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

# ====== Helpers ======
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

def norm_vendor(v):
    if pd.isna(v):
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', str(v)).lower()

def find_col(df, variants):
    for c in df.columns:
        if any(v.lower().replace(' ', '') in c.lower().replace(' ', '') for v in variants):
            return c
    return None


# ====== Streamlit Config ======
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export & Email Tool")

# ====== File Uploads ======
uploaded_file = st.file_uploader("📂 Upload Payment Excel (TEST.xlsx)", type=["xlsx"])
credit_file = st.file_uploader("📂 Optional: Upload Credit Notes Excel", type=["xlsx"])

credit_df = None
if credit_file:
    try:
        credit_df = pd.read_excel(credit_file)
        credit_df.columns = [c.strip() for c in credit_df.columns]
        credit_df = credit_df.loc[:, ~credit_df.columns.duplicated()]
    except Exception as e:
        st.error(f"❌ Error loading Credit Notes Excel: {e}")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment Excel loaded successfully")
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # Required columns
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
            subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount).fillna(0.0)
            vendor = subset["Supplier Name"].iloc[0]
            vendor_norm = norm_vendor(vendor)
            email_to = subset["Supplier's Email"].iloc[0]
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            payment_input = st.text_input("💶 Actual Payment Amount (optional)")
            payment_value = parse_amount(payment_input)

            # === Automatic CN logic (no messages, silent) ===
            if credit_df is not None and payment_value:
                alt_col = find_col(credit_df, ["Alt.Document", "Alt. Document", "Document"])
                val_col = find_col(credit_df, ["Invoice Value", "Amount", "Value"])
                vendor_col = find_col(credit_df, ["Supplier Name", "Vendor", "Supplier"])

                if alt_col and val_col:
                    cn = credit_df.copy()
                    cn[val_col] = cn[val_col].apply(parse_amount)
                    cn = cn.dropna(subset=[val_col])
                    cn["AbsVal"] = cn[val_col].abs().round(2)
                    cn["VendorNorm"] = cn[vendor_col].apply(norm_vendor) if vendor_col else ""

                    total_invoices = round(summary["Invoice Value"].sum(), 2)
                    diff = round(total_invoices - payment_value, 2)
                    target = abs(diff)

                    # Match by vendor & value
                    matches = cn[(cn["VendorNorm"] == vendor_norm) & (cn["AbsVal"] == target)]
                    if matches.empty:
                        matches = cn[cn["AbsVal"] == target]

                    if not matches.empty:
                        last = matches.iloc[-1]  # apply last found CN
                        cn_alt = str(last[alt_col])
                        cn_val = -abs(last[val_col])  # deduct
                        summary = pd.concat([
                            summary,
                            pd.DataFrame([{"Alt. Document": f"{cn_alt} (CN)", "Invoice Value": cn_val}])
                        ], ignore_index=True)

            # === Final total ===
            total_value = summary["Invoice Value"].sum()
            summary = pd.concat([
                summary,
                pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            ], ignore_index=True)

            summary["Invoice Value"] = summary["Invoice Value"].apply(lambda v: f"€{v:,.2f}")

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Vendor Email (from Excel):** {email_to}")
            st.dataframe(summary)

            # === Excel Export ===
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
            file_path = os.path.join(folder_path, f"{vendor_norm}_Payment_{pay_code}.xlsx")
            wb.save(file_path)

            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            st.download_button(
                "💾 Download Excel Summary",
                buffer,
                file_name=f"{vendor_norm}_Payment_{pay_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Upload your Excel file to begin.")
