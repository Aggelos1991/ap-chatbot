import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime


# ---------- HELPERS ----------
def parse_amount(value):
    """Parses European or US formatted numbers to float."""
    if pd.isna(value):
        return None
    s = str(value).strip()
    s = re.sub(r'[^\d,.\-]', '', s)
    # handle formats like 1.234,56 or 1,234.56
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
    v = re.sub(r'[^A-Za-z0-9]', '', str(v)).lower()
    return v


def find_col(df, variants):
    for c in df.columns:
        if any(v.lower().replace(' ', '') in c.lower().replace(' ', '') for v in variants):
            return c
    return None


# ---------- STREAMLIT CONFIG ----------
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export & Email Tool")

# ---------- FILE UPLOADS ----------
uploaded_file = st.file_uploader("📂 Upload Payment Excel (TEST.xlsx)", type=["xlsx"])
credit_file = st.file_uploader("📂 Optional: Upload Credit Notes Excel", type=["xlsx"])

credit_df = None
if credit_file:
    try:
        credit_df = pd.read_excel(credit_file)
        credit_df.columns = [c.strip() for c in credit_df.columns]
        credit_df = credit_df.loc[:, ~credit_df.columns.duplicated()]
        st.success("✅ Credit Notes file loaded")
        st.write("Columns:", list(credit_df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Credit Notes Excel: {e}")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment Excel loaded")
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

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

    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            subset = subset.copy()
            subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount).fillna(0.0)
            vendor = subset["Supplier Name"].iloc[0]
            vendor_norm = norm_vendor(vendor)
            email_to = subset["Supplier's Email"].iloc[0]

            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            st.divider()
            payment_amount = st.text_input("💶 Actual Payment Amount (optional)", value="")
            payment_value = parse_amount(payment_amount)

            # --- CN DETECTION ---
            if credit_df is not None:
                st.info("🔎 Checking for Credit Notes...")

                alt_col = find_col(credit_df, ["Alt.Document", "Alt. Document", "Document"])
                val_col = find_col(credit_df, ["Invoice Value", "Amount", "Value"])
                vendor_col = find_col(credit_df, ["Supplier Name", "Vendor", "Supplier"])

                if not alt_col or not val_col:
                    st.warning("⚠️ CN file missing 'Alt.Document' or 'Amount' column.")
                else:
                    cn = credit_df.copy()
                    cn[val_col] = cn[val_col].apply(parse_amount)
                    cn = cn.dropna(subset=[val_col])
                    cn["VendorNorm"] = cn[vendor_col].apply(norm_vendor) if vendor_col else ""
                    cn["AbsVal"] = cn[val_col].abs().round(2)

                    total_invoices = round(summary["Invoice Value"].sum(), 2)
                    if payment_value:
                        diff = round(total_invoices - payment_value, 2)
                        st.caption(f"Invoices total €{total_invoices:.2f} | Payment €{payment_value:.2f} | Diff €{diff:.2f}")
                        target = abs(diff)

                        if target > 0.01:
                            vendor_cn = cn[cn["VendorNorm"] == vendor_norm]
                            matches = vendor_cn[vendor_cn["AbsVal"].round(2) == round(target, 2)]

                            if matches.empty:
                                # fallback: match all vendors
                                matches = cn[cn["AbsVal"].round(2) == round(target, 2)]

                            if not matches.empty:
                                st.write("🔍 CN candidates found:")
                                st.dataframe(matches[[alt_col, val_col]])

                                # pick LAST if multiple
                                chosen = matches.iloc[-1]
                                cn_alt = str(chosen[alt_col])
                                cn_val = -abs(chosen[val_col])  # deduct
                                summary = pd.concat([
                                    summary,
                                    pd.DataFrame([{"Alt. Document": f"{cn_alt} (CN)", "Invoice Value": cn_val}])
                                ], ignore_index=True)

                                if len(matches) > 1:
                                    st.warning(
                                        f"⚠️ Found {len(matches)} CNs with same amount €{target:.2f}. "
                                        f"Applied LAST: {cn_alt}."
                                    )
                                else:
                                    st.success(f"✅ Applied CN '{cn_alt}' (deducted €{abs(cn_val):.2f}).")
                            else:
                                st.info("ℹ️ No Credit Note matches the exact difference.")
                        else:
                            st.info("No difference detected — no CN needed.")

            # --- FINAL TOTAL ---
            total_value = summary["Invoice Value"].sum()
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            # --- DISPLAY ---
            summary["Invoice Value"] = summary["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Vendor Email (from Excel):** {email_to}")
            st.dataframe(summary)

            # --- EXCEL EXPORT ---
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
