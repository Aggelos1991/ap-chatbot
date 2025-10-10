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

# ====================== Helpers ======================
def parse_number_to_cents(x):
    """Parse amounts like '1.234,56' / '1,234.56' / 167,09 / 167.09 -> int cents."""
    if pd.isna(x): return None
    s = str(x).strip()
    s = re.sub(r'[^\d,.\-\+]', '', s)
    if s.count(',') == 1 and (s.count('.') == 0 or s.find(',') > s.find('.')):
        s = s.replace('.', '')
        s = s.replace(',', '.')
    elif s.count(',') >= 1 and s.count('.') >= 1:
        s = s.replace(',', '')
    try:
        val = float(s)
    except:
        return None
    return int(round(val * 100))

def norm_vendor(s):
    if pd.isna(s): return ""
    s = str(s).upper().strip()
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def find_col(cols, candidates):
    norm = {c: c.strip().replace(" ", "").lower() for c in cols}
    for cand in candidates:
        target = cand.replace(" ", "").lower()
        for orig, n in norm.items():
            if n == target:
                return orig
    return None

# ================== Streamlit Config ==================
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export & Email Tool")

# ------------------- File Uploads --------------------
uploaded_file = st.file_uploader("📂 Upload Payment Excel (TEST.xlsx)", type=["xlsx"])
credit_file   = st.file_uploader("📂 Optional: Upload Credit Notes Excel", type=["xlsx"])

credit_df = None
if credit_file:
    try:
        credit_df = pd.read_excel(credit_file)
        credit_df.columns = [str(c).strip() for c in credit_df.columns]
        credit_df = credit_df.loc[:, ~credit_df.columns.duplicated()]
        st.success("✅ Credit Notes file loaded successfully")
        st.write("Credit Notes Columns detected:", list(credit_df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Credit Notes Excel: {e}")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment Excel loaded successfully")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # Required columns
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
            subset["Invoice Value (cents)"] = subset["Invoice Value"].apply(parse_number_to_cents).fillna(0).astype(int)
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value (cents)"].sum()

            vendor_raw = str(subset["Supplier Name"].dropna().iloc[0])
            vendor_norm = norm_vendor(vendor_raw)
            email_to = str(subset["Supplier's Email"].dropna().iloc[0])

            st.divider()
            payment_amount = st.text_input("💶 Actual Payment Amount (optional)", value="")
            payment_cents = parse_number_to_cents(payment_amount) if payment_amount.strip() else None

            # ---- CN handling ----
            applied_cn = None
            extra_cn_alts = []

            if credit_df is not None:
                st.info("🔎 Checking for Credit Notes...")

                alt_col = find_col(credit_df.columns, ["Alt.Document", "Alt. Document", "AltDoc", "Document"])
                val_col = find_col(credit_df.columns, ["Invoice Value", "Amount", "Value"])
                sup_col = find_col(credit_df.columns, ["Supplier Name", "Vendor", "Supplier"])

                if alt_col and val_col:
                    cn = credit_df.copy()
                    cn["__cents__"] = cn[val_col].apply(parse_number_to_cents)
                    cn = cn.dropna(subset=["__cents__"]).copy()
                    cn["__abs_cents__"] = cn["__cents__"].abs().astype(int)

                    if sup_col:
                        cn["__VENDOR_NORM__"] = cn[sup_col].apply(norm_vendor)
                    else:
                        cn["__VENDOR_NORM__"] = ""

                    total_inv_cents = int(summary["Invoice Value (cents)"].sum())

                    if payment_cents is not None and payment_cents >= 0:
                        diff_cents = total_inv_cents - int(payment_cents)
                        target_cents = abs(diff_cents)

                        st.caption(
                            f"• Invoices total: €{total_inv_cents/100:.2f} | Payment: €{(payment_cents or 0)/100:.2f} | "
                            f"Diff: €{diff_cents/100:.2f} (target CN = €{target_cents/100:.2f})"
                        )

                        if target_cents > 0:
                            # Match CNs (strict same vendor)
                            matches = cn[(cn["__VENDOR_NORM__"] == vendor_norm) & (cn["__abs_cents__"] == target_cents)]

                            # If none found, try loose vendor match
                            if matches.empty and sup_col:
                                matches = cn[cn["__abs_cents__"] == target_cents]

                            if not matches.empty:
                                matches = matches.sort_values(by=alt_col.astype(str) if alt_col in matches.columns else "__abs_cents__")
                                last = matches.iloc[-1]  # take the last one
                                applied_cn = {
                                    "alt": str(last[alt_col]),
                                    "cents": -target_cents  # deduct
                                }

                                if len(matches) > 1:
                                    extra_cn_alts = [str(a) for a in matches[alt_col].astype(str).tolist()[:-1]]

                                # Append CN as negative line
                                cn_row = pd.DataFrame([{
                                    "Alt. Document": f"{applied_cn['alt']} (CN)",
                                    "Invoice Value (cents)": applied_cn["cents"]
                                }])
                                summary = pd.concat([summary, cn_row], ignore_index=True)

                                st.success(f"✅ Applied CN '{applied_cn['alt']}' and deducted €{target_cents/100:.2f}.")
                                if extra_cn_alts:
                                    st.warning(f"⚠️ Found {len(extra_cn_alts)+1} CNs with the same amount — only the LAST one applied. Others: {', '.join(extra_cn_alts)}")
                            else:
                                st.info("ℹ️ No Credit Note matches the exact difference. No CN applied.")
                        else:
                            st.info("No difference detected — no CN needed.")
                else:
                    st.warning("⚠️ Credit Notes file missing recognizable columns (expected e.g. 'Alt.Document' and 'Amount').")

            # ---------- Final total ----------
            total_value_cents = int(summary["Invoice Value (cents)"].sum())
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value (cents)": total_value_cents}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            show = summary.copy()
            show["Invoice Value (€)"] = show["Invoice Value (cents)"].apply(lambda c: f"€{c/100:,.2f}")
            show = show[["Alt. Document", "Invoice Value (€)"]]

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor_raw}")
            st.write(f"**Vendor Email (from Excel):** {email_to}")
            st.dataframe(show)

            # ---------- Excel Export ----------
            wb = Workbook()
            ws_summary = wb.active
            ws_summary.title = "Summary"
            for r in dataframe_to_rows(show, index=False, header=True):
                ws_summary.append(r)

            ws_hidden = wb.create_sheet("HiddenMeta")
            ws_hidden["A1"], ws_hidden["B1"] = "Vendor", vendor_raw
            ws_hidden["A2"], ws_hidden["B2"] = "Vendor Email", email_to
            ws_hidden["A3"], ws_hidden["B3"] = "Payment Code", pay_code
            ws_hidden["A4"], ws_hidden["B4"] = "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws_hidden.sheet_state = "hidden"

            folder_path = os.path.join(os.getcwd(), "exports")
            os.makedirs(folder_path, exist_ok=True)
            file_path = os.path.join(folder_path, f"{norm_vendor(vendor_raw)}_Payment_{pay_code}.xlsx")
            wb.save(file_path)

            # ---------- Gmail Section ----------
            st.divider()
            st.subheader("📨 Send Excel by Gmail")

            sender_email = st.text_input("Your Gmail address:")
            app_password = st.text_input("Your Gmail App Password:", type="password")

            subject = f"Payment Summary — {vendor_raw}"
            body = f"Dear {vendor_raw},\n\nPlease find attached the payment summary for code {pay_code}.\n\nKind regards,\nAngelos"

            if st.button("✉️ Send Email"):
                try:
                    msg = MIMEMultipart()
                    msg["From"] = sender_email
                    msg["To"] = email_to
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))

                    with open(file_path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
                    msg.attach(part)

                    server = smtplib.SMTP("smtp.gmail.com", 587)
                    server.starttls()
                    server.login(sender_email, app_password)
                    server.send_message(msg)
                    server.quit()

                    st.success(f"✅ Email sent successfully to {email_to}")
                except Exception as e:
                    st.error(f"❌ Failed to send email: {e}")

            # ---------- Download ----------
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="💾 Download Excel Summary (with hidden email tab)",
                data=buffer,
                file_name=f"{norm_vendor(vendor_raw)}_Payment_{pay_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Upload your Excel file to begin.")
