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

# ===== Streamlit config =====
st.set_page_config(page_title="💼 Vendor Payment Reconciliation Exporter", layout="wide")
st.title("💼 Vendor Payment Reconciliation — Excel Export & Email Tool")

# --- FILE UPLOADS ---
uploaded_file = st.file_uploader("📂 Upload Payment Excel (TEST.xlsx)", type=["xlsx"])
credit_file = st.file_uploader("📂 Optional: Upload Credit Notes Excel", type=["xlsx"])

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

    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            subset = subset.copy()
            subset["Invoice Value"] = pd.to_numeric(subset["Invoice Value"], errors="coerce").fillna(0)
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()

            vendor = str(subset["Supplier Name"].dropna().iloc[0])
            email_to = str(subset["Supplier's Email"].dropna().iloc[0])

            st.divider()
            payment_amount = st.number_input("💶 Actual Payment Amount (optional)", min_value=0.0, step=0.01, format="%.2f")

            # === Handle Credit Notes if file uploaded ===
            if credit_df is not None:
                st.info("🔎 Checking for Credit Notes...")

                # --- Flexible column mapping for Credit Notes ---
                alt_col = next(
                    (c for c in credit_df.columns if c.strip().replace(" ", "").lower() in ["alt.document", "altdocument"]),
                    None,
                )
                val_col = next(
                    (c for c in credit_df.columns if c.strip().replace(" ", "").lower() in ["invoicevalue", "amount", "value"]),
                    None,
                )

                if alt_col and val_col:
                    credit_df[val_col] = pd.to_numeric(credit_df[val_col], errors="coerce").fillna(0)

                    # Match vendor if column exists
                    if "Supplier Name" in credit_df.columns:
                        vendor_cn_df = credit_df[credit_df["Supplier Name"].astype(str) == vendor].copy()
                    else:
                        vendor_cn_df = credit_df.copy()

                    total_invoices = float(summary["Invoice Value"].sum())
                    chosen_cn_added = False

                    if payment_amount and payment_amount > 0:
                        diff = round(total_invoices - float(payment_amount), 2)
                        if abs(diff) > 0.01:
                            tmp = vendor_cn_df.copy()
                            tmp["_abs_val"] = tmp[val_col].abs().round(2)
                            target = round(abs(diff), 2)
                            matches = tmp[tmp["_abs_val"] == target]

                            if not matches.empty:
                                if len(matches) > 1:
                                    st.warning(
                                        f"⚠️ Found {len(matches)} Credit Notes with the same value €{target:.2f}. "
                                        "Only the first one will be applied."
                                    )

                                cn_row = matches.iloc[0]
                                cn_alt = str(cn_row[alt_col])
                                cn_value = -target
                                cn_df = pd.DataFrame(
                                    [{"Alt. Document": f"{cn_alt} (CN)", "Invoice Value": cn_value}]
                                )
                                summary = pd.concat([summary, cn_df], ignore_index=True)
                                chosen_cn_added = True
                                st.success(f"✅ Applied CN '{cn_alt}' for {vendor}: deducted €{target:.2f}.")
                            else:
                                st.info("ℹ️ No Credit Note matches the exact difference. No CN applied.")
                        else:
                            st.info("No difference detected between invoices and payment — no CN needed.")

                    # If user did NOT enter payment amount → fallback (optional)
                    if not chosen_cn_added and (not payment_amount or payment_amount == 0):
                        st.info("No payment amount entered — skipping CN deduction logic.")

                else:
                    st.warning("⚠️ Credit Notes file missing recognizable columns (expected something like 'Alt.Document' and 'Amount').")

            # === Total row (includes CN deduction if any) ===
            total_value = summary["Invoice Value"].sum()
            total_row = pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_value}])
            summary = pd.concat([summary, total_row], ignore_index=True)

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Vendor Email (from Excel):** {email_to}")
            st.dataframe(summary.style.format({"Invoice Value": "€{:,.2f}".format}))

            # --- Create workbook ---
            wb = Workbook()
            ws_summary = wb.active
            ws_summary.title = "Summary"
            for r in dataframe_to_rows(summary, index=False, header=True):
                ws_summary.append(r)

            # Hidden metadata
            ws_hidden = wb.create_sheet("HiddenMeta")
            ws_hidden["A1"], ws_hidden["B1"] = "Vendor", vendor
            ws_hidden["A2"], ws_hidden["B2"] = "Vendor Email", email_to
            ws_hidden["A3"], ws_hidden["B3"] = "Payment Code", pay_code
            ws_hidden["A4"], ws_hidden["B4"] = "Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws_hidden.sheet_state = "hidden"

            # Prepare temp folder
            folder_path = os.path.join(os.getcwd(), "exports")
            os.makedirs(folder_path, exist_ok=True)
            file_path = os.path.join(folder_path, f"{vendor}_Payment_{pay_code}.xlsx")
            wb.save(file_path)

            # === EMAIL SECTION ===
            st.divider()
            st.subheader("📨 Send Excel by Gmail")

            sender_email = st.text_input("Your Gmail address:")
            app_password = st.text_input("Your Gmail App Password:", type="password")

            subject = f"Payment Summary — {vendor}"
            body = f"Dear {vendor},\n\nPlease find attached the payment summary for code {pay_code}.\n\nKind regards,\nAngelos"

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

            # === DOWNLOAD SECTION ===
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="💾 Download Excel Summary (with hidden email tab)",
                data=buffer,
                file_name=f"{vendor}_Payment_{pay_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Upload your Excel file to begin.")
