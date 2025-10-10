import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re

st.set_page_config(page_title="💼 Vendor Payment Reconciliation & Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # normalize headers
    df.columns = [
        re.sub(r'[^a-z0-9]+', ' ', c.lower()).strip()
        for c in df.columns
    ]
    df = df.loc[:, ~df.columns.duplicated()]

    st.success("✅ Excel file loaded successfully!")
    st.write("### 🧭 Normalized column headers:")
    st.dataframe(pd.DataFrame(df.columns, columns=["Normalized Header"]))

    # Flexible column detection
    col_payment = [c for c in df.columns if "payment" in c and "code" in c]
    col_invoice = [c for c in df.columns if "alt" in c and "document" in c]
    col_amount = [c for c in df.columns if "invoice" in c and "value" in c]
    col_vendor = [c for c in df.columns if "supplier" in c and "name" in c]
    col_email = [c for c in df.columns if "email" in c]

    if not all([col_payment, col_invoice, col_amount, col_vendor, col_email]):
        st.error(f"❌ Some required columns are missing. Detected:\n"
                 f"Payment Code: {col_payment}\nAlt.Document: {col_invoice}\n"
                 f"Invoice Value: {col_amount}\nSupplier Name: {col_vendor}\nSupplier Email: {col_email}")
        st.stop()

    # map detected names
    col_payment, col_invoice, col_amount, col_vendor, col_email = (
        col_payment[0], col_invoice[0], col_amount[0], col_vendor[0], col_email[0]
    )

    payment_code = st.text_input("🔎 Enter Payment Document Code:")

    if payment_code:
        subset = df[df[col_payment].astype(str).str.strip() == str(payment_code).strip()]

        if subset.empty:
            st.warning("⚠️ No records found for this Payment Document Code.")
        else:
            # group invoices
            summary = subset.groupby(col_invoice, as_index=False)[col_amount].sum()
            total = summary[col_amount].sum()
            vendor = subset[col_vendor].iloc[0]
            email = subset[col_email].iloc[0]

            st.write("### 🧾 Invoices related to this Payment Document Code:")
            st.dataframe(summary)
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Amount:** €{total:,.2f}")

            invoice_lines = "\n".join(
                f"- {row[col_invoice]}: €{row[col_amount]:,.2f}" for _, row in summary.iterrows()
            )

            email_body = f"""
Dear {vendor},

Please find below the invoices corresponding to the payment we made under payment document code {payment_code}.

{invoice_lines}

Total amount: €{total:,.2f}

Thank you for your cooperation.

Kind regards,
Angelos Keramaris
Accounts Payable Department
Ikos Resorts
"""

            st.text_area("📧 Email draft:", email_body, height=250)

            # ==============================
            # SMTP send (works anywhere)
            # ==============================
            st.write("---")
            st.write("### ✉️ Email Settings (Outlook SMTP)")
            smtp_server = st.text_input("SMTP server", "smtp.office365.com")
            smtp_port = st.number_input("SMTP port", min_value=1, max_value=9999, value=587)
            smtp_user = st.text_input("Your Outlook email address:")
            smtp_pass = st.text_input("Your Outlook password or app password:", type="password")

            if st.button("📨 Send Email"):
                try:
                    msg = MIMEMultipart()
                    msg["From"] = smtp_user
                    msg["To"] = email
                    msg["Subject"] = f"Payment details — Document {payment_code}"
                    msg.attach(MIMEText(email_body, "plain"))

                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls()
                        server.login(smtp_user, smtp_pass)
                        server.send_message(msg)

                    st.success(f"✅ Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"❌ Failed to send email: {e}")
