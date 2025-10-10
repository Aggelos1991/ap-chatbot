import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="💼 Vendor Payment Reconciliation & Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    st.success("✅ Excel file loaded successfully!")
    st.write("### 🧭 Columns detected in your Excel:")
    st.dataframe(pd.DataFrame(df.columns, columns=["Columns"]))

    # === Correct column mapping based on your file ===
    col_map = {
        "Payment Document Code": None,
        "Alt.Document": None,
        "Invoice Value": None,
        "Supplier Name": None,
        "Supplier's Email": None
    }

    missing_cols = [col for col in col_map if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing columns in your Excel: {missing_cols}")
        st.stop()

    payment_code = st.text_input("🔎 Enter Payment Document Code:")

    if payment_code:
        subset = df[df["Payment Document Code"].astype(str).str.strip() == str(payment_code).strip()]
        if subset.empty:
            st.warning("⚠️ No records found for this Payment Document Code.")
        else:
            # Pivot / summary
            summary = subset.groupby("Alt.Document", as_index=False)["Invoice Value"].sum()
            total = summary["Invoice Value"].sum()
            vendor = subset["Supplier Name"].iloc[0]
            email = subset["Supplier's Email"].iloc[0]

            st.write("### 🧾 Invoices related to this Payment Document Code:")
            st.dataframe(summary)
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Amount:** €{total:,.2f}")

            invoice_lines = "\n".join(
                f"- {row['Alt.Document']}: €{row['Invoice Value']:,.2f}" for _, row in summary.iterrows()
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
            # Email via Outlook (SMTP)
            # ==============================
            st.write("---")
            st.write("### ✉️ Email Settings (SMTP for Outlook)")
            smtp_server = st.text_input("SMTP server", "smtp.office365.com")
            smtp_port = st.number_input("SMTP port", 1, 9999, 587)
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
