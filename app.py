import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ====================== CONFIG ======================
st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Sender")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]   # Clean header names
        df = df.loc[:, ~df.columns.duplicated()]            # Remove duplicates
        st.success("✅ Excel file loaded successfully!")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {e}")
        st.stop()

    # ====================== REQUIRED COLUMNS ======================
    required_cols = [
        "Payment Document Code",
        "Alt. Document",
        "Invoice Value",
        "Supplier's Email",
        "Supplier Name"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"⚠️ Missing columns: {missing}. Please correct your Excel and re-upload.")
        st.stop()

    # ====================== USER INPUT ======================
    payment_code = st.text_input("🔎 Enter Payment Document Code:")

    if payment_code:
        subset = df[df["Payment Document Code"].astype(str) == str(payment_code)]

        if subset.empty:
            st.warning("⚠️ No records found for this Payment Document Code.")
        else:
            # ====================== SUMMARY ======================
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()
            total = summary["Invoice Value"].sum()
            supplier = subset["Supplier Name"].iloc[0]
            email = subset["Supplier's Email"].iloc[0]

            st.subheader(f"📋 Summary for Payment Document Code: {payment_code}")
            st.dataframe(summary.style.format({"Invoice Value": "€{:,.2f}".format}))
            st.write(f"**Supplier:** {supplier}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

            # ====================== EMAIL BODY ======================
            invoice_lines = "\n".join(
                f"- {row['Alt. Document']}: €{row['Invoice Value']:,.2f}"
                for _, row in summary.iterrows()
            )

            email_body = f"""
Dear {supplier},

Please find below the invoices corresponding to the payment we made under payment document code {payment_code}.

{invoice_lines}

Total amount: €{total:,.2f}

Thank you for your cooperation.

Kind regards,
Angelos Keramaris
Accounts Payable Department
Ikos Resorts
"""
            st.divider()
            st.subheader("✉️ Email Preview")
            st.text_area("Generated Email:", email_body.strip(), height=250)

            # ====================== EMAIL SENDER ======================
            st.divider()
            st.subheader("📧 Send Email")

            sender_email = st.text_input("Your Outlook 365 email:")
            sender_pass = st.text_input("Your Outlook App Password (or leave blank if IT handles relay):", type="password")

            if st.button("📨 Send Email"):
                try:
                    msg = MIMEMultipart()
                    msg["From"] = sender_email
                    msg["To"] = email
                    msg["Subject"] = f"Payment details — Code {payment_code}"
                    msg.attach(MIMEText(email_body, "plain"))

                    with smtplib.SMTP("smtp.office365.com", 587) as server:
                        server.starttls()
                        if sender_pass:
                            server.login(sender_email, sender_pass)
                        server.send_message(msg)

                    st.success(f"✅ Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"❌ Error sending email: {e}")
else:
    st.info("Please upload your Excel file to start.")
