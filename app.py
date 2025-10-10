import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText

# ====================== STREAMLIT CONFIG ======================
st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot (Outlook 365 Login)")

st.markdown("""
Upload your Excel file, enter a **Payment Code**, and the bot will:
1. Find all invoices linked to that payment.
2. Summarize invoice amounts and totals.
3. Retrieve the vendor email.
4. Send the message automatically via Outlook 365 (SMTP login).
""")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Read and clean Excel
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Excel file loaded successfully!")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {e}")
        st.stop()

    # ====================== REQUIRED COLUMNS ======================
    required_cols = ["Payment_Code", "Invoice_No", "Amount", "Vendor", "Email"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns: {missing}. Please correct your Excel and re-upload.")
        st.stop()

    # ====================== PAYMENT CODE INPUT ======================
    payment_code = st.text_input("🔎 Enter Payment Code:")

    if payment_code:
        subset = df[df["Payment_Code"].astype(str) == str(payment_code)]

        if subset.empty:
            st.warning("⚠️ No records found for this payment code.")
        else:
            # ====================== DATA SUMMARY ======================
            summary = subset.groupby("Invoice_No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Email"].iloc[0]

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {payment_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.dataframe(summary.style.format({"Amount": "€{:,.2f}".format}))
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

            # ====================== EMAIL GENERATION ======================
            invoice_lines = "\n".join(f"- {row.Invoice_No}: €{row.Amount:,.2f}" for _, row in summary.iterrows())

            email_body = f"""
Dear {vendor},

Please find below the invoices corresponding to the payment we made under payment code {payment_code}.

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

            # ====================== OUTLOOK SMTP LOGIN ======================
            st.divider()
            st.subheader("📧 Outlook 365 Login")
            sender_email = st.text_input("Enter your Outlook email:", placeholder="you@saniikos.com")
            sender_pass = st.text_input("Enter your Outlook password:", type="password")

            if st.button("📨 Send Email"):
                try:
                    if not sender_email or not sender_pass:
                        st.warning("Please enter both your Outlook email and password.")
                    else:
                        msg = MIMEText(email_body)
                        msg["Subject"] = f"Payment details — Code {payment_code}"
                        msg["From"] = sender_email
                        msg["To"] = email

                        with smtplib.SMTP("smtp.office365.com", 587) as server:
                            server.starttls()
                            server.login(sender_email, sender_pass)
                            server.send_message(msg)

                        st.success(f"✅ Email successfully sent to {email}")
                except smtplib.SMTPAuthenticationError:
                    st.error("❌ Authentication failed. Please check your email or password, or enable SMTP access in your Outlook settings.")
                except Exception as e:
                    st.error(f"❌ Error sending email: {e}")
else:
    st.info("Please upload your Excel file to start.")
