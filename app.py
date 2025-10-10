import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ====================== CONFIG ======================
st.set_page_config(page_title="💬 AP Payment Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Sender (Outlook 365 Compatible)")

st.markdown("""
This app:
1. Reads your Excel file with payment data.
2. Lets you search by **Payment Code**.
3. Shows invoice summary.
4. Sends an email securely via Outlook 365 (no MSAL, no win32).
""")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    st.success("✅ Excel loaded!")
    st.write("Detected columns:", list(df.columns))

    required_cols = ["Payment_Code", "Invoice_No", "Amount", "Vendor", "Email"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"⚠️ Missing columns: {missing}")
        st.stop()

    # ====================== INPUTS ======================
    payment_code = st.text_input("🔎 Enter Payment Code:")

    if payment_code:
        subset = df[df["Payment_Code"].astype(str) == str(payment_code)]
        if subset.empty:
            st.warning("⚠️ No records found for that code.")
        else:
            summary = subset.groupby("Invoice_No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Email"].iloc[0]

            st.subheader(f"📋 Summary for Payment Code: {payment_code}")
            st.dataframe(summary.style.format({"Amount": "€{:,.2f}".format}))
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Vendor Email:** {email}")
            st.write(f"**Total:** €{total:,.2f}")

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
            st.text_area("Email Body", email_body.strip(), height=250)

            # ====================== EMAIL SETTINGS ======================
            st.divider()
            st.subheader("📧 Email Sending Settings")

            sender_email = st.text_input("Your Outlook email address (e.g. aggelos@ikosresorts.com)")
            sender_password = st.text_input("Your Outlook App Password (if MFA disabled) or leave blank for manual auth", type="password")

            if st.button("📨 Send Email"):
                try:
                    msg = MIMEMultipart()
                    msg["From"] = sender_email
                    msg["To"] = email
                    msg["Subject"] = f"Payment details — Code {payment_code}"
                    msg.attach(MIMEText(email_body, "plain"))

                    # Send via Outlook SMTP relay
                    with smtplib.SMTP("smtp.office365.com", 587) as server:
                        server.starttls()
                        if sender_password:
                            server.login(sender_email, sender_password)
                        else:
                            st.info("If authentication fails, please use an app password or contact IT for SMTP relay access.")
                        server.send_message(msg)

                    st.success(f"✅ Email sent successfully to {email}")
                except Exception as e:
                    st.error(f"❌ Email sending failed: {e}")

else:
    st.info("Please upload your Excel file to begin.")
