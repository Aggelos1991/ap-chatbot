import pandas as pd
import streamlit as st
import win32com.client as win32

# ====================== STREAMLIT CONFIG ======================
st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

st.markdown("""
Upload your Excel file, enter a **Payment Code**, and the bot will:
1. Find all invoices linked to that payment.
2. Summarize invoice amounts and totals.
3. Retrieve the vendor email.
4. Generate and send a professional email automatically via Outlook.
""")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Read and clean Excel
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]        # remove extra spaces
        df = df.loc[:, ~df.columns.duplicated()]                 # remove duplicate columns
        st.success("✅ Excel file loaded successfully!")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {e}")
        st.stop()

    # ====================== COLUMN CHECK ======================
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

            # ====================== OUTLOOK EMAIL SENDER ======================
            if st.button("📨 Send Email via Outlook"):
                try:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = f"Payment details — Code {payment_code}"
                    mail.Body = email_body
                    mail.Send()
                    st.success(f"✅ Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"❌ Error sending email via Outlook: {e}")
else:
    st.info("Please upload your Excel file to start.")
