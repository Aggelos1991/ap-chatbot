import pandas as pd
import streamlit as st
import win32com.client as win32

st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

# Step 1. Upload Excel
uploaded_file = st.file_uploader("Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Excel file loaded successfully!")

    # Step 2. Input payment code
    payment_code = st.text_input("Enter payment code:")

    if payment_code:
        subset = df[df["Payment_Code"].astype(str) == str(payment_code)]

        if subset.empty:
            st.warning("No records found for this payment code.")
        else:
            # Pivot-style summary
            summary = subset.groupby("Invoice_No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Email"].iloc[0]

            st.write("### 🧾 Invoices related to this payment code:")
            st.dataframe(summary)
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

            # Step 3. Generate email text
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

            st.text_area("📧 Email draft:", email_body, height=250)

            # Step 4. Send via Outlook
            if st.button("📨 Send Email via Outlook"):
                try:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = f"Payment details — Code {payment_code}"
                    mail.Body = email_body
                    mail.Send()
                    st.success(f"Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"Error sending email: {e}")
