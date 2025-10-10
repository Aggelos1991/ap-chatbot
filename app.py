import pandas as pd
import streamlit as st
import platform
import re

st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

uploaded_file = st.file_uploader("Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Excel file loaded successfully!")

    st.write("### 🧭 Columns found in your Excel:")
    st.dataframe(pd.DataFrame({"Columns": df.columns}))

    # Helper: auto-detect column names
    def find_col(patterns):
        for col in df.columns:
            if any(re.search(p, col, re.IGNORECASE) for p in patterns):
                return col
        return None

    # Detect columns
    col_payment = find_col(["payment", "pay doc", "payment document", "code"])
    col_alt_doc = find_col(["alternative", "alt doc", "document"])
    col_amount = find_col(["amount", "importe", "value", "total"])
    col_vendor = find_col(["vendor", "supplier", "proveedor"])
    col_email = find_col(["email", "mail", "correo"])

    if not all([col_payment, col_alt_doc, col_amount, col_vendor, col_email]):
        st.error("❌ Some required columns could not be identified. Please verify your Excel headers.")
    else:
        payment_code = st.text_input("Enter payment code:")

        if payment_code:
            subset = df[df[col_payment].astype(str).str.strip() == str(payment_code).strip()]

            if subset.empty:
                st.warning("No records found for this payment code.")
            else:
                vendor = subset[col_vendor].iloc[0]
                email = subset[col_email].iloc[0]

                # Build summary pivot table by Payment Document + Alternative Document
                summary = subset[[col_payment, col_alt_doc, col_amount]]
                total = summary[col_amount].sum()

                st.write("### 🧾 Invoices (Alternative Documents) for this Payment Code:")
                st.dataframe(summary)
                st.write(f"**Vendor:** {vendor}")
                st.write(f"**Email:** {email}")
                st.write(f"**Total Payment Amount:** €{total:,.2f}")

                # Email table formatting
                table_header = "| Payment Doc | Alternative Document | Amount (€) |\n|--------------|---------------------|-------------|\n"
                invoice_rows = "\n".join(
                    f"| {row[col_payment]} | {row[col_alt_doc]} | {row[col_amount]:,.2f} |"
                    for _, row in summary.iterrows()
                )

                email_body = f"""
Dear {vendor},

Please find below the invoices corresponding to the payment we made under payment code {payment_code}.

{table_header}{invoice_rows}

**Total amount:** €{total:,.2f}

Thank you for your cooperation.

Kind regards,
Angelos Keramaris
Accounts Payable Department
Ikos Resorts
"""

                st.text_area("📧 Email draft:", email_body, height=350)

                if st.button("📨 Send Email"):
                    os_name = platform.system()
                    if os_name == "Windows":
                        try:
                            import win32com.client as win32
                            outlook = win32.Dispatch("outlook.application")
                            mail = outlook.CreateItem(0)
                            mail.To = email
                            mail.Subject = f"Payment details — Code {payment_code}"
                            mail.Body = email_body
                            mail.Send()
                            st.success(f"Email successfully sent to {email}")
                        except Exception as e:
                            st.error(f"Error sending email via Outlook: {e}")
                    else:
                        st.info("Outlook automation not available on this system.")
                        st.code(f"To: {email}\nSubject: Payment details — Code {payment_code}\n\n{email_body}")
