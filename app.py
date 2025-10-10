import pandas as pd
import streamlit as st
import win32com.client as win32

st.set_page_config(page_title="💼 Vendor Payment Reconciliation & Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot")

# ==============================
# STEP 1 — Load Excel
# ==============================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    # Load and clean columns
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]       # Clean whitespace
    df = df.loc[:, ~df.columns.duplicated()]                # Remove duplicates

    st.success("✅ Excel file loaded successfully!")
    st.write("### 🧭 Columns detected in your Excel:")
    st.dataframe(pd.DataFrame(df.columns, columns=["Columns"]))

    # ==============================
    # STEP 2 — Validate expected columns
    # ==============================
    required_cols = ["Payment Code", "Invoice No", "Amount", "Vendor", "Supplier's Email"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"❌ Missing required columns in Excel: {missing_cols}")
        st.stop()

    # ==============================
    # STEP 3 — Ask for Payment Code
    # ==============================
    payment_code = st.text_input("🔎 Enter Payment Code:")

    if payment_code:
        subset = df[df["Payment Code"].astype(str).str.strip() == str(payment_code).strip()]

        if subset.empty:
            st.warning("⚠️ No records found for this Payment Code.")
        else:
            # ==============================
            # STEP 4 — Generate Summary
            # ==============================
            summary = subset.groupby("Invoice No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Supplier's Email"].iloc[0]

            st.write("### 🧾 Invoices related to this Payment Code:")
            st.dataframe(summary)
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Amount:** €{total:,.2f}")

            # ==============================
            # STEP 5 — Generate Email
            # ==============================
            invoice_lines = "\n".join(f"- {row['Invoice No']}: €{row['Amount']:,.2f}" for _, row in summary.iterrows())
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

            # ==============================
            # STEP 6 — Test Outlook Connection
            # ==============================
            if st.button("🧠 Test Outlook Connection"):
                try:
                    outlook = win32.Dispatch("Outlook.Application")
                    _ = outlook.GetNamespace("MAPI")
                    st.success("✅ Outlook is connected successfully!")
                except Exception as e:
                    st.error(f"❌ Outlook connection failed: {e}")

            # ==============================
            # STEP 7 — Send Email via Outlook
            # ==============================
            if st.button("📨 Send Email via Outlook"):
                try:
                    outlook = win32.Dispatch("Outlook.Application")
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = f"Payment details — Code {payment_code}"
                    mail.Body = email_body
                    mail.Send()
                    st.success(f"✅ Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"❌ Failed to send email. Error: {e}")
