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
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    
    st.success("✅ Excel file loaded successfully!")
    st.write("### 🧭 Columns detected in your Excel:")
    st.dataframe(pd.DataFrame(df.columns, columns=["Columns"]))

    # ==============================
    # STEP 2 — Validate Columns (based on your file)
    # ==============================
    required_cols = [
        "Payment Document Code",  # instead of Payment Code
        "Alt Document",           # instead of Invoice No
        "Invoice Value",          # instead of Amount
        "Supplier Name",          # instead of Vendor
        "Supplier's Email"        # keep same if exists
    ]

    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing required columns in Excel: {missing_cols}")
        st.stop()

    # ==============================
    # STEP 3 — Ask for Payment Code
    # ==============================
    payment_code = st.text_input("🔎 Enter Payment Document Code:")

    if payment_code:
        subset = df[df["Payment Document Code"].astype(str).str.strip() == str(payment_code).strip()]

        if subset.empty:
            st.warning("⚠️ No records found for this Payment Document Code.")
        else:
            # ==============================
            # STEP 4 — Generate Summary
            # ==============================
            summary = subset.groupby("Alt Document", as_index=False)["Invoice Value"].sum()
            total = summary["Invoice Value"].sum()
            vendor = subset["Supplier Name"].iloc[0]

            # Try to get email from either Supplier's Email or alternative Greek column
            possible_email_cols = [col for col in df.columns if "email" in col.lower()]
            email = subset[possible_email_cols[0]].iloc[0] if possible_email_cols else "N/A"

            st.write("### 🧾 Invoices related to this Payment Document Code:")
            st.dataframe(summary)
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Amount:** €{total:,.2f}")

            # ==============================
            # STEP 5 — Generate Email
            # ==============================
            invoice_lines = "\n".join(
                f"- {row['Alt Document']}: €{row['Invoice Value']:,.2f}" for _, row in summary.iterrows()
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
            # STEP 6 — Send Email via Outlook
            # ==============================
            if st.button("📨 Send Email via Outlook"):
                try:
                    outlook = win32.Dispatch("Outlook.Application")
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = f"Payment details — Document {payment_code}"
                    mail.Body = email_body
                    mail.Send()
                    st.success(f"✅ Email successfully sent to {email}")
                except Exception as e:
                    st.error(f"❌ Failed to send email. Error: {e}")
