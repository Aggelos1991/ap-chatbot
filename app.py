import pandas as pd
import streamlit as st
import requests
from msal import PublicClientApplication

# ====================== CONFIG ======================
st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot (Outlook 365 Secure)")

CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID"
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/Mail.Send"]

# ====================== LOGIN ======================
st.sidebar.subheader("🔐 Microsoft 365 Login")

if "access_token" not in st.session_state:
    st.session_state.access_token = None

if st.button("Login with Microsoft 365"):
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    result = app.acquire_token_interactive(scopes=SCOPES)
    st.session_state.access_token = result["access_token"]
    st.sidebar.success("✅ Logged in successfully!")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    required_cols = ["Payment_Code", "Invoice_No", "Amount", "Vendor", "Email"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    payment_code = st.text_input("🔎 Enter Payment Code:")

    if payment_code:
        subset = df[df["Payment_Code"].astype(str) == str(payment_code)]
        if subset.empty:
            st.warning("⚠️ No records found for this payment code.")
        else:
            summary = subset.groupby("Invoice_No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Email"].iloc[0]

            st.subheader(f"📋 Summary for Payment Code: {payment_code}")
            st.dataframe(summary.style.format({"Amount": "€{:,.2f}".format}))
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

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

            st.text_area("✉️ Email Preview", email_body.strip(), height=250)

            # ====================== SEND EMAIL USING GRAPH API ======================
            if st.button("📨 Send Email"):
                if not st.session_state.access_token:
                    st.error("Please log in with Microsoft 365 first.")
                else:
                    email_json = {
                        "message": {
                            "subject": f"Payment details — Code {payment_code}",
                            "body": {
                                "contentType": "Text",
                                "content": email_body
                            },
                            "toRecipients": [{"emailAddress": {"address": email}}],
                        },
                        "saveToSentItems": "true"
                    }

                    response = requests.post(
                        "https://graph.microsoft.com/v1.0/me/sendMail",
                        headers={"Authorization": f"Bearer {st.session_state.access_token}",
                                 "Content-Type": "application/json"},
                        json=email_json
                    )

                    if response.status_code == 202:
                        st.success(f"✅ Email sent successfully to {email}")
                    else:
                        st.error(f"❌ Failed to send email: {response.text}")
