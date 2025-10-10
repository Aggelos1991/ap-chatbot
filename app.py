import pandas as pd
import streamlit as st
import requests
import msal
import json

# ====================== STREAMLIT CONFIG ======================
st.set_page_config(page_title="💬 Vendor Payment Chatbot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot (Microsoft 365 Secure Login)")

st.markdown("""
Upload your Excel file, enter a **Payment Code**, and the bot will:
1. Find all invoices linked to that payment.
2. Summarize invoice amounts and totals.
3. Retrieve the vendor email.
4. Generate and send a professional email via **Microsoft Graph API** (no SMTP or win32com required).
""")

# ====================== MICROSOFT GRAPH CONFIG ======================
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "common"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send"]

# Function to authenticate user and get access token
def get_access_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        result = None

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        st.write("To authorize, open the following link and enter the code shown below:")
        st.code(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        return result["access_token"]
    else:
        st.error("Authentication failed.")
        return None

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Excel file loaded successfully!")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel file: {e}")
        st.stop()

    required_cols = ["Payment_Code", "Invoice_No", "Amount", "Vendor", "Email"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns: {missing}. Please correct your Excel and re-upload.")
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

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {payment_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.dataframe(summary.style.format({"Amount": "€{:,.2f}".format}))
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

            st.divider()
            st.subheader("✉️ Email Preview")
            st.text_area("Generated Email:", email_body.strip(), height=250)

            # ====================== MICROSOFT GRAPH EMAIL SEND ======================
            if st.button("📨 Send Email via Microsoft 365"):
                token = get_access_token()
                if token:
                    graph_url = "https://graph.microsoft.com/v1.0/me/sendMail"
                    mail_data = {
                        "message": {
                            "subject": f"Payment details — Code {payment_code}",
                            "body": {"contentType": "Text", "content": email_body},
                            "toRecipients": [{"emailAddress": {"address": email}}],
                        },
                        "saveToSentItems": "true"
                    }
                    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
                    response = requests.post(graph_url, headers=headers, data=json.dumps(mail_data))

                    if response.status_code == 202:
                        st.success(f"✅ Email successfully sent to {email}")
                    else:
                        st.error(f"❌ Error sending email: {response.status_code} - {response.text}")
else:
    st.info("Please upload your Excel file to start.")
