import pandas as pd
import streamlit as st
import requests
from msal import PublicClientApplication

# ====================== APP CONFIG ======================
st.set_page_config(page_title="💬 AP Payment Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Outlook 365 Email Sender")

st.markdown("""
This app lets you:
1. Upload your Excel with invoices & payment codes.
2. Enter a **Payment Code** to see all related invoices.
3. Review the email preview.
4. Send the email directly via **Outlook 365 (Microsoft Graph API)**.
""")

# ====================== MICROSOFT AUTH CONFIG ======================
CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID"  # Replace with your registered Azure app ID
TENANT_ID = "common"  # Use 'common' if you want both work/school & personal Microsoft accounts
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/Mail.Send"]

# ====================== LOGIN SECTION ======================
st.sidebar.header("🔐 Microsoft 365 Login")

if "access_token" not in st.session_state:
    st.session_state.access_token = None

if st.sidebar.button("Login with Microsoft 365"):
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    result = app.acquire_token_interactive(scopes=SCOPES)
    st.session_state.access_token = result.get("access_token")
    if st.session_state.access_token:
        st.sidebar.success("✅ Successfully logged in with Microsoft 365")
    else:
        st.sidebar.error("❌ Login failed. Please try again.")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("📂 Upload your Excel file (e.g. TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]   # Clean headers
        df = df.loc[:, ~df.columns.duplicated()]            # Remove duplicates
        st.success("✅ Excel file loaded successfully!")
        st.write("Detected columns:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
        st.stop()

    # ====================== CHECK REQUIRED COLUMNS ======================
    required_cols = ["Payment_Code", "Invoice_No", "Amount", "Vendor", "Email"]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        st.error(f"⚠️ Missing columns in Excel: {missing_cols}")
        st.stop()

    # ====================== USER INPUT ======================
    payment_code = st.text_input("🔎 Enter Payment Code:")

    if payment_code:
        subset = df[df["Payment_Code"].astype(str) == str(payment_code)]

        if subset.empty:
            st.warning("⚠️ No invoices found for that payment code.")
        else:
            # ====================== PIVOT SUMMARY ======================
            summary = subset.groupby("Invoice_No", as_index=False)["Amount"].sum()
            total = summary["Amount"].sum()
            vendor = subset["Vendor"].iloc[0]
            email = subset["Email"].iloc[0]

            st.divider()
            st.subheader(f"📊 Summary for Payment Code: {payment_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email}")
            st.dataframe(summary.style.format({"Amount": "€{:,.2f}".format}))
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

            # ====================== EMAIL PREPARATION ======================
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

            # ====================== SEND EMAIL BUTTON ======================
            if st.button("📨 Send Email via Outlook 365"):
                if not st.session_state.access_token:
                    st.error("❌ Please log in with Microsoft 365 first.")
                else:
                    send_url = "https://graph.microsoft.com/v1.0/me/sendMail"
                    headers = {
                        "Authorization": f"Bearer {st.session_state.access_token}",
                        "Content-Type": "application/json"
                    }
                    email_payload = {
                        "message": {
                            "subject": f"Payment details — Code {payment_code}",
                            "body": {"contentType": "Text", "content": email_body},
                            "toRecipients": [{"emailAddress": {"address": email}}],
                        },
                        "saveToSentItems": "true"
                    }

                    try:
                        response = requests.post(send_url, headers=headers, json=email_payload)
                        if response.status_code == 202:
                            st.success(f"✅ Email successfully sent to {email}")
                        else:
                            st.error(f"❌ Email sending failed — {response.text}")
                    except Exception as e:
                        st.error(f"❌ Error sending email: {e}")
else:
    st.info("📁 Please upload your Excel file to begin.")
