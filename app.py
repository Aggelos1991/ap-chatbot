import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText

# ===== Streamlit config =====
st.set_page_config(page_title="💼 Vendor Payment Reconciliation & Email Bot", layout="wide")
st.title("💼 Vendor Payment Reconciliation & Email Bot (Outlook 365)")

uploaded_file = st.file_uploader("📂 Upload Excel (TEST.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # keep your column names, just trim and drop duplicate headers
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Excel loaded")
        st.write("Columns detected:", list(df.columns))
    except Exception as e:
        st.error(f"❌ Error loading Excel: {e}")
        st.stop()

    # --- REQUIRED columns (your exact names) ---
    REQ = [
        "Payment Document Code",
        "Alt. Document",
        "Invoice Value",
        "Supplier Name",
        "Supplier's Email",
    ]
    missing = [c for c in REQ if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")

    if pay_code:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)]
        if subset.empty:
            st.warning("⚠️ No rows found for this Payment Document Code.")
        else:
            # ensure numeric amounts
            subset = subset.copy()
            subset["Invoice Value"] = pd.to_numeric(subset["Invoice Value"], errors="coerce").fillna(0)

            # pivot-like summary by invoice
            summary = subset.groupby("Alt. Document", as_index=False)["Invoice Value"].sum()
            total = float(summary["Invoice Value"].sum())

            # take vendor/email from the subset (assumed same vendor per payment)
            vendor = str(subset["Supplier Name"].dropna().iloc[0])
            email_to = str(subset["Supplier's Email"].dropna().iloc[0])

            st.divider()
            st.subheader(f"📋 Summary for Payment Code: {pay_code}")
            st.write(f"**Vendor:** {vendor}")
            st.write(f"**Email:** {email_to}")
            st.dataframe(summary.style.format({"Invoice Value": "€{:,.2f}".format}))
            st.write(f"**Total Payment Amount:** €{total:,.2f}")

            # email body
            lines = "\n".join(
                f"- {row['Alt. Document']}: €{row['Invoice Value']:,.2f}"
                for _, row in summary.iterrows()
            )
            email_body = f"""Dear {vendor},

Please find below the invoices corresponding to the payment we made under payment code {pay_code}.

{lines}

Total amount: €{total:,.2f}

Thank you for your cooperation.

Kind regards,
Angelos Keramaris
Accounts Payable Department
Ikos Resorts
"""

            st.divider()
            st.subheader("✉️ Email Preview")
            st.text_area("Generated Email:", email_body, height=260)

            # ---- Outlook SMTP (uses your Microsoft 365 username & password) ----
            st.divider()
            st.subheader("📧 Outlook 365 Login")
            sender_email = st.text_input("Your Outlook email (From):", placeholder="you@yourdomain.com")
            sender_pass = st.text_input("Your Outlook password:", type="password")

            if st.button("📨 Send Email"):
                try:
                    if not sender_email or not sender_pass:
                        st.warning("Enter your Outlook email and password.")
                    else:
                        msg = MIMEText(email_body)
                        msg["Subject"] = f"Payment details — Code {pay_code}"
                        msg["From"] = sender_email
                        msg["To"] = email_to

                        with smtplib.SMTP("smtp.office365.com", 587) as server:
                            server.ehlo()
                            server.starttls()
                            server.ehlo()
                            server.login(sender_email, sender_pass)
                            server.send_message(msg)

                        st.success(f"✅ Email sent to {email_to}")
                except smtplib.SMTPAuthenticationError as e:
                    st.error("❌ Authentication failed. If your org blocks basic SMTP, create an Outlook **app password** and use it here.")
                except Exception as e:
                    st.error(f"❌ Error sending email: {e}")
else:
    st.info("Upload your Excel to begin.")
