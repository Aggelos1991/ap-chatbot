# ==========================================================
# Remitator — GLPI Integration (English Version • Final FIXED)
# ==========================================================
import os, re, requests
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------
st.set_page_config(page_title="Remitator — GLPI", layout="wide")
st.markdown(
    """
    <style>
        div.stButton > button:first-child {
            background-color: #007BFF;
            color: white;
            border-radius: 6px;
            height: 2.5em;
            width: 160px;
            font-size: 15px;
            border: none;
        }
        div.stButton > button:first-child:hover {
            background-color: #0069d9;
            color: white;
        }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("Remitator — Automatic Remittance Upload to GLPI")

# ----------------------------------------------------------
# ENVIRONMENT
# ----------------------------------------------------------
load_dotenv()
GLPI_URL   = os.getenv("GLPI_URL")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
    st.error("⚠️ Missing variables in .env file (GLPI_URL, APP_TOKEN, USER_TOKEN)")
    st.stop()

# ----------------------------------------------------------
# GLPI FUNCTIONS
# ----------------------------------------------------------
def login():
    r = requests.get(
        f"{GLPI_URL}/initSession",
        headers={"Authorization": f"user_token {USER_TOKEN}", "App-Token": APP_TOKEN}
    )
    return r.json().get("session_token")

def add_solution(token, ticket_id, html):
    requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILSolution",
        json={"input": {"tickets_id": ticket_id, "content": html, "solutiontypes_id": 1}},
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

def update_ticket(token, ticket_id, payload):
    requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={
            "Session-Token": token,
            "App-Token": APP_TOKEN,
            "Content-Type": "application/json"
        }
    )

# ----------------------------------------------------------
# HELPERS
# ----------------------------------------------------------
def parse_amount(value):
    if pd.isna(value): 
        return 0.0
    s = str(value).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0


def find_col(df, options):
    """Return the first matching column name from a list of possible names."""
    for c in df.columns:
        if c.strip().lower() in [o.lower() for o in options]:
            return c
    return None


# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
col1, col2, col3 = st.columns(3)
ticket_id = col1.text_input("Ticket ID", placeholder="101004")
vendor_email = col2.text_input("Vendor Email", placeholder="vendor@email.com")
payment_code = col3.text_input("Payment Code", placeholder="F2401223")

pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📄 Upload Credit Notes Excel (optional)", type=["xlsx"])

# ----------------------------------------------------------
# MAIN ACTION
# ----------------------------------------------------------
if st.button("Close Ticket"):
    if not ticket_id or not payment_code or not pay_file:
        st.error("Please fill in Ticket ID, Payment Code, and upload the Payment Excel file.")
        st.stop()

    # --- Load Payment Excel
    df_pay = pd.read_excel(pay_file)
    df_pay.columns = [c.strip() for c in df_pay.columns]

    pay_code_col = find_col(df_pay, ["Payment Document Code", "Payment Code", "Code"])
    doc_col = find_col(df_pay, ["Alt. Document", "Alternative Document", "Document", "Invoice Number"])
    value_col = find_col(df_pay, ["Invoice Value (€)", "Invoice Value", "Amount", "Charge", "Importe (€)"])

    if not all([pay_code_col, doc_col, value_col]):
        st.error("❌ Missing one of the required columns: Payment Code, Document, or Amount.")
        st.stop()

    df_pay[pay_code_col] = df_pay[pay_code_col].astype(str)
    subset = df_pay[df_pay[pay_code_col] == payment_code].copy()

    if subset.empty:
        st.error(f"No invoices found for payment code {payment_code}.")
        st.stop()

    # Confirmation message
    st.success(f"✅ Payment {payment_code} found — {len(subset)} invoice(s) loaded successfully.")

    all_rows = subset[[doc_col, value_col]].copy()
    all_rows.columns = ["Factura / Documento", "Importe (€)"]

    # --- Merge Credit Notes if uploaded
    if cn_file:
        df_cn = pd.read_excel(cn_file)
        df_cn.columns = [c.strip() for c in df_cn.columns]
        cn_doc = find_col(df_cn, ["Alt. Document", "Alternative Document", "Document"])
        cn_val = find_col(df_cn, ["Invoice Value (€)", "Invoice Value", "Amount", "Charge"])
        if cn_doc and cn_val:
            df_cn = df_cn[[cn_doc, cn_val]]
            df_cn.columns = ["Factura / Documento", "Importe (€)"]
            all_rows = pd.concat([all_rows, df_cn], ignore_index=True)

    # --- Compute total
    all_rows["Importe (€)"] = all_rows["Importe (€)"].apply(parse_amount)
    total = all_rows["Importe (€)"].sum()
    all_rows.loc[len(all_rows)] = ["TOTAL", total]

    # --- HTML Table + Email
    html_table = all_rows.to_html(index=False, border=0, justify="center", classes="table")
    html_message = f"""
    <p><strong>Estimado proveedor,</strong></p>
    <p>Adjuntamos el detalle de las facturas incluidas en el pago realizado.<br>
    <strong>Código de pago:</strong> {payment_code}</p>
    {html_table}
    <p>Quedamos a su disposición para cualquier aclaración.</p>
    <p>Saludos cordiales,<br><strong>Equipo Finance</strong></p>
    """

    # --- GLPI API
    token = login()
    if not token:
        st.error("❌ Failed to log in to GLPI. Check credentials or tokens.")
        st.stop()

    with st.spinner("Updating ticket and sending message to vendor..."):
        add_solution(token, ticket_id, html_message)
        update_ticket(token, ticket_id, {"input": {"status": 5}})  # Solved
        update_ticket(token, ticket_id, {"input": {"status": 6}})  # Closed

    st.success(f"✅ Ticket #{ticket_id} successfully updated and vendor message sent.")
    st.markdown("---")
    st.markdown("**Preview of the message sent:**")
    st.markdown(html_message, unsafe_allow_html=True)
