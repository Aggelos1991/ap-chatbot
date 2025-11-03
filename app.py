# ==========================================================
# Remitator — GLPI Integration (Final)
# ==========================================================
import os, json, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv

# ----------------------------------------------------------
# CONFIG
# ----------------------------------------------------------
st.set_page_config(page_title="Remitator — GLPI", layout="wide")
st.title("Remitator — Envío automático de remesas a GLPI")

load_dotenv()
GLPI_URL   = os.getenv("GLPI_URL")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
    st.error("⚠️ Faltan variables en tu archivo .env (GLPI_URL, APP_TOKEN, USER_TOKEN)")
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

def put(token, ticket_id, payload):
    requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

# ----------------------------------------------------------
# HELPER FUNCTIONS
# ----------------------------------------------------------
def parse_amount(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
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

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
col1, col2, col3 = st.columns(3)
ticket_id = col1.text_input("Ticket ID", placeholder="101004")
vendor_email = col2.text_input("Email del proveedor", placeholder="proveedor@empresa.com")
payment_code = col3.text_input("Código de pago", placeholder="F2401223")

pay_file = st.file_uploader("📂 Excel de pagos", type=["xlsx"])
cn_file  = st.file_uploader("📄 Excel de notas de crédito (opcional)", type=["xlsx"])

# ----------------------------------------------------------
# MAIN ACTION
# ----------------------------------------------------------
if st.button("🚀 Enviar detalle al ticket GLPI", type="primary", use_container_width=True):
    if not ticket_id or not payment_code or not pay_file:
        st.error("Completa Ticket ID, código de pago y sube el Excel de pagos.")
        st.stop()

    # --- Load Excel(s)
    df_pay = pd.read_excel(pay_file)
    df_pay.columns = [c.strip() for c in df_pay.columns]
    df_pay["Payment Document Code"] = df_pay["Payment Document Code"].astype(str)
    subset = df_pay[df_pay["Payment Document Code"] == payment_code].copy()

    if subset.empty:
        st.error(f"No se encontraron facturas con el código {payment_code}.")
        st.stop()

    all_rows = subset[["Alt. Document", "Invoice Value (€)"]].copy()
    all_rows.columns = ["Factura / Documento", "Importe (€)"]

    # If CN Excel uploaded, merge it for visual clarity
    if cn_file:
        df_cn = pd.read_excel(cn_file)
        df_cn.columns = [c.strip() for c in df_cn.columns]
        if "Alt. Document" in df_cn.columns and "Invoice Value (€)" in df_cn.columns:
            df_cn = df_cn[["Alt. Document", "Invoice Value (€)"]]
            df_cn.columns = ["Factura / Documento", "Importe (€)"]
            all_rows = pd.concat([all_rows, df_cn], ignore_index=True)

    # Compute total
    all_rows["Importe (€)"] = all_rows["Importe (€)"].apply(parse_amount)
    total = all_rows["Importe (€)"].sum()
    all_rows.loc[len(all_rows)] = ["TOTAL", total]

    # Build HTML table
    html_table = all_rows.to_html(index=False, border=0, justify="center", classes="table")

    # Build message
    html_message = f"""
    <p><strong>Estimado proveedor,</strong></p>
    <p>Adjuntamos el detalle de las facturas incluidas en el pago realizado.<br>
    <strong>Código de pago:</strong> {payment_code}</p>
    {html_table}
    <p>Quedamos a su disposición para cualquier aclaración.</p>
    <p>Saludos cordiales,<br><strong>Equipo Finance</strong></p>
    """

    # GLPI API calls
    token = login()
    if not token:
        st.error("❌ Error al iniciar sesión en GLPI.")
        st.stop()

    with st.spinner("Actualizando ticket y enviando mensaje al proveedor..."):
        # Add message
        add_solution(token, ticket_id, html_message)
        # Close ticket
        put(token, ticket_id, {"input": {"status": 5}})  # Solved
        put(token, ticket_id, {"input": {"status": 6}})  # Closed

    st.success(f"✅ Ticket #{ticket_id} actualizado con éxito y mensaje enviado.")
    st.markdown("---")
    st.markdown("**Vista previa del mensaje enviado:**")
    st.markdown(html_message, unsafe_allow_html=True)
