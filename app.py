# ==========================================================
# The Remitator — GLPI Integration (FINAL • Auto Email→UserID • Solution + AP Extras)
# ==========================================================
import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from dotenv import load_dotenv

# ========== UI ==========
st.set_page_config(page_title="The Remitator", layout="wide")
st.markdown("""
<style>
  div.stButton > button:first-child{
    background-color:#007BFF;color:white;border:none;border-radius:6px;
    height:2.4em;width:160px;font-size:15px
  }
  div.stButton > button:first-child:hover{background-color:#0069d9}
</style>
""", unsafe_allow_html=True)
st.title("💀 The Remitator — Hasta la vista, payment remittance. 💀")

# ========== ENV ==========
load_dotenv()
GLPI_URL   = os.getenv("GLPI_URL")            # e.g. https://glpi.ikosgroup.com/apirest.php
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")
PLUGIN_FIELD_ENDPOINT = f"{GLPI_URL}/PluginFieldsSolutioncategoryfield/"
SOLUTION_CATEGORY_ID = 10   # Payment Remittance Advice

# ========== HELPERS ==========
def parse_amount(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif s.count(",") == 1: s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def find_col(df, names):
    for c in df.columns:
        name = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            if n.replace(" ", "").replace(".", "").lower() in name:
                return c
    return None

# ========== GLPI API ==========
def glpi_login():
    r = requests.get(f"{GLPI_URL}/initSession",
                     headers={"Authorization": f"user_token {USER_TOKEN}", "App-Token": APP_TOKEN})
    return r.json().get("session_token")

def glpi_search_user(token, email):
    r = requests.get(
        f"{GLPI_URL}/User",
        params={"criteria[0][field]": 5, "criteria[0][searchtype]": "contains", "criteria[0][value]": email},
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )
    data = r.json()
    if isinstance(data, dict) and "data" in data and data["data"]:
        return data["data"][0].get("id")
    return None

def glpi_update_ticket(token, ticket_id, status=5, category_id=400):
    payload = {"input": {"status": status, "itilcategories_id": category_id}}
    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_assign_user(token, ticket_id, user_id):
    body = {"input": {"tickets_id": int(ticket_id), "users_id": int(user_id), "type": 2}}
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/Ticket_User/",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_add_solution(token, ticket_id, html, solution_type_id=10):
    body = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "solutiontypes_id": int(solution_type_id),
            "content": html
        }
    }
    r = requests.post(
        f"{GLPI_URL}/ITILSolution/",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )
    if r.status_code == 201 and "id" in r.json():
        sol_id = r.json()["id"]
        # set AP Extras → Solution Category = Payment Remittance Advice (10)
        requests.put(
            f"{PLUGIN_FIELD_ENDPOINT}{sol_id}",
            json={"input": {"plugin_fields_solutioncategoryfields_id": SOLUTION_CATEGORY_ID}},
            headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
        )
    return r

# ========== MAIN ==========
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])

if pay_file:
    df = pd.read_excel(pay_file)
    df.columns = [c.strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    pay_code = st.text_input("🔎 Payment Code:")
    if not pay_code: st.stop()
    subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty: st.warning("⚠️ No rows found for this Payment Code."); st.stop()

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)
    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]
    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    total_val = subset["Payment Value"].sum()
    summary.loc[len(summary)] = ["TOTAL", total_val]
    summary["Invoice Value (€)"] = summary["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = summary[["Alt. Document", "Invoice Value (€)"]]

    tab1, tab2 = st.tabs(["📋 Summary", "🔗 GLPI"])
    with tab1:
        st.dataframe(display_df, use_container_width=True)

    with tab2:
        c1, c2, c3 = st.columns(3)
        ticket_id = c1.text_input("Ticket ID", placeholder="101004")
        category_id = c2.text_input("Category ID", value="400")
        assigned_email = c3.text_input("Assign To Email", placeholder="finance@ikosgroup.com")

        html_table = display_df.to_html(index=False, border=0, justify="center", classes="table")
        html_message = f"""
        <p><strong>Estimado proveedor,</strong></p>
        <p>Por favor, encuentre a continuación las facturas que corresponden al pago realizado.</p>
        {html_table}
        <p>Quedamos a su disposición para cualquier aclaración.</p>
        <p>Saludos cordiales,<br><strong>Equipo Finance</strong></p>
        """
        st.markdown(html_message, unsafe_allow_html=True)

        if st.button("Send to GLPI"):
            if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]): st.error("Missing GLPI credentials."); st.stop()
            token = glpi_login()
            if not token: st.error("Failed GLPI session."); st.stop()

            with st.spinner("Processing..."):
                # lookup user ID by email
                user_id = glpi_search_user(token, assigned_email)
                if not user_id:
                    st.error(f"No user found for {assigned_email}")
                    st.stop()

                glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                glpi_assign_user(token, ticket_id, user_id)
                sol = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)

            if str(sol.status_code).startswith("2") or sol.status_code == 201:
                st.success(f"✅ Ticket #{ticket_id} solved, assigned to {assigned_email}, and solution added with AP Extras category 10.")
            else:
                st.error(f"❌ GLPI error: {sol.status_code} → {sol.text}")
else:
    st.info("📂 Upload Payment Excel to begin.")
