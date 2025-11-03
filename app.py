# ==========================================================
# The Remitator — GLPI Integration (FINAL Production Version)
# ==========================================================
import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime
from itertools import combinations
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
GLPI_URL   = os.getenv("GLPI_URL")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

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

def glpi_add_solution(token, ticket_id, html, solution_type_id=10):
    body = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "solutiontypes_id": int(solution_type_id),
            "content": html
        }
    }
    return requests.post(
        f"{GLPI_URL}/ITILSolution/",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_update_ticket(token, ticket_id, status=5, category_id=400):
    payload = {"input": {"status": status, "itilcategories_id": category_id}}
    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_assign_user(token, ticket_id, user_id):
    body = {
        "input": {
            "tickets_id": int(ticket_id),
            "users_id": int(user_id),
            "type": 2  # Assigned to
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/Ticket_User/",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

# ========== APP ==========
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📂 (Optional) Upload Credit Notes Excel", type=["xlsx"])

if pay_file:
    df = pd.read_excel(pay_file)
    df.columns = [c.strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    pay_code = st.text_input("🔎 Payment Code:")
    if not pay_code: st.stop()

    subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("⚠️ No rows found for this Payment Document Code.")
        st.stop()

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)
    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]
    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    # CN Logic simplified
    cn_rows = []
    if cn_file:
        cn = pd.read_excel(cn_file)
        cn.columns = [c.strip() for c in cn.columns]
        cn = cn.loc[:, ~cn.columns.duplicated()]
        cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
        cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "DEBE", "Invoice Value", "Invoice Value (€)"])
        if cn_alt_col and cn_val_col:
            cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
            for _, row in subset.iterrows():
                diff = round(row["Payment Value"] - row["Invoice Value"], 2)
                for _, r in cn.iterrows():
                    if round(abs(r[cn_val_col]), 2) == round(abs(diff), 2):
                        cn_rows.append({"Alt. Document": f"{r[cn_alt_col]} (CN)", "Invoice Value": -abs(r[cn_val_col])})
                        break

    all_rows = pd.concat([summary, pd.DataFrame(cn_rows)], ignore_index=True)
    total_val = subset["Payment Value"].sum()
    all_rows.loc[len(all_rows)] = ["TOTAL", total_val]
    all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = all_rows[["Alt. Document", "Invoice Value (€)"]]

    tab1, tab2 = st.tabs(["📋 Summary", "🔗 GLPI"])
    with tab1:
        st.dataframe(display_df, use_container_width=True)

    with tab2:
        c1, c2, c3 = st.columns(3)
        ticket_id = c1.text_input("Ticket ID", placeholder="101004")
        category_id = c2.text_input("Category ID", value="400")
        user_id = c3.text_input("Assigned User ID", value="22487")

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
            if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
                st.error("Missing GLPI credentials in .env."); st.stop()
            token = glpi_login()
            if not token:
                st.error("Failed to start GLPI session."); st.stop()

            with st.spinner("Posting to GLPI..."):
                r1 = glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                r2 = glpi_assign_user(token, ticket_id, user_id)
                r3 = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)

            if all(str(r.status_code).startswith("2") for r in [r1, r2, r3]):
                st.success(f"✅ Ticket #{ticket_id} solved, category {category_id}, assigned to {user_id}, and solution posted (type 10).")
            else:
                st.error(f"❌ GLPI error: {[r.status_code for r in [r1, r2, r3]]}")
else:
    st.info("📂 Upload Payment Excel to begin.")
