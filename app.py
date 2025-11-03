# ==========================================================
# The Remitator — GLPI Integration (Final+ Assign + Fix Solution)
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

def glpi_update_ticket(token, ticket_id, status=None, category_id=None):
    payload = {"input": {}}
    if status is not None: payload["input"]["status"] = status
    if category_id is not None: payload["input"]["itilcategories_id"] = category_id
    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_add_solution(token, ticket_id, html, solution_type_id=10):
    body = {
        "input": {
            "tickets_id": int(ticket_id),
            "content": html,
            "solutiontypes_id": int(solution_type_id)
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILSolution",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
    )

def glpi_assign_user(token, ticket_id, assigned_email):
    body = {
        "input": {
            "tickets_id": int(ticket_id),
            "use_notification": 1,
            "type": 2,  # Assigned To
            "alternative_email": assigned_email
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/Ticket_User",
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
    subset = df[df["Payment Document Code"].astype(str) == st.text_input("🔎 Payment Code:")].copy()
    if subset.empty: st.stop()

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)
    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]
    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    # CN Logic
    cn_rows, debug_rows, unmatched_invoices = [], [], []
    if cn_file:
        cn = pd.read_excel(cn_file)
        cn.columns = [c.strip() for c in cn.columns]
        cn = cn.loc[:, ~cn.columns.duplicated()]
        cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
        cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "DEBE", "Invoice Value", "Invoice Value (€)"])
        if cn_alt_col and cn_val_col:
            cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
            used = set()
            for _, row in subset.iterrows():
                inv = str(row["Alt. Document"])
                diff = round(row["Payment Value"] - row["Invoice Value"], 2)
                match = False
                for i, r in cn.iterrows():
                    if i in used: continue
                    if round(abs(r[cn_val_col]), 2) == round(abs(diff), 2):
                        cn_rows.append({"Alt. Document": f"{r[cn_alt_col]} (CN)", "Invoice Value": -abs(r[cn_val_col])})
                        used.add(i); match=True; break
                if not match and abs(diff) > 0.01:
                    unmatched_invoices.append({"Alt. Document": f"{inv} (Unmatched Diff)", "Invoice Value": diff})
                debug_rows.append({
                    "Invoice": inv, "Invoice Value": row["Invoice Value"],
                    "Payment Value": row["Payment Value"], "Difference": diff,
                    "Matched?": "✅" if match else "❌"
                })

    all_rows = pd.concat([summary, pd.DataFrame(cn_rows), pd.DataFrame(unmatched_invoices)], ignore_index=True)
    total_val = subset["Payment Value"].sum()
    all_rows.loc[len(all_rows)] = ["TOTAL", total_val]
    all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = all_rows[["Alt. Document", "Invoice Value (€)"]]

    tab1, tab2 = st.tabs(["📋 Summary", "🔗 GLPI"])
    with tab1:
        st.dataframe(display_df, use_container_width=True)
        if debug_rows:
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

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

            with st.spinner("Sending to GLPI..."):
                glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                glpi_assign_user(token, ticket_id, assigned_email)
                sol = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)

            if str(sol.status_code).startswith("2"):
                st.success(f"✅ Ticket #{ticket_id} solved, assigned to {assigned_email}, and solution added.")
            else:
                st.error(f"❌ GLPI error: {sol.status_code} → {sol.text}")
else:
    st.info("📂 Upload Payment Excel to begin.")
