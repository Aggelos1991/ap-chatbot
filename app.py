# ==========================================================
# The Remitator — GLPI Solution Sync (Final Clean)
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

# ---------- Streamlit base UI ----------
st.set_page_config(page_title="The Remitator", layout="wide")
st.markdown(
    """
    <style>
      div.stButton > button:first-child{
        background-color:#007BFF;color:white;border:none;border-radius:6px;
        height:2.4em;width:160px;font-size:15px
      }
      div.stButton > button:first-child:hover{background-color:#0069d9}
    </style>
    """, unsafe_allow_html=True
)
st.title("💀 The Remitator — Hasta la vista, payment remittance. 💀")

# ---------- GLPI ENV ----------
load_dotenv()
GLPI_URL   = os.getenv("GLPI_URL")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

# ---------- Helpers ----------
def parse_amount(v):
    if pd.isna(v): return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    try: return float(s)
    except: return 0.0

def find_col(df, names):
    for c in df.columns:
        name = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            if n.replace(" ", "").replace(".", "").lower() in name:
                return c
    return None

# ---------- GLPI API ----------
def glpi_login():
    if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]): return None
    r = requests.get(
        f"{GLPI_URL}/initSession",
        headers={"Authorization": f"user_token {USER_TOKEN}", "App-Token": APP_TOKEN},
        timeout=30
    )
    try:
        return r.json().get("session_token")
    except:
        return None

def glpi_update_ticket(token, ticket_id, status=None, category_id=None):
    payload = {"input": {}}
    if status is not None: payload["input"]["status"] = status
    if category_id is not None: payload["input"]["itilcategories_id"] = category_id
    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"},
        timeout=30
    )

def glpi_add_solution(token, ticket_id, html, solution_type_id=10):
    body = {"input": {"tickets_id": int(ticket_id), "content": html, "solutiontypes_id": int(solution_type_id)}}
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILSolution",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN},
        timeout=30
    )

# ---------- Uploads ----------
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📂 (Optional) Upload Credit Notes Excel", type=["xlsx"])

# ---------- Main Logic ----------
if pay_file:
    try:
        df = pd.read_excel(pay_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment file loaded successfully")
    except Exception as e:
        st.error(f"❌ Error loading Payment Excel: {e}")
        st.stop()

    req = ["Payment Document Code","Alt. Document","Invoice Value","Payment Value","Supplier Name","Supplier's Email"]
    missing = [c for c in req if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in Payment Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")
    if not pay_code: st.stop()

    subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("⚠️ No rows found for this Payment Document Code.")
        st.stop()

    cn = None
    if cn_file:
        try:
            cn = pd.read_excel(cn_file)
            cn.columns = [c.strip() for c in cn.columns]
            cn = cn.loc[:, ~cn.columns.duplicated()]
            st.info("📄 Credit Note file loaded and will be applied.")
        except Exception as e:
            st.warning(f"⚠️ Error loading CN file (will skip CN logic): {e}")
            cn = None
    else:
        st.info("ℹ️ No Credit Note file uploaded — showing payments only.")

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]

    summary = subset[["Alt. Document", "Invoice Value"]].copy()
    cn_rows, unmatched_invoices = [], []

    # ----- CN match simplified -----
    if cn is not None:
        cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
        cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "DEBE"])
        if cn_alt_col and cn_val_col:
            cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
            cn = cn[cn[cn_val_col].abs() > 0.01].reset_index(drop=True)

    # ----- Combine summary -----
    all_rows = summary.copy()
    if cn_rows: all_rows = pd.concat([all_rows, pd.DataFrame(cn_rows)], ignore_index=True)
    if unmatched_invoices: all_rows = pd.concat([all_rows, pd.DataFrame(unmatched_invoices)], ignore_index=True)

    total_val = all_rows["Invoice Value"].sum()
    all_rows = pd.concat([all_rows, pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_val}])], ignore_index=True)
    all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = all_rows[["Alt. Document", "Invoice Value (€)"]]

    # ----- Tabs -----
    tab1, tab2 = st.tabs(["📋 Summary", "🔗 GLPI"])
    with tab1:
        st.subheader(f"Final Summary for Payment Code: {pay_code}")
        st.write(f"**Vendor:** {vendor}")
        st.write(f"**Vendor Email (from file):** {vendor_email_in_file}")
        st.dataframe(display_df, use_container_width=True)

    with tab2:
        st.write("This will send the email (Spanish) with the table to GLPI as a **Solution** and mark the ticket **Solved** (Category 400, Solution Type 10).")
        c1, c2 = st.columns(2)
        ticket_id = c1.text_input("Ticket ID", placeholder="101004")
        category_id = c2.text_input("Category ID", value="400")

        # Spanish message
        html_table = display_df.to_html(index=False, border=0, justify="center", classes="table")
        html_message = f"""
        <p><strong>Estimado proveedor,</strong></p>
        <p>Por favor, encuentre a continuación las facturas que corresponden al pago realizado.</p>
        {html_table}
        <p>Quedamos a su disposición para cualquier aclaración.</p>
        <p>Saludos cordiales,<br><strong>Equipo Finance</strong></p>
        """

        st.markdown("**Preview of email to send:**")
        st.markdown(html_message, unsafe_allow_html=True)

        if st.button("Send to GLPI"):
            if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
                st.error("Missing GLPI credentials in .env (GLPI_URL, APP_TOKEN, USER_TOKEN).")
                st.stop()
            if not ticket_id.strip():
                st.error("Ticket ID is required.")
                st.stop()

            token = glpi_login()
            if not token:
                st.error("Failed to start GLPI session. Check tokens/URL.")
                st.stop()

            with st.spinner("Updating ticket and posting solution..."):
                glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                resp_sol = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)

            if str(resp_sol.status_code).startswith("2"):
                st.success(f"✅ Ticket #{ticket_id} updated: Category {category_id}, Status Solved (5), Solution posted (type 10).")
            else:
                st.error(f"❌ GLPI response: {resp_sol.status_code} — {resp_sol.text}")
else:
    st.info("📂 Please upload the Payment Excel to begin (Credit Note file optional).")
