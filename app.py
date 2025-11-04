# ==========================================================
# The Remitator — FINAL FINAL (Manual Comma Codes + AP Extras Solution Category ID=10)
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
st.title("The Remitator — Hasta la vista, payment remittance.")
# ========== ENV ==========
load_dotenv()
GLPI_URL = os.getenv("GLPI_URL")
APP_TOKEN = os.getenv("APP_TOKEN")
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
def glpi_set_apextras_category(token, ticket_id, solution_cat_id=10):
    body = {
        "input": {
            "id": int(ticket_id),
            "plugin_fields_solutioncategoryfielddropdowns_id": int(solution_cat_id)
        }
    }
    return requests.put(
        f"{GLPI_URL}/Ticket/{int(ticket_id)}",
        json=body,
        headers={
            "Session-Token": token,
            "App-Token": APP_TOKEN,
            "Content-Type": "application/json"
        },
        timeout=30
    )
def glpi_add_solution(token, ticket_id, html, solution_type_id=10):
    body = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": html,
            "solutiontypes_id": int(solution_type_id),
            "status": 5
        }
    }
    return requests.post(
        f"{GLPI_URL}/ITILSolution",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"},
        timeout=30
    )
def glpi_add_followup(token, ticket_id, html):
    body = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": html,
            "solutiontypes_id": 10
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILFollowup",
        json=body,
        headers={"Session-Token": token,
                "App-Token": APP_TOKEN,
                "Content-Type": "application/json"
        },
        timeout=30
    )
def glpi_assign_userid(token, ticket_id, user_id):
    tid = int(str(ticket_id).strip())
    body = {
        "input": {
            "tickets_id": tid,
            "users_id": int(user_id),
            "type": 2,
            "use_notification": 1
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{tid}/Ticket_User",
        json=body,
        headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"},
        timeout=30
    )
# ========== USER MAP ==========
USER_MAP = {
    "akeramaris@saniikos.com": 22487,
    "mmarquis@saniikos.com": 16207
}
# ========== MAIN ==========
pay_file = st.file_uploader("Upload Payment Excel", type=["xlsx"])
cn_file = st.file_uploader("(Optional) Upload Credit Notes Excel", type=["xlsx"])
if pay_file:
    df = pd.read_excel(pay_file)
    df.columns = [c.strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    st.success("Payment file loaded successfully")
    pay_input = st.text_input("Enter one or more Payment Document Codes (comma-separated):", "")
    if not pay_input.strip():
        st.stop()
    selected_codes = [x.strip() for x in pay_input.split(",") if x.strip()]
    if not selected_codes:
        st.stop()
    combined_html = ""
    combined_vendor_names = []
    debug_rows_all = []
    export_tables = {}
    for pay_code in selected_codes:
        subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
        if subset.empty:
            continue
        subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
        subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)
        vendor = subset["Supplier Name"].iloc[0]
        vendor_email_in_file = subset["Supplier's Email"].iloc[0]
        summary = subset[["Alt. Document", "Invoice Value"]].copy()
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
                        val = round(abs(r[cn_val_col]), 2)
                        if val == 0: continue
                        if round(val, 2) == round(abs(diff), 2):
                            cn_rows.append({"Alt. Document": f"{r[cn_alt_col]} (CN)", "Invoice Value": -val})
                            used.add(i); match=True; break
                    if not match and abs(diff) > 0.01:
                        unmatched_invoices.append({"Alt. Document": f"{inv} (Adj. Diff)", "Invoice Value": diff})
                    debug_rows.append({
                        "Invoice": inv, "Invoice Value": row["Invoice Value"],
                        "Payment Value": row["Payment Value"], "Difference": diff,
                        "Matched?": "Yes" if match else "No"
                    })
        valid_cn_df = pd.DataFrame([r for r in cn_rows if r["Invoice Value"] != 0])
        unmatched_df = pd.DataFrame(unmatched_invoices)
        all_rows = pd.concat([summary, valid_cn_df, unmatched_df], ignore_index=True)
        total_val = subset["Payment Value"].sum()
        all_rows.loc[len(all_rows)] = ["TOTAL", total_val]
        all_rows["Invoice Value (€)"] = all_rows["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
        display_df = all_rows[["Alt. Document", "Invoice Value (€)"]]
        debug_rows_all.extend(debug_rows)
        html_table = display_df.to_html(index=False, border=0, justify="center", classes="table")
        combined_html += f"<h4>Payment Code: {pay_code} — Vendor: {vendor}</h4>{html_table}<br>"
        combined_vendor_names.append(vendor)
        export_tables[pay_code] = display_df
    tab1, tab2 = st.tabs(["Summary", "GLPI"])
    with tab1:
        st.markdown(combined_html, unsafe_allow_html=True)
        if debug_rows_all:
            st.subheader("Debug breakdown — invoice vs. CN matching")
            st.dataframe(pd.DataFrame(debug_rows_all), use_container_width=True)

        # EXCEL FIXED — ONLY THIS PART CHANGED
        if export_tables:
            from openpyxl.styles import Font, Alignment, PatternFill
            wb = Workbook()
            ws = wb.active
            ws.title = "Payment Summary"
            ws.append(["The Remitator – Payment Summary"])
            ws.append([f"Payment Codes: {', '.join(selected_codes)}"])
            ws.append([f"Vendors: {', '.join(set(combined_vendor_names))}"])
            ws.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
            ws.append([])
            bold = Font(bold=True)
            money = '#,##0.00 €'
            row = 6
            for code in selected_codes:
                vendor = df[df["Payment Document Code"].astype(str) == code]["Supplier Name"].iloc[0]
                block = export_tables[code]
                ws.cell(row,1).value = f"Payment Code: {code} – {vendor}"
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
                ws.cell(row,1).font = bold
                row += 2
                ws.cell(row,1).value = "Document"
                ws.cell(row,2).value = "Amount (€)"
                for c in (1,2):
                    ws.cell(row,c).font = Font(bold=True,color="FFFFFF")
                    ws.cell(row,c).fill = PatternFill("solid","1E88E5")
                row += 1
                total = 0.0
                for _, r in block.iterrows():
                    doc = r["Alt. Document"]
                    amt_str = r["Invoice Value (€)"].replace("€","").replace(",","").strip()
                    amt = float(amt_str) if amt_str else 0.0
                    total += amt
                    ws.cell(row,1).value = doc
                    ws.cell(row,2).value = amt
                    ws.cell(row,2).number_format = money
                    if "(CN)" in doc: ws.cell(row,1).font = Font(color="2E8B57")
                    if "(Adj." in doc: ws.cell(row,1).font = Font(color="D32F2F")
                    row += 1
                ws.cell(row,1).value = "TOTAL"
                ws.cell(row,2).value = total
                ws.cell(row,1).font = bold
                ws.cell(row,2).font = bold
                ws.cell(row,2).number_format = money
                ws.cell(row,2).fill = PatternFill("solid","E3F2FD")
                row += 2
            for col in ws.columns:
                max_len = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len+2,50)
            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            st.download_button(
                "Download Excel Summary (FIXED)",
                buf,
                f"Remittance_{'_'.join(selected_codes)}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        # END OF EXCEL FIX

    with tab2:
        c1, c2, c3 = st.columns(3)
        ticket_id = c1.text_input("Ticket ID", placeholder="101004")
        category_id = c2.text_input("Category ID", value="400")
        assigned_email = c3.text_input("Assign To Email", placeholder="akeramaris@saniikos.com")
        html_message = f"""
        <p><strong>Estimado proveedor,</strong></p>
        <p>Por favor, encuentre a continuación las facturas que corresponden a los pagos realizados:</p>
        {combined_html}
        <p>Quedamos a su disposición para cualquier aclaración.</p>
        <p>Saludos cordiales,<br><strong>Equipo Finance</strong></p>
        """
        st.markdown(html_message, unsafe_allow_html=True)
        if st.button("Send to GLPI"):
            if not str(ticket_id).strip().isdigit():
                st.error("Invalid or empty Ticket ID. Please enter a numeric ID.")
                st.stop()
            if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
                st.error("Missing GLPI credentials."); st.stop()
            token = glpi_login()
            if not token:
                st.error("Failed GLPI session."); st.stop()
            user_id = USER_MAP.get(assigned_email.lower())
            if not user_id:
                st.error(f"No mapped GLPI user ID for email: {assigned_email}")
                st.stop()
            with st.spinner("Posting to GLPI..."):
                glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                glpi_set_apextras_category(token, ticket_id, solution_cat_id=10)
                glpi_assign_userid(token, ticket_id, user_id)
                resp_sol = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)
                if resp_sol.status_code == 400 or "already solved" in resp_sol.text.lower():
                    st.warning("Ticket already solved — posting as comment instead.")
                    resp_sol = glpi_add_followup(token, ticket_id, html_message)
            if str(resp_sol.status_code).startswith("2"):
                st.success(f"Ticket #{ticket_id} updated — Solution/Comment added successfully (AP Extras ID 10).")
            else:
                st.error(f"GLPI error: {resp_sol.status_code} → {resp_sol.text}")
else:
    st.info("Upload Payment Excel to begin.")
