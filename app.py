# ==========================================================
# The Remitator — GLPI Integration (FINAL • CN + Debug + Excel + Email→UserID)
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
GLPI_URL   = os.getenv("GLPI_URL")            # e.g. https://your.domain/apirest.php
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")
SOLUTION_CATEGORY_ID = 10   # AP Extras → PAYMENT REMITTANCE ADVICE

# ========== Helpers ==========
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

def glpi_add_solution(token, ticket_id, html, solution_type_id=SOLUTION_CATEGORY_ID):
    """
    Create a visible Solution (like clicking 'Add' in UI), then set AP Extras Solution Category = 10.
    """
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
    # If created, GLPI returns 201 with {"id": <solution_id>}
    if r.status_code in (200, 201):
        try:
            sol_id = r.json().get("id")
            if sol_id:
                # AP Extras plugin field: set Solution Category to 10
                requests.put(
                    f"{GLPI_URL}/PluginFieldsSolutioncategoryfield/{sol_id}",
                    json={"input": {"plugin_fields_solutioncategoryfields_id": SOLUTION_CATEGORY_ID}},
                    headers={"Session-Token": token, "App-Token": APP_TOKEN, "Content-Type": "application/json"}
                )
        except Exception:
            pass
    return r

# ========== Uploads ==========
pay_file = st.file_uploader("📂 Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("📂 (Optional) Upload Credit Notes Excel", type=["xlsx"])

# ========== Main ==========
if pay_file:
    try:
        df = pd.read_excel(pay_file)
        df.columns = [c.strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]
        st.success("✅ Payment file loaded successfully")
    except Exception as e:
        st.error(f"❌ Error loading Payment Excel: {e}")
        st.stop()

    required = ["Payment Document Code","Alt. Document","Invoice Value","Payment Value","Supplier Name","Supplier's Email"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in Payment Excel: {missing}")
        st.stop()

    pay_code = st.text_input("🔎 Enter Payment Document Code:")
    if not pay_code:
        st.stop()

    subset = df[df["Payment Document Code"].astype(str) == str(pay_code)].copy()
    if subset.empty:
        st.warning("⚠️ No rows found for this Payment Document Code.")
        st.stop()

    # Parse numeric
    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]

    # ===== CN matching (single + 2–3 combos) + debug =====
    summary = subset[["Alt. Document", "Invoice Value"]].copy()
    cn_rows, debug_rows, unmatched_invoices = [], [], []

    if cn_file:
        try:
            cn = pd.read_excel(cn_file)
            cn.columns = [c.strip() for c in cn.columns]
            cn = cn.loc[:, ~cn.columns.duplicated()]
            cn_alt_col = find_col(cn, ["Alt.Document", "Alt. Document"])
            cn_val_col = find_col(cn, ["Amount", "Debit", "Charge", "Cargo", "DEBE", "Invoice Value", "Invoice Value (€)"])
            if cn_alt_col and cn_val_col:
                cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)
                cn = cn[cn[cn_val_col].abs() > 0.01].reset_index(drop=True)
                cn = cn.drop_duplicates(subset=[cn_alt_col], keep="last").reset_index(drop=True)
                used_indices = set()

                for _, row in subset.iterrows():
                    inv = str(row["Alt. Document"])
                    payment_val = row["Payment Value"]
                    invoice_val = row["Invoice Value"]
                    diff = round(payment_val - invoice_val, 2)
                    matched_cns = []
                    if abs(diff) < 0.01:
                        debug_rows.append({"Invoice": inv, "Invoice Value": invoice_val,
                                           "Payment Value": payment_val, "Difference": diff,
                                           "Matched CNs": "—", "Matched?": "✅ (no diff)"})
                        continue

                    match_found = False
                    # Try single CN
                    for i, r in cn.iterrows():
                        if i in used_indices: continue
                        if round(abs(r[cn_val_col]), 2) == round(abs(diff), 2):
                            cn_no = str(r[cn_alt_col]); cn_amt = -abs(r[cn_val_col])
                            cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                            matched_cns.append(cn_no); used_indices.add(i); match_found = True; break
                    # Try 2–3 combo CNs
                    if not match_found:
                        available = [(i, abs(r[cn_val_col]), r) for i, r in cn.iterrows() if i not in used_indices]
                        for n in [2, 3]:
                            for combo in combinations(available, n):
                                total = round(sum(x[1] for x in combo), 2)
                                if abs(total - abs(diff)) < 0.05:
                                    for i, _, r in combo:
                                        cn_no = str(r[cn_alt_col]); cn_amt = -abs(r[cn_val_col])
                                        cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                                        matched_cns.append(cn_no); used_indices.add(i)
                                    match_found = True; break
                            if match_found: break
                    # Record diff if still unmatched
                    if not match_found:
                        unmatched_invoices.append({"Alt. Document": f"{inv} (Unmatched Diff)", "Invoice Value": diff})
                    debug_rows.append({
                        "Invoice": inv, "Invoice Value": invoice_val, "Payment Value": payment_val,
                        "Difference": diff, "Matched CNs": ", ".join(matched_cns) if matched_cns else "—",
                        "Matched?": "✅" if match_found else "❌"
                    })
                st.success(f"✅ Applied {len(cn_rows)} CNs (single/combo)")
            else:
                st.warning("⚠️ CN file missing expected columns ('Alt.Document', 'Amount/Debit'). CN logic skipped.")
        except Exception as e:
            st.warning(f"⚠️ Error loading CN file (will skip CN logic): {e}")

    # ===== Final table (Invoices + CNs + unmatched diffs) + correct TOTAL =====
    all_rows = summary.copy()
    if cn_rows:
        all_rows = pd.concat([all_rows, pd.DataFrame(cn_rows)], ignore_index=True)
    if unmatched_invoices:
        all_rows = pd.concat([all_rows, pd.DataFrame(unmatched_invoices)], ignore_index=True)

    total_val = subset["Payment Value"].sum()
    all_rows = pd.concat([all_rows, pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_val}])], ignore_index=True)

    display_df = all_rows.copy()
    display_df["Invoice Value (€)"] = display_df["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = display_df[["Alt. Document", "Invoice Value (€)"]]

    # ===== Tabs =====
    tab1, tab2 = st.tabs(["📋 Summary", "🔗 GLPI"])
    with tab1:
        st.subheader(f"Final Summary for Payment Code: {pay_code}")
        st.write(f"**Vendor:** {vendor}")
        st.write(f"**Vendor Email (from file):** {vendor_email_in_file}")
        st.dataframe(display_df, use_container_width=True)

        if debug_rows:
            st.subheader("🔍 Debug breakdown — invoice vs. CN matching")
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

        # Excel export with hidden meta
        wb = Workbook(); ws = wb.active; ws.title = "Final Summary"
        for r in dataframe_to_rows(display_df, index=False, header=True): ws.append(r)
        ws_hidden = wb.create_sheet("HiddenMeta")
        meta = [
            ["Vendor", vendor],
            ["Vendor Email", vendor_email_in_file],
            ["Payment Code", pay_code],
            ["Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ]
        for row in meta: ws_hidden.append(row)
        tab_ = Table(displayName="MetaTable", ref=f"A1:B{len(meta)}")
        tab_.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws_hidden.add_table(tab_); ws_hidden.sheet_state = "hidden"
        buf = BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button("💾 Download Excel Summary", buf,
                           file_name=f"{vendor}_Payment_{pay_code}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.write("Send the Spanish email + table to GLPI as a **Solution** (type 10), mark **Solved (5)**, set **Category 400**, and **assign by email**.")
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

        st.markdown("**Preview of email to send:**")
        st.markdown(html_message, unsafe_allow_html=True)

        if st.button("Send to GLPI"):
            if not all([GLPI_URL, APP_TOKEN, USER_TOKEN]):
                st.error("Missing GLPI credentials in .env."); st.stop()
            if not ticket_id.strip():
                st.error("Ticket ID required."); st.stop()

            token = glpi_login()
            if not token:
                st.error("Failed to start GLPI session."); st.stop()

            with st.spinner("Posting to GLPI..."):
                # 1) Resolve user by email → ID
                user_id = glpi_search_user(token, assigned_email)
                if not user_id:
                    st.error(f"❌ No GLPI user found for email: {assigned_email}")
                    st.stop()

                # 2) Update ticket (status + category)
                r1 = glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                # 3) Assign to resolved user
                r2 = glpi_assign_user(token, ticket_id, user_id)
                # 4) Create visible Solution + set AP Extras category=10
                r3 = glpi_add_solution(token, ticket_id, html_message, solution_type_id=SOLUTION_CATEGORY_ID)

            if all(str(r.status_code).startswith("2") or r.status_code == 201 for r in [r1, r2, r3]):
                st.success(f"✅ Ticket #{ticket_id} solved, category {category_id}, assigned to {assigned_email}, and solution posted (AP Extras = 10).")
            else:
                st.error(f"❌ GLPI error codes: {[r.status_code for r in [r1, r2, r3]]}")
else:
    st.info("📂 Please upload the Payment Excel to begin (Credit Note file optional).")
