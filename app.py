# ==========================================================
# The Remitator — GLPI Solution Sync (Final • CN Restored)
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
GLPI_URL   = os.getenv("GLPI_URL")        # e.g. https://glpi.domain/apirest.php
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

    # Parse amounts
    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor = subset["Supplier Name"].iloc[0]
    vendor_email_in_file = subset["Supplier's Email"].iloc[0]

    # Base summary (invoices)
    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    # ===== RESTORED CN LOGIC + DEBUG =====
    cn_rows, debug_rows, unmatched_invoices = [], [], []
    unmatched_cns_df = pd.DataFrame()

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
                        debug_rows.append({
                            "Invoice": inv, "Invoice Value": invoice_val, "Payment Value": payment_val,
                            "Difference": diff, "Matched CNs": "—", "Matched?": "✅ (no diff)"
                        })
                        continue

                    match_found = False

                    # 1) Single CN exact match
                    for i, r in cn.iterrows():
                        if i in used_indices: continue
                        if round(abs(r[cn_val_col]), 2) == round(abs(diff), 2):
                            cn_no = str(r[cn_alt_col])
                            cn_amt = -abs(r[cn_val_col]) if diff < 0 else abs(r[cn_val_col]) * (-1)
                            # Regardless of sign, present CN as negative adjustment
                            cn_amt = -abs(r[cn_val_col])
                            cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                            matched_cns.append(cn_no)
                            used_indices.add(i)
                            match_found = True
                            break

                    # 2) 2–3 CN combinations
                    if not match_found:
                        available = [(i, abs(r[cn_val_col]), r) for i, r in cn.iterrows() if i not in used_indices]
                        for n in [2, 3]:
                            for combo in combinations(available, n):
                                total = round(sum(x[1] for x in combo), 2)
                                if abs(total - abs(diff)) < 0.05:
                                    for i, _, r in combo:
                                        cn_no = str(r[cn_alt_col])
                                        cn_amt = -abs(r[cn_val_col])
                                        cn_rows.append({"Alt. Document": f"{cn_no} (CN)", "Invoice Value": cn_amt})
                                        matched_cns.append(cn_no)
                                        used_indices.add(i)
                                    match_found = True
                                    break
                            if match_found: break

                    # 3) If still not matched, record unmatched diff
                    if not match_found:
                        unmatched_invoices.append({
                            "Alt. Document": f"{inv} (Unmatched Diff)",
                            "Invoice Value": diff
                        })

                    debug_rows.append({
                        "Invoice": inv,
                        "Invoice Value": invoice_val,
                        "Payment Value": payment_val,
                        "Difference": diff,
                        "Matched CNs": ", ".join(matched_cns) if matched_cns else "—",
                        "Matched?": "✅" if match_found else "❌"
                    })

                # Unused CNs table
                unmatched_cns_df = cn.loc[~cn.index.isin(used_indices), [cn_alt_col, cn_val_col]].copy()
                unmatched_cns_df.rename(columns={cn_alt_col: "CN Number", cn_val_col: "Amount"}, inplace=True)
                unmatched_cns_df["Amount"] = unmatched_cns_df["Amount"].apply(lambda v: f"€{v:,.2f}")

                st.success(f"✅ Applied {len(cn_rows)} CNs (single/combo)")
            else:
                st.warning("⚠️ CN file missing expected columns ('Alt.Document', 'Amount/Debit'). CN logic skipped.")
        except Exception as e:
            st.warning(f"⚠️ Error loading CN file (will skip CN logic): {e}")

    # ===== Combine into final summary (Invoices + CNs + Unmatched Diffs) =====
    all_rows = summary.copy()
    if cn_rows:
        all_rows = pd.concat([all_rows, pd.DataFrame(cn_rows)], ignore_index=True)
    if unmatched_invoices:
        all_rows = pd.concat([all_rows, pd.DataFrame(unmatched_invoices)], ignore_index=True)

    # ✅ TOTAL should reflect actual payment sent
    total_val = subset["Payment Value"].sum()
    all_rows = pd.concat([all_rows, pd.DataFrame([{"Alt. Document": "TOTAL", "Invoice Value": total_val}])], ignore_index=True)

    # Pretty display
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

        # 🔍 Debug table restored
        if debug_rows:
            st.subheader("🔍 Debug breakdown — invoice vs. CN matching")
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)
        if not unmatched_cns_df.empty:
            st.subheader("🧾 Unused CNs")
            st.dataframe(unmatched_cns_df, use_container_width=True)

        # Optional Excel export of the display table + meta
        wb = Workbook(); ws = wb.active; ws.title = "Final Summary"
        for r in dataframe_to_rows(display_df, index=False, header=True): ws.append(r)
        ws_hidden = wb.create_sheet("HiddenMeta")
        meta = [["Vendor", vendor],["Vendor Email", vendor_email_in_file],["Payment Code", pay_code],
                ["Exported At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]]
        for row in meta: ws_hidden.append(row)
        tab_ = Table(displayName="MetaTable", ref=f"A1:B{len(meta)}")
        tab_.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws_hidden.add_table(tab_); ws_hidden.sheet_state = "hidden"
        buf = BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button("💾 Download Excel Summary", buf,
            file_name=f"{vendor}_Payment_{pay_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===== GLPI Tab =====
    with tab2:
        st.write("This will send the email (Spanish) with the table to GLPI as a **Solution** and mark the ticket **Solved** (Category 400, Solution Type 10).")
        c1, c2 = st.columns(2)
        ticket_id = c1.text_input("Ticket ID", placeholder="101004")
        category_id = c2.text_input("Category ID", value="400")

        # Spanish email body + table
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
                st.error("Missing GLPI credentials in .env (GLPI_URL, APP_TOKEN, USER_TOKEN)."); st.stop()
            if not ticket_id.strip():
                st.error("Ticket ID is required."); st.stop()

            token = glpi_login()
            if not token:
                st.error("Failed to start GLPI session. Check tokens/URL."); st.stop()

            with st.spinner("Updating ticket and posting solution..."):
                glpi_update_ticket(token, ticket_id, status=5, category_id=int(category_id))
                resp_sol = glpi_add_solution(token, ticket_id, html_message, solution_type_id=10)

            if str(resp_sol.status_code).startswith("2"):
                st.success(f"✅ Ticket #{ticket_id} updated: Category {category_id}, Status Solved (5), Solution posted (type 10).")
            else:
                st.error(f"❌ GLPI response: {resp_sol.status_code} — {resp_sol.text}")
else:
    st.info("📂 Please upload the Payment Excel to begin (Credit Note file optional).")
