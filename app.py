# ==========================================================
# THE REMITATOR — OLD FINAL HYBRID VERSION
# OLD FINAL + ADVANCED DEBUG B (✓/✗)
# ==========================================================

import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime
from dotenv import load_dotenv

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.set_page_config(page_title="The Remitator", layout="wide")
st.title("💀 The Remitator — Old Final Hybrid")

# ----------------------------------------------------------
# ENV
# ----------------------------------------------------------
load_dotenv()
GLPI_URL = os.getenv("GLPI_URL")
APP_TOKEN = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")

# ----------------------------------------------------------
# HELPERS
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

def find_col(df, names):
    for c in df.columns:
        clean = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            if n.replace(" ", "").replace(".", "").lower() in clean:
                return c
    return None

# ----------------------------------------------------------
# GLPI BASIC FUNCTIONS
# ----------------------------------------------------------
def glpi_login():
    r = requests.get(
        f"{GLPI_URL}/initSession",
        headers={"Authorization": f"user_token {USER_TOKEN}", "App-Token": APP_TOKEN}
    )
    return r.json().get("session_token")

def glpi_update_ticket(token, ticket_id, status=5, category_id=None):
    payload = {"input": {"status": int(status)}}
    if category_id:
        payload["input"]["itilcategories_id"] = int(category_id)
    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

def glpi_add_solution(token, ticket_id, html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": html,
            "solutiontypes_id": 10,
            "status": 5
        }
    }
    return requests.post(
        f"{GLPI_URL}/ITILSolution",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )

def glpi_add_followup(token, ticket_id, html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "content": html,
            "solutiontypes_id": 10
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILFollowup",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN}
    )


# ----------------------------------------------------------
# USER MAP (OLD FINAL)
# ----------------------------------------------------------
USER_MAP = {
    "akeramaris@saniikos.com": 22487,
    "mmarquis@saniikos.com": 16207
}

# ----------------------------------------------------------
# MAIN INPUTS
# ----------------------------------------------------------
pay_file = st.file_uploader("Upload Payment Excel", type=["xlsx"])
cn_file = st.file_uploader("Upload Credit Notes (optional)", type=["xlsx"])

if not pay_file:
    st.info("Upload Payment Excel to start.")
    st.stop()

df = pd.read_excel(pay_file)
df.columns = [c.strip() for c in df.columns]
df = df.loc[:, ~df.columns.duplicated()]

pay_input = st.text_input("Enter Payment Document Codes (comma separated):")
if not pay_input.strip():
    st.stop()

selected_codes = [x.strip() for x in pay_input.split(",") if x.strip()]
if not selected_codes:
    st.stop()

combined_html = ""
combined_vendor_names = []
export_data = {}
debug_rows_all = []   # <--- FOR ADVANCED DEBUG B

# ----------------------------------------------------------
# PROCESS EACH PAYMENT CODE
# ----------------------------------------------------------
for pay_code in selected_codes:
    col = find_col(df, ["PaymentDocumentCode", "PaymentDocument"])
    if not col:
        st.error("Cannot find Payment Document Code column.")
        st.stop()

    subset = df[df[col].astype(str) == str(pay_code)]
    if subset.empty:
        continue

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor_col = find_col(df, ["Vendor", "SupplierName", "Supplier"])
    vendor = subset[vendor_col].iloc[0] if vendor_col else "Unknown Vendor"

    combined_vendor_names.append(vendor)

    summary = subset[["Alt. Document", "Invoice Value"]].copy()

    cn_rows = []
    unmatched = []

    # ------------------------------------------------------
    # CREDIT NOTES MATCHING
    # ------------------------------------------------------
    if cn_file:
        cn = pd.read_excel(cn_file)
        cn.columns = [c.strip() for c in cn.columns]
        cn = cn.loc[:, ~cn.columns.duplicated()]

        cn_alt = find_col(cn, ["AltDocument", "Alt.Document"])
        cn_val = find_col(cn, ["Amount", "InvoiceValue", "DEBE", "Cargo"])

        if cn_alt and cn_val:
            cn[cn_val] = cn[cn_val].apply(parse_amount)
            used = set()

            for _, row in subset.iterrows():
                inv = str(row["Alt. Document"])
                inv_val = row["Invoice Value"]
                pay_val = row["Payment Value"]
                diff = round(pay_val - inv_val, 2)

                # Debug storage
                debug_entry = {
                    "Payment Code": pay_code,
                    "Vendor": vendor,
                    "Alt. Document": inv,
                    "Invoice Value": inv_val,
                    "Payment Value": pay_val,
                    "Difference": diff,
                    "Matched": "✓" if abs(diff) < 0.01 else "✗"
                }

                # Try match CN
                matched = False
                for i, r in cn.iterrows():
                    if i in used: continue
                    if round(abs(r[cn_val]),2) == round(abs(diff),2):
                        cn_rows.append({
                            "Alt. Document": f"{r[cn_alt]} (CN)",
                            "Invoice Value": -abs(r[cn_val])
                        })
                        used.add(i)
                        matched = True
                        break

                if not matched and abs(diff) > 0.01:
                    unmatched.append({
                        "Alt. Document": f"{inv} (Adj. Diff)",
                        "Invoice Value": diff
                    })

                debug_rows_all.append(debug_entry)

    # ------------------------------------------------------
    # FINAL ROW TABLE
    # ------------------------------------------------------
    full = pd.concat([
        summary,
        pd.DataFrame(cn_rows),
        pd.DataFrame(unmatched)
    ], ignore_index=True)

    total_value = full["Invoice Value"].sum()
    full.loc[len(full)] = ["TOTAL", total_value]

    export_data[pay_code] = {"vendor": vendor, "rows": full.copy()}

    display_df = full.copy()
    display_df["Invoice Value (€)"] = display_df["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
    display_df = display_df[["Alt. Document", "Invoice Value (€)"]]

    combined_html += f"""
<b>Payment Code:</b> {pay_code}<br>
<b>Vendor:</b> {vendor}<br>
<b>Total Amount:</b> €{total_value:,.2f}<br><br>
{display_df.to_html(index=False, border=0)}
<br><hr><br>
"""


# ----------------------------------------------------------
# OUTPUT SUMMARY
# ----------------------------------------------------------
if combined_html.endswith("<br><hr><br>"):
    combined_html = combined_html[:-12]

tab1, tab2, tab3 = st.tabs(["Summary", "Advanced Debug", "GLPI"])

# ----------------------------------------------------------
# TAB 1 — SUMMARY
# ----------------------------------------------------------
with tab1:
    st.markdown(combined_html, unsafe_allow_html=True)

# ----------------------------------------------------------
# TAB 2 — ADVANCED DEBUG (VERSION B)
# ----------------------------------------------------------
with tab2:
    st.subheader("Advanced Debug Breakdown (Unicode Icons ✓ / ✗)")
    dbg_df = pd.DataFrame(debug_rows_all)
    dbg_df = dbg_df.sort_values(by=["Payment Code", "Vendor", "Alt. Document"]).reset_index(drop=True)
    st.dataframe(dbg_df, use_container_width=True)

    st.download_button(
        "⬇️ Download Debug CSV",
        dbg_df.to_csv(index=False).encode("utf-8"),
        file_name="debug_breakdown.csv",
        mime="text/csv"
    )


# ----------------------------------------------------------
# TAB 3 — GLPI
# ----------------------------------------------------------
with tab3:
    language = st.radio("Language", ["Spanish", "English"], horizontal=True)

    ticket_id = st.text_input("Ticket ID")
    category_id = st.text_input("Category ID")
    assigned_email = st.text_input("Assign Email (optional)")

    if language == "Spanish":
        intro = "Estimado proveedor,<br><br>Adjuntamos las facturas correspondientes a los pagos realizados:<br><br>"
        outro = "<br>Quedamos a su disposición para cualquier aclaración.<br><br>Saludos,<br>Finance"
    else:
        intro = "Dear supplier,<br><br>Please find below the invoices corresponding to the executed payments:<br><br>"
        outro = "<br>Should you need any clarification, we remain available.<br><br>Regards,<br>Finance Team"

    html_message = intro + combined_html + outro
    st.markdown(html_message, unsafe_allow_html=True)

    if st.button("Send to GLPI"):
        if not ticket_id.isdigit():
            st.error("Invalid Ticket ID")
            st.stop()

        token = glpi_login()
        if not token:
            st.error("GLPI login error.")
            st.stop()

        glpi_update_ticket(token, ticket_id, 5, category_id)

        resp = glpi_add_solution(token, ticket_id, html_message)

        if resp.status_code == 400 or "already solved" in resp.text.lower():
            glpi_add_followup(token, ticket_id, html_message)
            st.warning("Ticket solved already — posted as follow-up.")
        else:
            st.success("Solution added.")
