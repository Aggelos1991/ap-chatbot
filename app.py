# ==========================================================
# THE REMITATOR — OLD FINAL HYBRID VERSION (FIXED)
# OLD FINAL + ADVANCED DEBUG B (✓/✗) + CN GROUPING FIX
# ==========================================================

import os, re, requests
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime
from dotenv import load_dotenv

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.set_page_config(page_title="The Remitator", layout="wide")
st.title("💀 The Remitator — Old Final Hybrid (CN Fix)")

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
        cc = c.strip().lower().replace(" ", "").replace(".", "")
        for n in names:
            nn = n.strip().lower().replace(" ", "").replace(".", "")
            if nn in cc:
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
cn_file  = st.file_uploader("Upload Credit Notes (optional)", type=["xlsx"])

if not pay_file:
    st.info("Upload Payment Excel to begin.")
    st.stop()

df = pd.read_excel(pay_file)
df.columns = [c.strip() for c in df.columns]
df = df.loc[:, ~df.columns.duplicated()]

pay_input = st.text_input("Enter Payment Document Codes (comma separated):")
if not pay_input.strip(): st.stop()

selected_codes = [c.strip() for c in pay_input.split(",") if c.strip()]
if not selected_codes: st.stop()

combined_html = ""
combined_vendor_names = []
export_data = {}
debug_rows_all = []  # for ADVANCED DEBUG B

# ----------------------------------------------------------
# PROCESS EACH PAYMENT CODE
# ----------------------------------------------------------
for pay_code in selected_codes:

    col = find_col(df, ["paymentdocumentcode", "paymentdocument"])
    if not col:
        st.error("Payment Document column not found.")
        st.stop()

    subset = df[df[col].astype(str) == str(pay_code)]
    if subset.empty:
        continue

    subset["Invoice Value"] = subset["Invoice Value"].apply(parse_amount)
    subset["Payment Value"] = subset["Payment Value"].apply(parse_amount)

    vendor_col = find_col(df, ["vendor", "suppliername", "supplier"])
    vendor = subset[vendor_col].iloc[0] if vendor_col else "Unknown Vendor"
    combined_vendor_names.append(vendor)

    summary_df = subset[["Alt. Document", "Invoice Value"]].copy()

    cn_rows = []
    unmatched_rows = []

    # ------------------------------------------------------
    # CN GROUPING FIX — only last correction counts
    # ------------------------------------------------------
    grouped_cn = None
    cn_alt = None
    cn_val_col = None

    if cn_file:
        cn = pd.read_excel(cn_file)
        cn.columns = [c.strip() for c in cn.columns]
        cn = cn.loc[:, ~cn.columns.duplicated()]

        cn_alt = find_col(cn, ["altdocument", "alt.document"])
        cn_val_col = find_col(cn, ["amount", "invoicevalue", "debe", "cargo"])

        if cn_alt and cn_val_col:
            cn[cn_val_col] = cn[cn_val_col].apply(parse_amount)

            # 🟩 CRITICAL FIX — keep ONLY last correction per Alt.Document
            grouped_cn = (
                cn
                .sort_values(by=cn.index)
                .groupby(cn_alt, as_index=False)
                .last()
            )

    # ------------------------------------------------------
    # PROCESS INVOICES (NO GROUPING HERE)
    # ------------------------------------------------------
    for _, row in subset.iterrows():
        inv = str(row["Alt. Document"])
        inv_val = row["Invoice Value"]
        pay_val = row["Payment Value"]
        diff = round(pay_val - inv_val, 2)

        debug_entry = {
            "Payment Code": pay_code,
            "Vendor": vendor,
            "Alt. Document": inv,
            "Invoice Value": inv_val,
            "Payment Value": pay_val,
            "Difference": diff,
            "Matched": "✓" if abs(diff) < 0.01 else "✗"
        }

        matched = False

        # --------------------------------------------------
        # CN MATCHING (using ONLY grouped CN file)
        # --------------------------------------------------
        if grouped_cn is not None:
            for _, cn_row in grouped_cn.iterrows():

                cn_val = abs(cn_row[cn_val_col])
                if round(cn_val,2) == round(abs(diff),2):

                    cn_rows.append({
                        "Alt. Document": f"{cn_row[cn_alt]} (CN)",
                        "Invoice Value": -cn_val
                    })
                    matched = True
                    break

        if not matched and abs(diff) > 0.01:
            unmatched_rows.append({
                "Alt. Document": f"{inv} (Adj. Diff)",
                "Invoice Value": diff
            })

        debug_rows_all.append(debug_entry)

    # ------------------------------------------------------
    # BUILD FINAL TABLE BLOCK
    # ------------------------------------------------------
    final_df = pd.concat(
        [summary_df, pd.DataFrame(cn_rows), pd.DataFrame(unmatched_rows)],
        ignore_index=True
    )

    total_value = final_df["Invoice Value"].sum()
    final_df.loc[len(final_df)] = ["TOTAL", total_value]

    export_data[pay_code] = {"vendor": vendor, "rows": final_df.copy()}

    html_df = final_df.copy()
    html_df["Invoice Value (€)"] = html_df["Invoice Value"].apply(lambda v: f"€{v:,.2f}")

    html_df = html_df[["Alt. Document", "Invoice Value (€)"]]

    combined_html += f"""
<b>Payment Code:</b> {pay_code}<br>
<b>Vendor:</b> {vendor}<br>
<b>Total Amount:</b> €{total_value:,.2f}<br><br>
{html_df.to_html(index=False, border=0)}
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
# TAB 2 — ADVANCED DEBUG (✓ / ✗)
# ----------------------------------------------------------
with tab2:
    st.subheader("Advanced Debug Breakdown (Unicode ✓ / ✗)")
    dbg_df = pd.DataFrame(debug_rows_all)
    dbg_df = dbg_df.sort_values(by=["Payment Code", "Vendor", "Alt. Document"])
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
    email_assign = st.text_input("Assign To Email")

    if language == "Spanish":
        intro = "Estimado proveedor,<br><br>Adjuntamos las facturas correspondientes a los pagos realizados:<br><br>"
        outro = "<br>Quedamos a su disposición para cualquier aclaración.<br><br>Saludos,<br>Finance"
    else:
        intro = "Dear supplier,<br><br>Please find below the invoices related to the executed payments:<br><br>"
        outro = "<br>Should you require any clarification, we remain available.<br><br>Finance Team"

    html_message = intro + combined_html + outro
    st.markdown(html_message, unsafe_allow_html=True)

    if st.button("Send to GLPI"):
        if not ticket_id.isdigit():
            st.error("Invalid Ticket ID.")
            st.stop()

        token = glpi_login()
        if not token:
            st.error("GLPI Login Failed")
            st.stop()

        glpi_update_ticket(token, ticket_id, 5, category_id)

        resp = glpi_add_solution(token, ticket_id, html_message)

        if resp.status_code == 400 or "already solved" in resp.text.lower():
            glpi_add_followup(token, ticket_id, html_message)
            st.warning("Ticket was solved — posted as follow-up.")
        else:
            st.success("Solution added.")
