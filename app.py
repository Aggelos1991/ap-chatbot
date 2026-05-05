# ==========================================================
# THE REMITATOR — FINAL HYBRID (HARDENED GLPI EDITION)
# DEFAULT USER ID = 22487 (ANGELOS KERAMARIS)
# ==========================================================

import os, re, requests
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.set_page_config(page_title="The Remitator", layout="wide")
st.title("💀 The Remitator — Final Hybrid (Hardened)")

DEFAULT_USER_ID = 22487
SIGNATURE = """<br><br>Saludos,<br><b>Angelos Keramaris<br>Accounts Payable Iberia</b>"""

# ----------------------------------------------------------
# ENV
# ----------------------------------------------------------
load_dotenv()
GLPI_URL   = (os.getenv("GLPI_URL") or "").rstrip("/")
APP_TOKEN  = os.getenv("APP_TOKEN")
USER_TOKEN = os.getenv("USER_TOKEN")


# ----------------------------------------------------------
# HELPERS
# ----------------------------------------------------------
def parse_amount(v):
    if pd.isna(v):
        return 0.0
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
            n_clean = n.lower().replace(" ", "").replace(".", "")
            if n_clean in clean:
                return c
    return None


def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return None


# ----------------------------------------------------------
# GLPI FUNCTIONS — HARDENED
# ----------------------------------------------------------
def glpi_login():
    if not GLPI_URL or not APP_TOKEN or not USER_TOKEN:
        return None, (
            "Missing GLPI credentials. Check your .env / Streamlit secrets:\n"
            f"- GLPI_URL set:   {bool(GLPI_URL)}\n"
            f"- APP_TOKEN set:  {bool(APP_TOKEN)}\n"
            f"- USER_TOKEN set: {bool(USER_TOKEN)}"
        )

    try:
        r = requests.get(
            f"{GLPI_URL}/initSession",
            headers={
                "Authorization": f"user_token {USER_TOKEN}",
                "App-Token": APP_TOKEN,
                "Content-Type": "application/json",
            },
            timeout=20,
        )
    except requests.RequestException as e:
        return None, f"Network error contacting GLPI: {e}"

    data = safe_json(r)

    if isinstance(data, list):
        return None, f"GLPI rejected login: {data}"

    if not isinstance(data, dict):
        return None, (
            f"Unexpected GLPI response (status {r.status_code}). "
            f"Body: {r.text[:300]}"
        )

    token = data.get("session_token")
    if not token:
        return None, f"GLPI response had no session_token. Body: {data}"

    return token, None


def glpi_update_ticket(token, ticket_id, status=5, category_id=None):
    payload = {
        "input": {
            "id": int(ticket_id),
            "status": int(status),
            "users_id_lastupdater": DEFAULT_USER_ID,
            "users_id_recipient": DEFAULT_USER_ID,
        }
    }
    if category_id:
        payload["input"]["itilcategories_id"] = int(category_id)

    return requests.put(
        f"{GLPI_URL}/Ticket/{ticket_id}",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN},
        timeout=20,
    )


def glpi_add_solution(token, ticket_id, html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "users_id": DEFAULT_USER_ID,
            "users_id_recipient": DEFAULT_USER_ID,
            "content": html,
            "solutiontypes_id": 10,
            "status": 5,
        }
    }
    return requests.post(
        f"{GLPI_URL}/ITILSolution",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN},
        timeout=20,
    )


def glpi_add_followup(token, ticket_id, html):
    payload = {
        "input": {
            "itemtype": "Ticket",
            "items_id": int(ticket_id),
            "users_id": DEFAULT_USER_ID,
            "users_id_recipient": DEFAULT_USER_ID,
            "content": html,
        }
    }
    return requests.post(
        f"{GLPI_URL}/Ticket/{ticket_id}/ITILFollowup",
        json=payload,
        headers={"Session-Token": token, "App-Token": APP_TOKEN},
        timeout=20,
    )


def glpi_kill_session(token):
    try:
        requests.get(
            f"{GLPI_URL}/killSession",
            headers={"Session-Token": token, "App-Token": APP_TOKEN},
            timeout=10,
        )
    except Exception:
        pass


# ----------------------------------------------------------
# INPUT FILES
# ----------------------------------------------------------
pay_file = st.file_uploader("Upload Payment Excel", type=["xlsx"])
cn_file  = st.file_uploader("Upload Credit Notes (optional)", type=["xlsx"])

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

# ----------------------------------------------------------
# COLUMN DETECTION (once, up front)
# ----------------------------------------------------------
pay_doc_col = find_col(df, ["PaymentDocumentCode", "PaymentDocument"])
if not pay_doc_col:
    st.error("Cannot find Payment Document Code column.")
    st.stop()

alt_col    = find_col(df, ["Alt.Document", "AltDocument", "Alt. Document"]) or "Alt. Document"
inv_col    = find_col(df, ["InvoiceValue", "Invoice Value"]) or "Invoice Value"
payv_col   = find_col(df, ["PaymentValue", "Payment Value"]) or "Payment Value"
vendor_col = find_col(df, ["Vendor", "SupplierName", "Supplier"])

# ----------------------------------------------------------
# CREDIT NOTES — load ONCE outside loop
# ----------------------------------------------------------
cn_df = None
cn_alt = cn_val = None
if cn_file:
    cn_df = pd.read_excel(cn_file)
    cn_df.columns = [c.strip() for c in cn_df.columns]
    cn_df = cn_df.loc[:, ~cn_df.columns.duplicated()]
    cn_alt = find_col(cn_df, ["AltDocument", "Alt.Document", "Alt. Document"])
    cn_val = find_col(cn_df, ["Amount", "InvoiceValue", "DEBE", "Cargo"])
    if cn_alt and cn_val:
        cn_df[cn_val] = cn_df[cn_val].apply(parse_amount)
    else:
        st.warning(
            f"Credit-notes file uploaded but columns not detected "
            f"(alt={cn_alt!r}, val={cn_val!r}). CN matching disabled."
        )
        cn_df = None

cn_used_global = set()

combined_html = ""
export_data = {}
debug_rows_all = []

# ----------------------------------------------------------
# PROCESS EACH PAYMENT CODE
# ----------------------------------------------------------
for pay_code in selected_codes:
    subset = df[df[pay_doc_col].astype(str) == str(pay_code)].copy()
    if subset.empty:
        continue

    subset[inv_col]  = subset[inv_col].apply(parse_amount)
    subset[payv_col] = subset[payv_col].apply(parse_amount)

    vendor = subset[vendor_col].iloc[0] if vendor_col else "Unknown Vendor"

    summary_rows = []
    cn_rows = []
    unmatched = []

    for _, row in subset.iterrows():
        inv     = str(row[alt_col])
        inv_val = row[inv_col]
        pay_val = row[payv_col]
        diff    = round(pay_val - inv_val, 2)

        summary_rows.append({"Alt. Document": inv, "Invoice Value": inv_val})

        dbg = {
            "Payment Code": pay_code,
            "Vendor": vendor,
            "Alt. Document": inv,
            "Invoice Value": inv_val,
            "Payment Value": pay_val,
            "Difference": diff,
            "Matched CN": "",
            "CN Value": "",
            "Status": "",
        }

        if abs(diff) < 0.01:
            dbg["Status"] = "✓ Exact"
            debug_rows_all.append(dbg)
            continue

        matched = False
        if cn_df is not None:
            for i, r in cn_df.iterrows():
                if i in cn_used_global:
                    continue
                if round(abs(r[cn_val]), 2) == round(abs(diff), 2):
                    cn_doc = str(r[cn_alt])
                    cn_amount = -abs(r[cn_val]) if diff < 0 else abs(r[cn_val])
                    cn_rows.append({
                        "Alt. Document": f"{cn_doc} (CN)",
                        "Invoice Value": cn_amount,
                    })
                    cn_used_global.add(i)
                    matched = True
                    dbg["Status"]     = "✓ CN matched"
                    dbg["Matched CN"] = cn_doc
                    dbg["CN Value"]   = cn_amount
                    break

        if not matched:
            unmatched.append({
                "Alt. Document": f"{inv} (Adj. Diff)",
                "Invoice Value": diff,
            })
            dbg["Status"] = "✗ No CN — adjustment"

        debug_rows_all.append(dbg)

    full = pd.concat(
        [pd.DataFrame(summary_rows), pd.DataFrame(cn_rows), pd.DataFrame(unmatched)],
        ignore_index=True,
    )

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
# FINAL OUTPUT
# ----------------------------------------------------------
if combined_html.endswith("<br><hr><br>"):
    combined_html = combined_html[:-12]

tab1, tab2, tab3 = st.tabs(["Summary", "Advanced Debug", "GLPI"])

# ------------------------ SUMMARY -------------------------
with tab1:
    st.markdown(combined_html, unsafe_allow_html=True)

# ------------------------ DEBUG ---------------------------
with tab2:
    if debug_rows_all:
        dbg_df = pd.DataFrame(debug_rows_all)
        st.dataframe(dbg_df, use_container_width=True)

        st.download_button(
            "⬇️ Download Debug CSV",
            dbg_df.to_csv(index=False).encode("utf-8"),
            file_name="debug_breakdown.csv",
            mime="text/csv",
        )
    else:
        st.info("No rows processed — check that the Payment Document Codes match the Excel.")

# ------------------------ GLPI OUTPUT ----------------------
with tab3:
    language = st.radio("Language", ["Spanish", "English"], horizontal=True)

    ticket_id   = st.text_input("Ticket ID")
    category_id = st.text_input("Category ID")

    if language == "Spanish":
        intro = (
            "Estimado proveedor,<br><br>"
            "Adjuntamos el detalle de facturas correspondientes a los pagos realizados:"
            "<br><br>"
        )
        outro = SIGNATURE
    else:
        intro = (
            "Dear supplier,<br><br>"
            "Please find below the invoice breakdown corresponding to the executed payments:"
            "<br><br>"
        )
        outro = "<br><br>Regards,<br><b>Angelos Keramaris<br>Accounts Payable Iberia</b>"

    html_message = intro + combined_html + outro
    st.markdown(html_message, unsafe_allow_html=True)

    if st.button("Send to GLPI"):
        if not ticket_id.isdigit():
            st.error("Invalid Ticket ID")
            st.stop()

        token, err = glpi_login()
        if err:
            st.error(err)
            st.stop()

        try:
            upd = glpi_update_ticket(token, ticket_id, 5, category_id)
            if upd.status_code >= 400:
                st.warning(f"Ticket update returned {upd.status_code}: {upd.text[:300]}")

            resp = glpi_add_solution(token, ticket_id, html_message)

            if resp.status_code == 400 or "already solved" in resp.text.lower():
                fu = glpi_add_followup(token, ticket_id, html_message)
                if fu.status_code >= 400:
                    st.error(f"Follow-up failed ({fu.status_code}): {fu.text[:300]}")
                else:
                    st.warning("Ticket already solved — posted as follow-up.")
            elif resp.status_code >= 400:
                st.error(f"Solution failed ({resp.status_code}): {resp.text[:300]}")
            else:
                st.success("Solution added to GLPI.")
        finally:
            glpi_kill_session(token)
