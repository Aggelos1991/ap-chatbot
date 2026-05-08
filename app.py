# ==========================================================
# THE REMITATOR — FINAL HYBRID (HARDENED GLPI EDITION)
# DEFAULT USER ID = 22487 (ANGELOS KERAMARIS)
# ==========================================================

import os, re, io, requests
from itertools import combinations
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
# ENV / SECRETS
# ----------------------------------------------------------
load_dotenv()

def _get(name, default=""):
    """Streamlit Cloud secrets first, then environment variable, then default."""
    try:
        v = st.secrets.get(name)
        if v:
            return v
    except Exception:
        pass
    return os.getenv(name, default)

GLPI_URL   = (_get("GLPI_URL") or "").rstrip("/")
APP_TOKEN  = _get("APP_TOKEN")
USER_TOKEN = _get("USER_TOKEN")


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


def col_by_letter(df, letter):
    letter = (letter or "").strip().upper()
    if not letter or not letter.isalpha():
        return None
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    idx -= 1
    if 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None


def find_cn_combo(pool, used, target, max_combo=3):
    """
    pool: list of (idx, doc, abs_amount) tuples (pre-built once).
    used: set of CN indices already consumed.
    target: signed difference (we match abs).
    """
    target = round(abs(target), 2)
    if target <= 0:
        return None

    avail = [t for t in pool if t[0] not in used and 0 < t[2] <= target]
    if not avail:
        return None

    # Size 1
    for entry in avail:
        if entry[2] == target:
            return [entry]
    if max_combo < 2:
        return None

    # Size 2
    n = len(avail)
    for i in range(n):
        a = avail[i][2]
        for j in range(i + 1, n):
            if round(a + avail[j][2], 2) == target:
                return [avail[i], avail[j]]
    if max_combo < 3:
        return None

    # Size 3+
    for size in range(3, min(max_combo, len(avail)) + 1):
        for combo in combinations(avail, size):
            if round(sum(c[2] for c in combo), 2) == target:
                return list(combo)
    return None


def safe_json(resp):
    try:
        return resp.json()
    except Exception:
        return None


def build_excel_export(export_data):
    """One sheet per payment code + a combined Summary sheet."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Summary sheet
        summary_rows = []
        for code, info in export_data.items():
            rows = info["rows"]
            total_row = rows[rows["Alt. Document"] == "TOTAL"]
            total = float(total_row["Invoice Value"].iloc[0]) if not total_row.empty else 0.0
            summary_rows.append({"Payment Code": code, "Vendor": info["vendor"], "Total (€)": total})
        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

        # Per-payment-code sheets (Excel sheet names: max 31 chars, no \ / * ? : [ ])
        used = set()
        for code, info in export_data.items():
            base = re.sub(r'[\\/*?:\[\]]', '_', str(code))[:28] or "Sheet"
            name, n = base, 1
            while name in used:
                n += 1
                name = f"{base[:25]}_{n}"
            used.add(name)
            out = info["rows"].copy()
            out.insert(0, "Vendor", info["vendor"])
            out.insert(0, "Payment Code", code)
            out.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


# ----------------------------------------------------------
# GLPI FUNCTIONS — HARDENED
# ----------------------------------------------------------
def glpi_login():
    if not GLPI_URL or not APP_TOKEN or not USER_TOKEN:
        return None, (
            "Missing GLPI credentials. Check your Streamlit secrets / .env:\n"
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
        return None, f"Unexpected GLPI response (status {r.status_code}). Body: {r.text[:300]}"
    token = data.get("session_token")
    if not token:
        return None, f"GLPI response had no session_token. Body: {data}"
    return token, None


def glpi_update_ticket(token, ticket_id, status=5, category_id=None):
    payload = {"input": {"id": int(ticket_id), "status": int(status),
                         "users_id_lastupdater": DEFAULT_USER_ID,
                         "users_id_recipient": DEFAULT_USER_ID}}
    if category_id:
        payload["input"]["itilcategories_id"] = int(category_id)
    return requests.put(f"{GLPI_URL}/Ticket/{ticket_id}", json=payload,
                        headers={"Session-Token": token, "App-Token": APP_TOKEN}, timeout=20)


def glpi_add_solution(token, ticket_id, html):
    payload = {"input": {"itemtype": "Ticket", "items_id": int(ticket_id),
                         "users_id": DEFAULT_USER_ID, "users_id_recipient": DEFAULT_USER_ID,
                         "content": html, "solutiontypes_id": 10, "status": 5}}
    return requests.post(f"{GLPI_URL}/ITILSolution", json=payload,
                         headers={"Session-Token": token, "App-Token": APP_TOKEN}, timeout=20)


def glpi_add_followup(token, ticket_id, html):
    payload = {"input": {"itemtype": "Ticket", "items_id": int(ticket_id),
                         "users_id": DEFAULT_USER_ID, "users_id_recipient": DEFAULT_USER_ID,
                         "content": html}}
    return requests.post(f"{GLPI_URL}/Ticket/{ticket_id}/ITILFollowup", json=payload,
                         headers={"Session-Token": token, "App-Token": APP_TOKEN}, timeout=20)


def glpi_kill_session(token):
    try:
        requests.get(f"{GLPI_URL}/killSession",
                     headers={"Session-Token": token, "App-Token": APP_TOKEN}, timeout=10)
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
# COLUMN DETECTION (payment file)
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
# CREDIT NOTES — load and pre-build pool ONCE
# ----------------------------------------------------------
cn_df = None
cn_pool = None
cn_alt = cn_credit = cn_charge = cn_reason = None
if cn_file:
    cn_df = pd.read_excel(cn_file)
    cn_df.columns = [c.strip() for c in cn_df.columns]
    cn_df = cn_df.loc[:, ~cn_df.columns.duplicated()]

    cn_alt = find_col(cn_df, [
        "AltDocument", "Alt.Document", "Alt. Document",
        "Documento", "Nº Documento", "NumDocumento",
        "Factura", "Nº Factura", "Referencia", "Ref",
        "Document", "DocNumber", "InvoiceNumber"
    ])
    cn_credit = find_col(cn_df, ["Credit", "CreditAmount", "CreditValue", "Haber", "Abono"])
    cn_charge = find_col(cn_df, ["Charge", "ChargeAmount", "Cargo", "DEBE", "Debit"])
    cn_reason = find_col(cn_df, [
        "Reason", "Razon", "Razón", "Motivo",
        "Description", "Descripcion", "Descripción",
        "Concept", "Concepto"
    ])

    with st.expander("🔧 Credit Notes — column override",
                     expanded=(cn_alt is None or (cn_credit is None and cn_charge is None))):
        st.write("**Columns in CN file:**", list(cn_df.columns))
        st.write(
            f"Auto-detected — document: `{cn_alt}` | "
            f"credit: `{cn_credit}` | charge: `{cn_charge}` | reason: `{cn_reason}`"
        )

        c1, c2 = st.columns(2)
        with c1:
            letter_alt    = st.text_input("Document column (letter)", "")
            letter_credit = st.text_input("Credit column (letter)", "F")
        with c2:
            letter_charge = st.text_input("Charge column (letter)", "E")
            letter_reason = st.text_input("Reason column (letter)", "G")

        if letter_alt:
            r = col_by_letter(cn_df, letter_alt)
            if r: cn_alt = r
            else: st.error(f"Letter '{letter_alt}' out of range for document.")
        if letter_credit:
            r = col_by_letter(cn_df, letter_credit)
            if r: cn_credit = r
            else: st.error(f"Letter '{letter_credit}' out of range for credit.")
        if letter_charge:
            r = col_by_letter(cn_df, letter_charge)
            if r: cn_charge = r
            else: st.error(f"Letter '{letter_charge}' out of range for charge.")
        if letter_reason:
            r = col_by_letter(cn_df, letter_reason)
            if r: cn_reason = r
            else: st.error(f"Letter '{letter_reason}' out of range for reason.")

    if cn_alt and (cn_credit or cn_charge):
        if cn_credit: cn_df[cn_credit] = cn_df[cn_credit].apply(parse_amount)
        if cn_charge: cn_df[cn_charge] = cn_df[cn_charge].apply(parse_amount)

        CN_KEYWORDS = (
            "credit note", "credit", "nota credito", "nota de credito",
            "nota crédito", "nota de crédito", "abono", "ncr", "n/c", "cn"
        )

        cn_pool = []
        skipped_no_reason = 0
        for i in cn_df.index:
            if cn_reason:
                reason = str(cn_df.at[i, cn_reason] or "").lower()
                if not any(k in reason for k in CN_KEYWORDS):
                    skipped_no_reason += 1
                    continue

            val = 0.0
            if cn_credit:
                val = parse_amount(cn_df.at[i, cn_credit])
            if val == 0 and cn_charge:
                val = parse_amount(cn_df.at[i, cn_charge])
            if val == 0:
                continue

            doc = str(cn_df.at[i, cn_alt])
            cn_pool.append((int(i), doc, round(abs(val), 2)))

        msg = (
            f"CN matching enabled — {len(cn_pool)} CNs loaded "
            f"(document=`{cn_alt}`, credit=`{cn_credit}`, "
            f"charge=`{cn_charge}`, reason=`{cn_reason}`)."
        )
        if cn_reason and skipped_no_reason:
            msg += f" Skipped {skipped_no_reason} non-CN rows by Reason."
        st.success(msg)
    else:
        st.warning("CN matching disabled — set the document and at least one of credit/charge columns above.")
        cn_df = None

cn_used_global = set()
combined_html = ""
export_data = {}
debug_rows_all = []

# ----------------------------------------------------------
# PROCESS EACH PAYMENT CODE
# ----------------------------------------------------------
MAX_COMBO = 3  # combine up to 3 CNs to cover a single invoice diff

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
        inv_val = float(row[inv_col])
        pay_val = float(row[payv_col])
        diff    = round(pay_val - inv_val, 2)

        summary_rows.append({"Alt. Document": inv, "Invoice Value": inv_val})

        dbg = {
            "Payment Code": str(pay_code),
            "Vendor": str(vendor),
            "Alt. Document": inv,
            "Invoice Value": inv_val,
            "Payment Value": pay_val,
            "Difference": diff,
            "Matched CN(s)": "",
            "CN Value(s)": "",
            "CN Count": 0,
            "Status": "",
        }

        if abs(diff) < 0.01:
            dbg["Status"] = "✓ Exact"
            debug_rows_all.append(dbg)
            continue

        combo = None
        if cn_pool is not None:
            combo = find_cn_combo(cn_pool, cn_used_global, diff, max_combo=MAX_COMBO)

        if combo:
            sign = -1 if diff < 0 else 1
            docs, vals = [], []
            for idx, doc, amt in combo:
                signed = sign * amt
                cn_rows.append({"Alt. Document": f"{doc} (CN)", "Invoice Value": signed})
                cn_used_global.add(idx)
                docs.append(doc)
                vals.append(f"{signed:.2f}")
            dbg["Status"]        = "✓ CN matched" if len(combo) == 1 else f"✓ {len(combo)} CNs combined"
            dbg["Matched CN(s)"] = ", ".join(docs)
            dbg["CN Value(s)"]   = ", ".join(vals)
            dbg["CN Count"]      = len(combo)
        else:
            unmatched.append({"Alt. Document": f"{inv} (Adj. Diff)", "Invoice Value": diff})
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

with tab1:
    st.markdown(combined_html, unsafe_allow_html=True)
    if export_data:
        st.download_button(
            "⬇️ Download Payment Analysis (Excel)",
            data=build_excel_export(export_data),
            file_name="payment_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab2:
    if debug_rows_all:
        dbg_df = pd.DataFrame(debug_rows_all)
        for c in ["Payment Code", "Vendor", "Alt. Document",
                  "Matched CN(s)", "CN Value(s)", "Status"]:
            if c in dbg_df.columns:
                dbg_df[c] = dbg_df[c].astype(str)
        st.dataframe(dbg_df, width="stretch")
        st.download_button(
            "⬇️ Download Debug CSV",
            dbg_df.to_csv(index=False).encode("utf-8"),
            file_name="debug_breakdown.csv",
            mime="text/csv",
        )
    else:
        st.info("No rows processed — check that the Payment Document Codes match the Excel.")

with tab3:
    language = st.radio("Language", ["Spanish", "English"], horizontal=True)

    ticket_id   = st.text_input("Ticket ID")
    category_id = st.text_input("Category ID")

    if language == "Spanish":
        intro = ("Estimado proveedor,<br><br>"
                 "Adjuntamos el detalle de facturas correspondientes a los pagos realizados:"
                 "<br><br>")
        outro = SIGNATURE
    else:
        intro = ("Dear supplier,<br><br>"
                 "Please find below the invoice breakdown corresponding to the executed payments:"
                 "<br><br>")
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
