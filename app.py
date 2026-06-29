# ==========================================================
# THE REMITATOR — FINAL HYBRID (HARDENED GLPI EDITION)
# DEFAULT USER ID = 22487 (ANGELOS KERAMARIS)
# + EDITABLE TABLES (st.data_editor)
# + BULK TICKETS: post the same answer to many tickets at once
# + BULK EMAIL mode (write your own message, no Excel needed)
# + AUTO-ASSIGN every ticket to Angelos (GLPI 10 _actors + Ticket_User)
# + Sidebar GLPI connection tester
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

# .strip() removes stray spaces/newlines that sneak into the Streamlit secret and break login
GLPI_URL   = (_get("GLPI_URL") or "").strip().rstrip("/")
APP_TOKEN  = (_get("APP_TOKEN") or "").strip()
USER_TOKEN = (_get("USER_TOKEN") or "").strip()


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


def fmt_date(v):
    """Format a payment date from the Excel into dd/mm/yyyy (European). Falls back to raw text."""
    if pd.isna(v):
        return ""
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%d/%m/%Y")
        except Exception:
            pass
    try:
        ts = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.notna(ts):
            return ts.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(v).strip()


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
    target = round(abs(target), 2)
    if target <= 0:
        return None

    avail = [t for t in pool if t[0] not in used and 0 < t[2] <= target]
    if not avail:
        return None

    for entry in avail:
        if entry[2] == target:
            return [entry]
    if max_combo < 2:
        return None

    n = len(avail)
    for i in range(n):
        a = avail[i][2]
        for j in range(i + 1, n):
            if round(a + avail[j][2], 2) == target:
                return [avail[i], avail[j]]
    if max_combo < 3:
        return None

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


def _body_no_total(rows):
    """Return rows without any TOTAL line, with clean numeric Invoice Value."""
    body = rows[rows["Alt. Document"].astype(str) != "TOTAL"].copy()
    body["Alt. Document"] = body["Alt. Document"].fillna("").astype(str)
    body["Invoice Value"] = pd.to_numeric(body["Invoice Value"], errors="coerce").fillna(0.0)
    return body.reset_index(drop=True)


def build_combined_html(export_data):
    """Rebuild the HTML summary from (possibly edited) export_data. Totals recomputed."""
    html = ""
    for code, info in export_data.items():
        body = _body_no_total(info["rows"])
        total_value = body["Invoice Value"].sum()

        full = body.copy()
        full.loc[len(full)] = ["TOTAL", total_value]

        disp = full.copy()
        disp["Invoice Value (€)"] = disp["Invoice Value"].apply(lambda v: f"€{v:,.2f}")
        disp = disp[["Alt. Document", "Invoice Value (€)"]]

        pay_date_line = f"<b>Payment Date:</b> {info.get('pay_date','')}<br>" if info.get("pay_date") else ""
        html += f"""
<b>Payment Code:</b> {code}<br>
<b>Vendor:</b> {info['vendor']}<br>
{pay_date_line}<b>Total Amount:</b> €{total_value:,.2f}<br><br>
{disp.to_html(index=False, border=0)}
<br><hr><br>
"""
    if html.endswith("<br><hr><br>"):
        html = html[:-12]
    return html


def build_excel_export(export_data):
    """One sheet per payment code + a combined Summary sheet. Totals recomputed."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_rows = []
        for code, info in export_data.items():
            body = _body_no_total(info["rows"])
            total = body["Invoice Value"].sum()
            summary_rows.append({
                "Payment Code": code,
                "Payment Date": info.get("pay_date", ""),
                "Vendor": info["vendor"],
                "Total (€)": total,
            })
        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

        used = set()
        for code, info in export_data.items():
            base = re.sub(r'[\\/*?:\[\]]', '_', str(code))[:28] or "Sheet"
            name, n = base, 1
            while name in used:
                n += 1
                name = f"{base[:25]}_{n}"
            used.add(name)

            body = _body_no_total(info["rows"])
            total = body["Invoice Value"].sum()
            out = body.copy()
            out.loc[len(out)] = ["TOTAL", total]
            out.insert(0, "Vendor", info["vendor"])
            out.insert(0, "Payment Date", info.get("pay_date", ""))
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


def glpi_assign_ticket(token, ticket_id, user_id=DEFAULT_USER_ID):
    """Set the ticket's 'Assigned to' technician. Tries both:
      1) GLPI 10 `_actors` structure (PUT /Ticket) — what the web form uses.
      2) Universal Ticket_User type=2 (POST /Ticket_User) — works on 9.x & 10.x.
    Returns True if either method was accepted (or the actor already exists)."""
    hdr = {"Session-Token": token, "App-Token": APP_TOKEN}
    ok = False

    # Method 1 — GLPI 10 _actors (only touches the 'assign' role; requester/observer untouched)
    try:
        r1 = requests.put(
            f"{GLPI_URL}/Ticket/{ticket_id}",
            json={"input": {"id": int(ticket_id),
                            "_actors": {"assign": [
                                {"itemtype": "User", "items_id": int(user_id), "use_notification": 1}
                            ]}}},
            headers=hdr, timeout=20,
        )
        if r1.status_code < 400:
            ok = True
    except Exception:
        pass

    # Method 2 — Ticket_User type=2 (additive). 'already exists' (400) still means assigned.
    try:
        r2 = requests.post(
            f"{GLPI_URL}/Ticket_User",
            json={"input": {"tickets_id": int(ticket_id), "users_id": int(user_id), "type": 2}},
            headers=hdr, timeout=20,
        )
        if r2.status_code < 400 or "already" in (r2.text or "").lower():
            ok = True
    except Exception:
        pass

    return ok


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


def glpi_send_one(token, ticket_id, html_message, category_id):
    """Update one ticket + assign to Angelos + post the solution
    (falls back to follow-up if already solved). Returns a short result string."""
    try:
        upd = glpi_update_ticket(token, ticket_id, 5, category_id)
        upd_warn = "" if upd.status_code < 400 else f" (update {upd.status_code})"

        # always assign the ticket to Angelos — robust across GLPI 9.x / 10.x
        assigned = glpi_assign_ticket(token, ticket_id, DEFAULT_USER_ID)
        assign_note = "" if assigned else " (assign?)"

        resp = glpi_add_solution(token, ticket_id, html_message)
        if resp.status_code == 400 or "already solved" in (resp.text or "").lower():
            fu = glpi_add_followup(token, ticket_id, html_message)
            if fu.status_code >= 400:
                return f"❌ Follow-up failed ({fu.status_code})" + upd_warn + assign_note
            return "⚠️ Already solved — follow-up posted" + upd_warn + assign_note
        if resp.status_code >= 400:
            return f"❌ Solution failed ({resp.status_code})" + upd_warn + assign_note
        return "✅ Solution added" + upd_warn + assign_note
    except Exception as e:
        return f"❌ Error: {e}"


# ----------------------------------------------------------
# SIDEBAR — GLPI connection check / debugger
# ----------------------------------------------------------
with st.sidebar:
    st.header("🔑 GLPI connection")

    def _mask(t):
        if not t:
            return "❌ NOT SET"
        return f"✅ set · {len(t)} chars · ends …{t[-4:]}"

    st.write(f"**GLPI_URL:** {GLPI_URL or '❌ NOT SET'}")
    st.write(f"**APP_TOKEN:** {_mask(APP_TOKEN)}")
    st.write(f"**USER_TOKEN:** {_mask(USER_TOKEN)}")

    if st.button("Test GLPI login"):
        tok, err = glpi_login()
        if err:
            st.error(err)
            st.info(
                "If it says *user_token seems invalid*:\n"
                "1. GLPI → your avatar → **My settings → Remote access keys**.\n"
                "2. **Regenerate** the *API token* and copy it.\n"
                "3. Streamlit Cloud → **Manage app → Settings → Secrets** → set "
                "`USER_TOKEN = \"...\"` (no trailing spaces).\n"
                "4. Save → app reboots → press Test again."
            )
        else:
            st.success("✅ Login OK — tokens are valid.")
            glpi_kill_session(tok)

    st.caption("Tokens come from Streamlit **Secrets** / .env — never stored in the code.")


# ----------------------------------------------------------
# MODE SELECTOR
# ----------------------------------------------------------
app_mode = st.radio(
    "What do you want to do?",
    ["✉️ Bulk Email to Tickets", "💶 Payment Analysis"],
    horizontal=True,
)

# ==========================================================
# MODE 1: BULK EMAIL — write your own message, post to many tickets
#         (no Excel required — this is the standalone bulk sender)
# ==========================================================
if app_mode.startswith("✉️"):
    st.subheader("✉️ Bulk Email to Tickets")
    st.caption("Write your own message once and post it to as many GLPI tickets as you like. No Excel needed.")

    lang = st.radio("Language (for greeting & signature)", ["Spanish", "English"], horizontal=True)

    c1, c2 = st.columns(2)
    with c1:
        use_greeting = st.checkbox("Add greeting line", value=True)
    with c2:
        use_sig = st.checkbox("Add my signature", value=True)

    greeting = ""
    if use_greeting:
        greeting = "Estimado proveedor,<br><br>" if lang == "Spanish" else "Dear supplier,<br><br>"

    if lang == "English":
        sig = "<br><br>Regards,<br><b>Angelos Keramaris<br>Accounts Payable Iberia</b>"
    else:
        sig = SIGNATURE
    signature = sig if use_sig else ""

    body = st.text_area(
        "Your message  —  HTML supported (use <br> for a new line, <b>…</b> for bold)",
        height=300,
        placeholder="Write your email here…",
    )

    html_message = greeting + (body or "") + signature

    st.markdown("**Preview — this exact message goes to every ticket:**")
    st.markdown(html_message or "_(empty — type your message above)_", unsafe_allow_html=True)

    st.markdown("---")
    ticket_input = st.text_area(
        "Ticket IDs  —  paste one or many (comma, space, or new-line separated)",
        height=90,
        placeholder="100245, 100246, 100247",
    )
    category_id = st.text_input("Category ID (optional — applied to every ticket)")

    ticket_ids = list(dict.fromkeys(re.findall(r"\d+", ticket_input or "")))
    has_body = bool((body or "").strip())

    if ticket_ids:
        st.info(f"Ready to post to **{len(ticket_ids)}** ticket(s): {', '.join(ticket_ids)}")
    else:
        st.caption("Enter at least one Ticket ID to enable sending.")
    if not has_body:
        st.caption("Type your message above to enable sending.")

    confirm = st.checkbox(
        f"I confirm posting this message to {len(ticket_ids)} ticket(s).",
        value=False,
        disabled=not (ticket_ids and has_body),
    )

    if st.button("🚀 Send to GLPI", disabled=not (ticket_ids and has_body and confirm)):
        token, err = glpi_login()
        if err:
            st.error(err)
            st.stop()

        results = []
        total = len(ticket_ids)
        progress = st.progress(0.0)
        status_box = st.empty()
        try:
            for i, tid in enumerate(ticket_ids):
                status_box.write(f"Processing ticket {tid}  ({i + 1}/{total}) …")
                res = glpi_send_one(token, tid, html_message, category_id)
                results.append({"Ticket": tid, "Result": res})
                progress.progress((i + 1) / total)
        finally:
            glpi_kill_session(token)

        status_box.empty()
        res_df = pd.DataFrame(results)
        st.dataframe(res_df, use_container_width=True)

        ok   = sum(1 for r in results if r["Result"].startswith("✅"))
        warn = sum(1 for r in results if r["Result"].startswith("⚠️"))
        bad  = sum(1 for r in results if r["Result"].startswith("❌"))
        line = f"Done — ✅ {ok} solved · ⚠️ {warn} follow-up · ❌ {bad} failed  (of {total})."
        (st.success if bad == 0 else st.warning)(line)

        st.download_button(
            "⬇️ Download results CSV",
            res_df.to_csv(index=False).encode("utf-8"),
            file_name="glpi_bulk_results.csv",
            mime="text/csv",
        )

    st.stop()  # bulk-email mode ends here — payment flow below does not run

# ==========================================================
# MODE 2: PAYMENT ANALYSIS  (original flow)
# ==========================================================
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

alt_col     = find_col(df, ["Alt.Document", "AltDocument", "Alt. Document"]) or "Alt. Document"
inv_col     = find_col(df, ["InvoiceValue", "Invoice Value"]) or "Invoice Value"
payv_col    = find_col(df, ["PaymentValue", "Payment Value"]) or "Payment Value"
vendor_col  = find_col(df, ["Vendor", "SupplierName", "Supplier"])
paydate_col = find_col(df, [
    "PaymentDate", "Payment Date", "PaymentDt",
    "FechaPago", "Fecha de Pago", "Fecha Pago",
    "FechaValor", "Fecha Valor", "ValueDate", "Value Date",
    "Fecha", "Date",
])

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
export_data = {}
debug_rows_all = []

# ----------------------------------------------------------
# PROCESS EACH PAYMENT CODE  (builds base rows, NO total yet)
# ----------------------------------------------------------
MAX_COMBO = 3  # combine up to 3 CNs to cover a single invoice diff

for pay_code in selected_codes:
    subset = df[df[pay_doc_col].astype(str) == str(pay_code)].copy()
    if subset.empty:
        continue

    subset[inv_col]  = subset[inv_col].apply(parse_amount)
    subset[payv_col] = subset[payv_col].apply(parse_amount)

    vendor = subset[vendor_col].iloc[0] if vendor_col else "Unknown Vendor"
    pay_date = fmt_date(subset[paydate_col].iloc[0]) if paydate_col else ""

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
            "Payment Date": pay_date,
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

    body = pd.concat(
        [pd.DataFrame(summary_rows), pd.DataFrame(cn_rows), pd.DataFrame(unmatched)],
        ignore_index=True,
    )
    # store base rows WITHOUT the TOTAL line — totals are recomputed on render/export
    export_data[pay_code] = {"vendor": vendor, "pay_date": pay_date, "rows": body.copy()}

# ----------------------------------------------------------
# TABS
# ----------------------------------------------------------
combined_html = ""
tab1, tab2, tab3 = st.tabs(["Summary", "Advanced Debug", "GLPI"])

with tab1:
    st.caption("✏️ This table is editable. Double-click a cell to change it · "
               "➕ row at the bottom to add · select a row's checkbox + press ⌫ to delete · "
               "totals recalc automatically and edits flow into the Excel + GLPI message.")

    # --- THE EDITABLE TABLE (always on, this IS the summary) ---
    for code in list(export_data.keys()):
        info = export_data[code]
        label = f"**{code} — {info['vendor']}**"
        if info.get("pay_date"):
            label += f"  ·  {info['pay_date']}"
        st.markdown(label)

        base = _body_no_total(info["rows"])
        edited = st.data_editor(
            base,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{code}",
            column_config={
                "Alt. Document": st.column_config.TextColumn("Alt. Document", width="large"),
                "Invoice Value": st.column_config.NumberColumn(
                    "Invoice Value", format="%.2f", step=0.01
                ),
            },
        )
        # clean + push edits back into export_data so HTML/Excel/GLPI all use them
        edited = edited.dropna(how="all")
        edited["Alt. Document"] = edited["Alt. Document"].fillna("").astype(str)
        edited["Invoice Value"] = pd.to_numeric(edited["Invoice Value"], errors="coerce").fillna(0.0)
        export_data[code]["rows"] = edited.reset_index(drop=True)

        st.markdown(f"**Total: €{edited['Invoice Value'].sum():,.2f}**")
        st.markdown("---")

    # rebuild the GLPI/email HTML from the (edited) data
    combined_html = build_combined_html(export_data)

    with st.expander("👁️ Preview formatted message (what gets sent to GLPI)"):
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
        for c in ["Payment Code", "Payment Date", "Vendor", "Alt. Document",
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
        st.caption("Note: debug reflects the automatic match pass, not your manual edits.")
    else:
        st.info("No rows processed — check that the Payment Document Codes match the Excel.")

with tab3:
    language = st.radio("Language", ["Spanish", "English"], horizontal=True)

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

    # ---- what to send: the payment analysis, or any custom text ----
    msg_mode = st.radio(
        "Message to post",
        ["Payment analysis (from Summary tab)", "Custom message (free text)"],
        horizontal=True,
    )
    if msg_mode.startswith("Custom"):
        custom_msg = st.text_area(
            "Your message — HTML supported (use <br> for line breaks)",
            height=220,
            placeholder="Estimado proveedor, ...",
        )
        append_sig = st.checkbox("Append signature", value=True)
        html_message = (custom_msg or "") + (SIGNATURE if append_sig else "")
    else:
        # uses the edited combined_html from tab1
        html_message = intro + combined_html + outro

    # ---- BULK ticket IDs ----
    ticket_input = st.text_area(
        "Ticket IDs — paste one or many (comma, space, or new-line separated)",
        height=90,
        placeholder="100245, 100246, 100247",
    )
    category_id = st.text_input("Category ID (optional — applied to every ticket)")

    # robust: pull every number out of whatever the user pasted, keep order, dedupe
    ticket_ids = list(dict.fromkeys(re.findall(r"\d+", ticket_input or "")))

    st.markdown("**Preview — this exact message goes to every ticket below:**")
    st.markdown(html_message, unsafe_allow_html=True)

    if ticket_ids:
        st.info(f"Ready to post to **{len(ticket_ids)}** ticket(s): {', '.join(ticket_ids)}")
    else:
        st.caption("Enter at least one Ticket ID to enable sending.")

    confirm = st.checkbox(
        f"I confirm posting this message to {len(ticket_ids)} ticket(s).",
        value=False,
        disabled=not ticket_ids,
    )

    if st.button("🚀 Send to GLPI", disabled=not (ticket_ids and confirm)):
        token, err = glpi_login()
        if err:
            st.error(err)
            st.stop()

        results = []
        total = len(ticket_ids)
        progress = st.progress(0.0)
        status_box = st.empty()
        try:
            for i, tid in enumerate(ticket_ids):
                status_box.write(f"Processing ticket {tid}  ({i + 1}/{total}) …")
                res = glpi_send_one(token, tid, html_message, category_id)
                results.append({"Ticket": tid, "Result": res})
                progress.progress((i + 1) / total)
        finally:
            glpi_kill_session(token)

        status_box.empty()
        res_df = pd.DataFrame(results)
        st.dataframe(res_df, use_container_width=True)

        ok   = sum(1 for r in results if r["Result"].startswith("✅"))
        warn = sum(1 for r in results if r["Result"].startswith("⚠️"))
        bad  = sum(1 for r in results if r["Result"].startswith("❌"))
        if bad == 0:
            st.success(f"Done — ✅ {ok} solved · ⚠️ {warn} follow-up · ❌ {bad} failed  (of {total}).")
        else:
            st.warning(f"Done — ✅ {ok} solved · ⚠️ {warn} follow-up · ❌ {bad} failed  (of {total}). See table above.")

        st.download_button(
            "⬇️ Download results CSV",
            res_df.to_csv(index=False).encode("utf-8"),
            file_name="glpi_bulk_results.csv",
            mime="text/csv",
        )
