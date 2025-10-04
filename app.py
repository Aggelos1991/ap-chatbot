import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'vendor name for INV-1003', 'emails for open', 'due before 2025-10-11', 'open over 1000'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no": ["invoice number", "invoice", "inv no", "inv#", "document", "doc no"],
    "vendor_name": ["vendor", "supplier", "vendor name", "supplier name"],
    "vendor_email": ["email", "vendor email", "supplier email", "mail", "correo"],
    "status": ["status", "state", "payment status"],
    "amount": ["amount", "total", "invoice amount", "value", "importe"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "fecha vencimiento", "vencimiento"],
    "payment_date": ["payment date", "fecha pago", "paid date"],
    "po_number": ["po", "po number", "purchase order"]
}

def _clean_id(s: str) -> str:
    # normalize for id-like comparisons (invoice nos)
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def _clean_col(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    colmap = {}
    for c in df.columns:
        c_clean = _clean_col(c)
        mapped = None
        for std, alts in SYNONYMS.items():
            for alt in [std] + alts:
                if _clean_col(alt) == c_clean:
                    mapped = std
                    break
            if mapped:
                break
        colmap[c] = mapped if mapped else c
    df = df.rename(columns=colmap)
    for std in SYNONYMS.keys():
        if std not in df.columns:
            df[std] = None
    return df

# ---------------------------
# Utilities
# ---------------------------
def fmt_money(a, c):
    try:
        return f"{float(str(a).replace(',','')):,.2f} {c or 'EUR'}"
    except:
        return str(a)

def parse_date(s):
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except:
        return pd.NaT

def extract_date_from_text(q):
    q = q.lower()
    # between X and Y
    m_between = re.search(r"between\s+([\d\-\/]+)\s+and\s+([\d\-\/]+)", q)
    if m_between:
        d1 = parse_date(m_between.group(1))
        d2 = parse_date(m_between.group(2))
        return {"mode": "between", "d1": d1, "d2": d2}

    # single date
    m = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})", q)
    if not m:
        return None
    d = parse_date(m.group(1))
    if "before" in q or "smaller" in q or "<" in q or "earlier" in q:
        return {"mode": "before", "d1": d}
    if "after" in q or "greater" in q or ">" in q or "later" in q:
        return {"mode": "after", "d1": d}
    if "on" in q:
        return {"mode": "on", "d1": d}
    # default: treat as "on" if no keyword
    return {"mode": "on", "d1": d}

def find_invoices_in_query(df: pd.DataFrame, q: str):
    """
    Robust invoice match:
    - Normalize the whole query and each invoice_no, then check containment.
    - Works for INV1003, INV-1003, inv 1003, etc.
    """
    q_norm = _clean_id(q)
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    temp = df.copy()
    temp["__inv_norm__"] = temp["invoice_no"].astype(str).map(_clean_id)
    mask = temp["__inv_norm__"].apply(lambda inv: inv != "" and inv in q_norm)
    hits = temp[mask].drop(columns="__inv_norm__", errors="ignore")
    return hits

# ---------------------------
# Core query
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower().strip()

    # Prep numeric/date/status
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")
    working["status"] = working["status"].astype(str).str.lower()

    # ---------- Specific invoice first (so we don't fall back to 'all vendors') ----------
    invoice_hits = find_invoices_in_query(working, ql)
    if not invoice_hits.empty:
        r = invoice_hits.iloc[0]
        if "vendor" in ql and "email" not in ql:
            return f"Vendor name for **{r['invoice_no']}**: **{r.get('vendor_name','-')}**", None
        if "email" in ql:
            return f"Email for **{r['invoice_no']}**: **{r.get('vendor_email','-')}**", None
        if "amount" in ql:
            return f"Amount for **{r['invoice_no']}**: **{fmt_money(r.get('amount'), r.get('currency'))}**", None
        if "due" in ql:
            return f"Due date for **{r['invoice_no']}**: **{r.get('due_date','-')}**", None
        # generic single-invoice summary
        return (
            f"Invoice **{r.get('invoice_no','-')}** — vendor **{r.get('vendor_name','-')}**, "
            f"status **{r.get('status','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
            f"due **{r.get('due_date','-')}**.",
            invoice_hits.reset_index(drop=True),
        )

    # ---------- Broad filters ----------
    # Status
    if any(w in ql for w in ["open", "unpaid", "pending"]):
        working = working[working["status"].str.contains("open|unpaid|pending", case=False, na=False)]
    elif "paid" in ql and not any(w in ql for w in ["unpaid", "not paid", "open", "pending"]):
        working = working[working["status"].str.contains("paid", case=False, na=False)]

    # Amounts
    m_over = re.search(r"(over|above|greater than|>=)\s*([0-9][0-9,\.]*)", ql)
    if m_over:
        val = float(m_over.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    m_under = re.search(r"(under|below|less than|<=)\s*([0-9][0-9,\.]*)", ql)
    if m_under:
        val2 = float(m_under.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Due date filters
    di = extract_date_from_text(ql)
    if di and pd.notna(working["due_date_parsed"]).any():
        mode = di["mode"]
        d1 = di.get("d1")
        d2 = di.get("d2")
        if mode == "before" and pd.notna(d1):
            working = working[working["due_date_parsed"] <= d1]
        elif mode == "after" and pd.notna(d1):
            working = working[working["due_date_parsed"] >= d1]
        elif mode == "on" and pd.notna(d1):
            working = working[working["due_date_parsed"].dt.date == d1.date()]
        elif mode == "between" and pd.notna(d1) and pd.notna(d2):
            working = working[(working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)]

    # Vendor-only (distinct list) — but only when no specific invoice matched
    if "vendor" in ql and "email" not in ql:
        names = (
            working["vendor_name"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
        if not names:
            return "No vendor names match your filters.", None
        names_sorted = sorted(names, key=str.lower)
        return f"🏢 Vendors: {', '.join(names_sorted)}", None

    # Emails (distinct)
    if "email" in ql or "emails" in ql:
        emails = (
            working["vendor_email"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
        emails = sorted(emails, key=str.lower)
        if not emails:
            return "No vendor emails found for this query.", None
        return f"📧 Emails: {'; '.join(emails)}", None

    # Totals
    if "sum" in ql or "total" in ql or "amount" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount for matching invoices: **{total:,.2f}**", working.reset_index(drop=True)

    # Due date list view
    if "due" in ql:
        subset = working[["invoice_no", "vendor_name", "due_date"]].dropna()
        if subset.empty:
            return "No due dates match your filters.", None
        return "📅 Due dates for matching invoices:", subset.reset_index(drop=True)

    if working.empty:
        return "No invoices match your filters.", None

    return f"Found {len(working)} invoice(s) matching your query.", working.reset_index(drop=True)

# ---------------------------
# Streamlit UI
# ---------------------------
st.sidebar.header("📦 Upload Excel")
uploaded = st.sidebar.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success("✅ Excel loaded and columns normalized.")
        st.dataframe(df.head(), use_container_width=True)
    except Exception as e:
        st.error(f"Failed to load file: {e}")

st.subheader("Chat")
if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask about invoices...")

if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    ans, res = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", ans))
    st.chat_message("assistant").write(ans)
    if isinstance(res, pd.DataFrame) and not res.empty:
        st.dataframe(res, use_container_width=True)
