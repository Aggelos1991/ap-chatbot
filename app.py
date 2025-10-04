import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'show open over 560', 'emails for unpaid invoices', 'due before 2025-10-11', 'vendor names for open invoices'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no": ["invoice number", "invoice", "inv no", "inv#", "document"],
    "vendor_name": ["vendor", "supplier", "vendor name"],
    "vendor_email": ["email", "vendor email", "supplier email", "mail"],
    "status": ["status", "state", "payment status"],
    "amount": ["amount", "total", "invoice amount", "value"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "fecha vencimiento", "vencimiento"],
    "payment_date": ["payment date", "fecha pago"],
    "po_number": ["po", "po number"]
}

def normalize_columns(df):
    colmap = {}
    for c in df.columns:
        c_clean = re.sub(r"[^a-z0-9]+", " ", c.lower()).strip()
        mapped = None
        for std, alts in SYNONYMS.items():
            for alt in [std] + alts:
                if c_clean == re.sub(r"[^a-z0-9]+", " ", alt.lower()).strip():
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
    between = re.findall(r"between\s+([\d\-\/]+)\s+and\s+([\d\-\/]+)", q)
    if between:
        d1, d2 = [parse_date(x) for x in between[0]]
        return {"mode": "between", "d1": d1, "d2": d2}
    m = re.search(r"(\d{4}-\d{2}-\d{2})|(\d{2}/\d{2}/\d{4})", q)
    if not m:
        return None
    d = parse_date(m.group(0))
    if "before" in q or "smaller" in q or "<" in q:
        return {"mode": "before", "d1": d}
    if "after" in q or "greater" in q or ">" in q:
        return {"mode": "after", "d1": d}
    if "between" in q:
        return {"mode": "between", "d1": d, "d2": None}
    if "on" in q:
        return {"mode": "on", "d1": d}
    return None

def detect_invoice_ids(q):
    return re.findall(r"\b[a-z]{2,}[0-9]+\b", q.lower())

# ---------------------------
# Query Core
# ---------------------------
def run_query(q, df):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
    df["status"] = df["status"].astype(str).str.lower()

    working = df.copy()

    # Filter by status
    if any(w in ql for w in ["open", "unpaid", "pending"]):
        working = working[working["status"].str.contains("open|unpaid|pending", case=False, na=False)]
    elif "paid" in ql and not any(w in ql for w in ["unpaid", "not paid"]):
        working = working[working["status"].str.contains("paid", case=False, na=False)]

    # Filter by amount
    over = re.search(r"(over|above|greater than)\s*([0-9][0-9,\.]*)", ql)
    under = re.search(r"(under|below|less than)\s*([0-9][0-9,\.]*)", ql)
    if over:
        val = float(over.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    if under:
        val = float(under.group(2).replace(",", ""))
        working = working[working["amount"] <= val]

    # Filter by due date
    di = extract_date_from_text(ql)
    if di:
        mode = di["mode"]
        d1, d2 = di.get("d1"), di.get("d2")
        if mode == "before":
            working = working[working["due_date_parsed"] <= d1]
        elif mode == "after":
            working = working[working["due_date_parsed"] >= d1]
        elif mode == "between" and d1 is not None and d2 is not None:
            working = working[
                (working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)
            ]
        elif mode == "on":
            working = working[working["due_date_parsed"].dt.date == d1.date()]

    # Filter by invoice id
    invs = detect_invoice_ids(ql)
    if invs:
        matched = working[working["invoice_no"].astype(str).str.lower().isin(invs)]
        if not matched.empty:
            r = matched.iloc[0]
            if "vendor" in ql:
                return f"Vendor name for **{r['invoice_no']}**: **{r['vendor_name']}**", matched
            elif "email" in ql:
                return f"Email for **{r['invoice_no']}**: **{r['vendor_email']}**", matched
            elif "amount" in ql:
                return f"Amount for **{r['invoice_no']}**: **{fmt_money(r['amount'], r['currency'])}**", matched
            elif "due" in ql:
                return f"Due date for **{r['invoice_no']}**: **{r['due_date']}**", matched
        return f"No matching invoice for {invs}", None

    # Return filtered fields
    if "vendor" in ql and not "email" in ql:
        vendors = sorted(set(working["vendor_name"].dropna().astype(str)))
        if not vendors:
            return "No vendor names found for your query.", None
        return f"🏢 Vendors: {', '.join(vendors)}", working

    if "email" in ql:
        emails = sorted(set(working["vendor_email"].dropna().astype(str)))
        if not emails:
            return "No vendor emails found for this query.", None
        return f"📧 Emails: {'; '.join(emails)}", working

    if "amount" in ql:
        total = working["amount"].sum()
        return f"💰 Total amount for matching invoices: **{total:,.2f} EUR**", working

    if "due" in ql:
        result = working[["invoice_no", "vendor_name", "due_date"]].dropna()
        return f"📅 Due dates for matching invoices:", result

    if working.empty:
        return "No invoices match your filters.", None

    return f"Found {len(working)} invoice(s) matching your query.", working

# ---------------------------
# Streamlit UI
# ---------------------------
st.sidebar.header("📦 Upload Excel")
uploaded = st.sidebar.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        df = pd.read_excel(uploaded)
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
    if res is not None and not res.empty:
        st.dataframe(res, use_container_width=True)
