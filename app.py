import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'emails for invoices due before 2024-10-01', 'vendor names for unpaid invoices', 'amounts due after 2025-01-01'")

# -----------------------------------
# COLUMN NORMALIZATION
# -----------------------------------
SYNONYMS = {
    "invoice_no": ["invoice number", "invoice", "inv", "doc no"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name"],
    "vendor_email": ["email", "vendor email", "supplier email", "mail"],
    "status": ["status", "state", "payment status"],
    "amount": ["amount", "total", "invoice amount", "value"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "fecha vencimiento"],
    "payment_date": ["payment date", "fecha pago"],
    "po_number": ["po", "po number", "purchase order"]
}

def normalize_columns(df):
    colmap = {}
    for c in df.columns:
        c_clean = re.sub(r"[^a-z0-9]+", " ", c.lower().strip())
        for std, alts in SYNONYMS.items():
            if c_clean == std or any(c_clean == a.lower() for a in alts):
                colmap[c] = std
                break
    df = df.rename(columns=colmap)
    for std in SYNONYMS.keys():
        if std not in df.columns:
            df[std] = None
    return df

# -----------------------------------
# PARSING UTILITIES
# -----------------------------------
def extract_date_from_query(q):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", q)
    if not m:
        return None
    try:
        return pd.to_datetime(m.group(1))
    except:
        return None

def fmt_money(amount, currency="EUR"):
    try:
        a = float(str(amount).replace(",", ""))
        return f"{a:,.2f} {currency}"
    except:
        return str(amount)

# -----------------------------------
# CORE LOGIC
# -----------------------------------
def run_query(q, df):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower()
    df = df.copy()

    # Normalize
    df["status"] = df["status"].astype(str).str.lower().fillna("")
    df["vendor_name"] = df["vendor_name"].astype(str)
    df["vendor_email"] = df["vendor_email"].astype(str)
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")

    # Filters
    user_date = extract_date_from_query(ql)
    if user_date is not None:
        if any(k in ql for k in ["before", "earlier", "<"]):
            df = df[df["due_date_parsed"] < user_date]
        elif any(k in ql for k in ["after", "later", "since", ">"]):
            df = df[df["due_date_parsed"] > user_date]

    if any(k in ql for k in ["open", "unpaid", "pending"]):
        df = df[df["status"].str.contains("open|unpaid|pending", na=False)]
    elif "paid" in ql and not "unpaid" in ql:
        df = df[df["status"].str.contains("paid", na=False)]

    # What user wants
    wants_email = "email" in ql or "emails" in ql
    wants_vendor = "vendor" in ql or "supplier" in ql
    wants_amount = "amount" in ql or "total" in ql or "value" in ql
    wants_due = "due" in ql
    wants_status = "status" in ql

    # Handle specific data
    if wants_email:
        emails = sorted(set(df["vendor_email"].dropna().astype(str).tolist()))
        if not emails:
            return "📭 No vendor emails found for this query.", None
        return f"📧 Vendor emails matching your criteria:\n\n" + "; ".join(emails), df

    if wants_vendor:
        vendors = sorted(set(df["vendor_name"].dropna().astype(str).tolist()))
        if not vendors:
            return "No vendor names found for this query.", None
        return f"🏢 Vendors matching your criteria:\n\n" + ", ".join(vendors), df

    if wants_amount:
        if df.empty:
            return "No invoices match your filters.", None
        out = "\n".join([f"{r['invoice_no']}: {fmt_money(r['amount'], r.get('currency', 'EUR'))}" for _, r in df.iterrows()])
        return f"💰 Amounts for matching invoices:\n\n{out}", df

    if wants_due and not user_date:
        out = "\n".join([f"{r['invoice_no']} → {r['due_date']}" for _, r in df.iterrows()])
        return f"📅 Due dates for invoices:\n\n{out}", df

    if wants_status:
        out = "\n".join([f"{r['invoice_no']} → {r['status'].capitalize()}" for _, r in df.iterrows()])
        return f"📊 Status overview:\n\n{out}", df

    if "total" in ql or "sum" in ql:
        total = df["amount"].sum()
        return f"💰 Total of filtered invoices: **{total:,.2f} EUR**", df

    if df.empty:
        return "No invoices found for your criteria.", None

    return f"Found {len(df)} invoice(s) matching your query.", df

# -----------------------------------
# STREAMLIT INTERFACE
# -----------------------------------
st.sidebar.header("📦 Upload Excel")
uploaded = st.sidebar.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success("✅ Excel loaded and columns normalized.")
        st.dataframe(df.head(50), use_container_width=True)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")

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
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)