import re
import pandas as pd
import streamlit as st
from datetime import datetime

# ----------------------------------------------------------
# Streamlit setup
# ----------------------------------------------------------
st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption(
    "Try: 'open amount for vendor test', 'emails for paid invoices', "
    "'due date invoices < today', 'vendor Technogym Iberia summary'"
)

# ----------------------------------------------------------
# Column normalization
# ----------------------------------------------------------
SYNONYMS = {
    "alternative_document": ["alternative document", "alt doc", "alt document", "document", "invoice"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name"],
    "vendor_email": ["email", "vendor email", "supplier email", "correo", "mail"],
    "amount": ["amount", "total", "invoice amount", "importe", "value"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "vencimiento", "fecha vencimiento"],
    "payment_date": ["payment date", "fecha pago", "paid date"],
    "agreed": ["agreed", "is agreed", "approved", "paid flag"],
}
STANDARD_COLS = list(SYNONYMS.keys())

def _clean(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    colmap = {}
    for c in df.columns:
        c_clean = _clean(c)
        mapped = None
        for std, alts in SYNONYMS.items():
            for alt in [std] + alts:
                if _clean(alt) == c_clean:
                    mapped = std
                    break
            if mapped:
                break
        colmap[c] = mapped if mapped else c
    df = df.rename(columns=colmap)
    for sc in STANDARD_COLS:
        if sc not in df.columns:
            df[sc] = None
    return df

# ----------------------------------------------------------
# Helpers
# ----------------------------------------------------------
def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

def parse_date(s):
    try:
        return pd.to_datetime(s, errors="coerce")
    except Exception:
        return pd.NaT

def extract_date_query(q):
    q = q.lower().strip()
    today = pd.Timestamp.today().normalize()

    m_between = re.search(r"between\s+(\d{4}-\d{2}-\d{2})\s+(?:and|to)\s+(\d{4}-\d{2}-\d{2})", q)
    if m_between:
        d1, d2 = parse_date(m_between.group(1)), parse_date(m_between.group(2))
        if pd.notna(d1) and pd.notna(d2):
            return ("between", d1, d2)

    if "< today" in q:
        return ("before", today, None)
    m_before = re.search(r"<\s*(\d{4}-\d{2}-\d{2})", q)
    if m_before:
        return ("before", parse_date(m_before.group(1)), None)

    if "> today" in q:
        return ("after", today, None)
    m_after = re.search(r">\s*(\d{4}-\d{2}-\d{2})", q)
    if m_after:
        return ("after", parse_date(m_after.group(1)), None)

    return None

# ----------------------------------------------------------
# Query Logic
# ----------------------------------------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower()
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["agreed"] = pd.to_numeric(working["agreed"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")

    # Vendor filter
    vendor_match = None
    for v in working["vendor_name"].dropna().unique():
        if v.lower() in ql:
            vendor_match = v
            break
    if vendor_match:
        working = working[working["vendor_name"].astype(str).str.lower() == vendor_match.lower()]

    # Agreed logic (1 = paid, 0 = open)
    if "open" in ql or "unpaid" in ql:
        working = working[working["agreed"] == 0]
    elif "paid" in ql or "approved" in ql:
        working = working[working["agreed"] == 1]

    # Due date filters
    if "due date" in ql:
        cond = extract_date_query(ql)
        if cond:
            mode, d1, d2 = cond
            if mode == "before" and pd.notna(d1):
                working = working[working["due_date_parsed"] < d1]
            elif mode == "after" and pd.notna(d1):
                working = working[working["due_date_parsed"] > d1]
            elif mode == "between" and pd.notna(d1) and pd.notna(d2):
                working = working[(working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)]

    # Email queries
    if "email" in ql or "emails" in ql:
        emails = working["vendor_email"].dropna().astype(str).str.strip().unique()
        if len(emails) == 0:
            return "❌ No vendor emails found for this query.", None
        return f"📧 Vendor emails: {'; '.join(emails)}", None

    # Vendor summary (open & paid)
    if vendor_match:
        cur = working["currency"].dropna().iloc[0] if working["currency"].notna().any() else "EUR"
        open_df = working[working["agreed"] == 0]
        paid_df = working[working["agreed"] == 1]

        msg = f"📊 **Vendor {vendor_match} summary:**\n"
        if len(open_df) > 0:
            msg += f"- Open invoices: {len(open_df)}, total {fmt_money(open_df['amount'].sum(), cur)}\n"
        else:
            msg += "- No open invoices.\n"
        if len(paid_df) > 0:
            msg += f"- Paid invoices: {len(paid_df)}, total {fmt_money(paid_df['amount'].sum(), cur)}"
        else:
            msg += "- No paid invoices."

        details = working[["alternative_document", "due_date", "amount", "currency", "agreed"]]
        return msg, details

    # Amount summary
    if "amount" in ql or "total" in ql:
        if working.empty:
            return "❌ No matching invoices found.", None
        total = working["amount"].sum()
        cur = working["currency"].dropna().iloc[0] if working["currency"].notna().any() else "EUR"
        header = "💰 "
        if "open" in ql:
            header += "Total open amount"
        elif "paid" in ql:
            header += "Total paid amount"
        else:
            header += "Total amount"
        if vendor_match:
            header += f" for {vendor_match}"
        return f"{header}: **{fmt_money(total, cur)}**", working

    if working.empty:
        return "❌ No invoices match your query.", None

    return f"Found **{len(working)}** matching invoices.**", working

# ----------------------------------------------------------
# UI
# ----------------------------------------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Alternative Document, Vendor Name, Vendor Email, Amount, Currency, Due Date, Agreed.")

uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        # Read Excel safely from bytes (bypass Streamlit internal check)
        file_bytes = uploaded.getvalue()
        df = pd.read_excel(file_bytes, dtype=str, header=0, mangle_dupe_cols=True)

        # Normalize and store
        df = normalize_columns(df)
        st.session_state.df = df

        st.success("✅ Excel loaded successfully.")
        st.dataframe(df.head(25), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")

st.subheader("Chat")

if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask: e.g. 'vendor Technogym Iberia summary', 'open amount for vendor test', 'due date invoices < today'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if isinstance(result_df, pd.DataFrame) and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)
        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download results as CSV", csv, file_name="results.csv", mime="text/csv")
