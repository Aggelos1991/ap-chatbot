import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'emails for unpaid', 'vendor names for open', 'due smaller than 2025-10-01', 'oldest invoice'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no": ["invoice no", "invoice number", "inv", "document no"],
    "vendor_name": ["vendor", "vendor name", "supplier"],
    "vendor_email": ["email", "vendor email", "supplier email", "mail"],
    "status": ["status", "state", "payment status"],
    "amount": ["amount", "total", "value"],
    "currency": ["currency", "curr"],
    "due_date": ["due date", "vencimiento"],
    "payment_date": ["payment date", "paid date"],
    "po_number": ["po", "purchase order"]
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

# ---------------------------
# Extract date from text
# ---------------------------
def parse_user_date(s: str):
    try:
        return pd.to_datetime(s, dayfirst=True, errors="raise")
    except Exception:
        try:
            return pd.to_datetime(s, dayfirst=False, errors="raise")
        except Exception:
            return pd.NaT

def extract_date_from_query(ql: str):
    m = re.search(r"(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", ql)
    if not m:
        return None
    dt = parse_user_date(m.group(1))
    return dt if pd.notna(dt) else None

# ---------------------------
# Format amount
# ---------------------------
def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

# ---------------------------
# Main query engine
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel file first.", None

    ql = q.lower().strip()
    df = df.copy()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
    df["status_norm"] = df["status"].astype(str).str.lower().str.strip()

    # Filter by status
    if "open" in ql or "unpaid" in ql or "pending" in ql:
        df = df[df["status_norm"].str.contains("open|unpaid|pending", na=False)]
    elif "paid" in ql and not ("unpaid" in ql or "not paid" in ql):
        df = df[df["status_norm"].str.contains("paid", na=False)]

    # Filter by amount
    over = re.search(r"(over|above|greater than|>=|more than)\s*([0-9]+)", ql)
    under = re.search(r"(under|below|less than|<=|smaller than)\s*([0-9]+)", ql)
    if over:
        val = float(over.group(2))
        df = df[df["amount"] >= val]
    elif under:
        val = float(under.group(2))
        df = df[df["amount"] <= val]

    # Filter by due date
    user_date = extract_date_from_query(ql)
    if user_date is not None:
        if any(k in ql for k in ["before", "earlier", "<", "smaller than", "less than", "until"]):
            df = df[df["due_date_parsed"] < user_date]
        elif any(k in ql for k in ["after", "later", ">", "greater than", "bigger than", "from", "since"]):
            df = df[df["due_date_parsed"] > user_date]

    if df.empty:
        return "No invoices found for this query.", None

    # Handle focused requests
    if "email" in ql:
        emails = sorted(set(df["vendor_email"].dropna().astype(str)))
        if not emails:
            return "No vendor emails found for this query.", None
        return f"📧 Found {len(emails)} email(s):\n\n" + "; ".join(emails), df

    if "vendor name" in ql or "vendors" in ql or "supplier" in ql:
        vendors = sorted(set(df["vendor_name"].dropna().astype(str)))
        return f"🏢 Found {len(vendors)} vendor(s):\n\n" + ", ".join(vendors), df

    if "amount" in ql or "total" in ql or "value" in ql:
        total = pd.to_numeric(df["amount"], errors="coerce").sum()
        return f"💰 Total amount for this query: **{total:,.2f} EUR**", df

    if "due" in ql or "oldest" in ql:
        df = df.sort_values("due_date_parsed", ascending=True)
        oldest = df.head(1)
        if oldest.empty:
            return "No invoices with valid due date.", None
        r = oldest.iloc[0]
        return f"📄 The oldest invoice is **{r['invoice_no']}** from **{r['vendor_name']}**, due on **{r['due_date']}**, amount **{fmt_money(r['amount'], r['currency'])}**.", oldest

    return f"Found {len(df)} invoice(s) matching your query.", df

# ---------------------------
# Streamlit UI
# ---------------------------
st.sidebar.header("📦 Upload Excel File")
uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        df = pd.read_excel(uploaded, dtype=str)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success("✅ Excel loaded and columns normalized.")
        st.dataframe(df, use_container_width=True)
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

prompt = st.chat_input("Ask about invoices: e.g., 'emails for unpaid', 'vendor names for open', 'due before 2025-05-10'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)