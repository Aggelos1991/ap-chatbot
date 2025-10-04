import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'show open over 560', 'emails for unpaid invoices', 'due before 2024-10-01', 'oldest unpaid invoice'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no": ["invoice no", "invoice number", "invoice", "inv", "inv no", "inv#", "inv num", "document no", "doc no", "document"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name", "proveedor"],
    "vendor_email": ["email", "vendor email", "supplier email", "correo", "mail"],
    "status": ["status", "state", "paid?", "open?", "payment status"],
    "amount": ["amount", "total", "invoice amount", "importe", "value"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "vencimiento", "fecha vencimiento"],
    "payment_date": ["payment date", "fecha pago", "paid date"],
    "po_number": ["po", "po number", "purchase order"]
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
# Detect invoice IDs
# ---------------------------
def detect_invoice_ids(text: str):
    text = text.lower()
    candidates = re.findall(r"\b[a-z]{2,}[0-9]+[-/0-9a-z]*\b", text)
    ignore_words = {
        "paid","open","pending","invoice","invoices","inv","unpaid","status",
        "email","emails","mail","for","the","can","you","bring","vendor",
        "amount","currency","due","payment","date","show","list","over","under",
        "below","older","oldest","newest","what","is","tell","give","please",
        "find","greater","than","less","more","sum","total","before","after","since","on"
    }
    filtered = []
    for t in candidates:
        if t in ignore_words or t.isalpha():
            continue
        filtered.append(t)
    if not filtered:
        return []
    seen, result = set(), []
    for t in filtered:
        if t not in seen:
            seen.add(t)
            result.append(t)
    return result

# ---------------------------
# Invoice match helper
# ---------------------------
def find_best_invoice_match(df: pd.DataFrame, inv: str):
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    def normalize(x):
        return re.sub(r"[-_\s]", "", str(x).strip().lower())
    inv_norm = normalize(inv)
    df["__inv_norm__"] = df["invoice_no"].astype(str).apply(normalize)
    exact = df[df["__inv_norm__"] == inv_norm]
    if not exact.empty:
        return exact
    return df[df["__inv_norm__"].str.contains(inv_norm, na=False)]

# ---------------------------
# Date utilities
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
# Helpers for oldest/newest
# ---------------------------
def extract_top_n(ql: str):
    m = re.search(r"(?:top\s*)?(\d+)\s*(oldest|newest)", ql)
    if not m:
        if "oldest" in ql or "newest" in ql:
            return 10
        return None
    return int(m.group(1))

def wants_oldest(ql: str):
    return "oldest" in ql or "earliest" in ql or "older" in ql

def wants_newest(ql: str):
    return "newest" in ql or "latest" in ql or "newer" in ql

# ---------------------------
# Format money
# ---------------------------
def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

# ---------------------------
# Run query
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel first.", None

    ql = q.lower()
    inv_ids = detect_invoice_ids(ql)

    # 1️⃣ Specific invoice query
    if inv_ids:
        answers, hits = [], pd.DataFrame()
        for inv in inv_ids:
            res = find_best_invoice_match(df, inv)
            if res.empty:
                answers.append(f"❓ Could not find invoice **{inv}**.")
            else:
                row = res.iloc[0]
                vend = row.get("vendor_name", "-")
                status = str(row.get("status", "-"))
                amount = fmt_money(row.get("amount"), row.get("currency"))
                email = row.get("vendor_email", "-")
                due = row.get("due_date", "-")
                if "email" in ql:
                    answers.append(f"The vendor email for **{inv}** ({vend}) is **{email}**.")
                elif "amount" in ql:
                    answers.append(f"Invoice **{inv}** ({vend}) has an amount of **{amount}**.")
                elif "currency" in ql:
                    answers.append(f"Currency for **{inv}** ({vend}): **{row.get('currency', '-') or '-'}**.")
                elif "due" in ql:
                    answers.append(f"Due date for **{inv}** ({vend}): **{due or '-'}**.")
                else:
                    answers.append(f"Invoice **{inv}** from **{vend}** has status: **{status}**, amount **{amount}**, due **{due}**.")
                hits = pd.concat([hits, res], axis=0)
        return "\n\n".join(answers), hits.reset_index(drop=True)

    # 2️⃣ Broader filters
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")

    # Filter by status
    if "open" in ql or "unpaid" in ql or "pending" in ql:
        working = working[working["status"].astype(str).str.contains("open|unpaid|pending", case=False, na=False)]
    if "paid" in ql and not ("unpaid" in ql or "not paid" in ql):
        working = working[working["status"].astype(str).str.contains("paid", case=False, na=False)]

    # Filter by amount
    m = re.search(r"(over|above|greater than|>=|more than)\s*([0-9][0-9,\.]*)", ql)
    if m:
        val = float(m.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    m2 = re.search(r"(under|below|less than|<=)\s*([0-9][0-9,\.]*)", ql)
    if m2:
        val2 = float(m2.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Filter by vendor
    vm = re.search(r"for\s+([a-z0-9 ._-]+)", ql)
    if vm:
        needle = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(re.escape(needle), case=False, na=False)]

    # Filter by due date
    user_date = extract_date_from_query(ql)
    if user_date is not None:
        if any(k in ql for k in ["before", "earlier than", "<"]):
            working = working[working["due_date_parsed"] < user_date]
        elif any(k in ql for k in ["after", "since", ">", "over"]):
            working = working[working["due_date_parsed"] > user_date]
        elif "on" in ql:
            working = working[working["due_date_parsed"].dt.date == user_date.date()]

    # Oldest / newest invoice logic
    if "due_date_parsed" in working.columns:
        n = extract_top_n(ql)
        if wants_oldest(ql):
            working = working.sort_values("due_date_parsed", ascending=True)
            if "oldest invoice" in ql or "the oldest" in ql:
                oldest_row = working.head(1)
                if oldest_row.empty:
                    return "No invoices found with a valid due date.", None
                r = oldest_row.iloc[0]
                return (
                    f"📄 The oldest invoice is **{r.get('invoice_no','-')}** from **{r.get('vendor_name','-')}**, "
                    f"due on **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                    f"status: **{r.get('status','-')}**.",
                    oldest_row.reset_index(drop=True)
                )
            if n:
                working = working.head(n)
        elif wants_newest(ql):
            working = working.sort_values("due_date_parsed", ascending=False)
            if "newest invoice" in ql or "latest invoice" in ql:
                newest_row = working.head(1)
                if newest_row.empty:
                    return "No invoices found with a valid due date.", None
                r = newest_row.iloc[0]
                return (
                    f"📄 The newest invoice is **{r.get('invoice_no','-')}** from **{r.get('vendor_name','-')}**, "
                    f"due on **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                    f"status: **{r.get('status','-')}**.",
                    newest_row.reset_index(drop=True)
                )
            if n:
                working = working.head(n)

    # Emails
    if "email" in ql or "emails" in ql:
        emails = working["vendor_email"].dropna().astype(str).str.strip()
        emails = [e for e in emails if e]
        emails = sorted(set(emails), key=str.lower)
        if not emails:
            return "No emails found for this query.", None
        return f"📧 **{len(emails)} emails:**\n\n" + "; ".join(emails), working.reset_index(drop=True)

    if working.empty:
        return "No invoices match your query.", None

    if "sum" in ql or "total" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount: **{total:,.2f}**", working.reset_index(drop=True)

    return f"Found **{len(working)}** invoices matching your filters.", working.reset_index(drop=True)

# ---------------------------
# Streamlit UI
# ---------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Invoice No, Vendor Name, Vendor Email, Status, Amount, Currency, Due Date, Payment Date, PO Number.")

uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

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

prompt = st.chat_input("Ask about invoices: e.g., 'show open over 560', 'emails for unpaid', 'due before 2024-10-01', 'oldest unpaid invoice'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)