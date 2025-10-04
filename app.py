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
        "paid","open","pending","invoice","invoices","inv","unpaid","status","email","emails",
        "mail","for","the","can","you","bring","vendor","amount","currency","due","payment",
        "date","show","list","over","under","below","older","oldest","newest","what","is",
        "tell","give","please","find","greater","than","less","more","sum","total","before",
        "after","since","on"
    }
    filtered = []
    for t in candidates:
        if t in ignore_words or t.isalpha():
            continue
        filtered.append(t)
    seen, result = set(), []
    for t in filtered:
        if t not in seen:
            seen.add(t)
            result.append(t)
    return result

# ---------------------------
# Helpers
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

def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

def find_best_invoice_match(df, inv):
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    def norm(x): return re.sub(r"[-_\s]", "", str(x).strip().lower())
    inv_norm = norm(inv)
    df["__n"] = df["invoice_no"].astype(str).apply(norm)
    exact = df[df["__n"] == inv_norm]
    return exact if not exact.empty else df[df["__n"].str.contains(inv_norm, na=False)]

# ---------------------------
# Run query
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel file first.", None

    ql = q.lower()
    inv_ids = detect_invoice_ids(ql)
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")

    # Filter by status
    if "open" in ql or "unpaid" in ql or "pending" in ql:
        working = working[working["status"].astype(str).str.contains("open|unpaid|pending", case=False, na=False)]
    elif "paid" in ql:
        working = working[working["status"].astype(str).str.contains("paid", case=False, na=False)]

    # Filter by amount
    m = re.search(r"(over|above|greater than|>=|more than)\s*([0-9][0-9,\.]*)", ql)
    if m:
        val = float(m.group(2).replace(",", ""))
        working = working[working["amount"] >= val]

    # Filter by date
    user_date = extract_date_from_query(ql)
    if user_date is not None:
        if "before" in ql:
            working = working[working["due_date_parsed"] < user_date]
        elif "after" in ql or "since" in ql:
            working = working[working["due_date_parsed"] > user_date]

    # Filter by vendor
    vm = re.search(r"for\s+([a-z0-9 ._-]+)", ql)
    if vm:
        vendor = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(vendor, case=False, na=False)]

    # 1️⃣ Specific invoice lookup
    if inv_ids:
        answers, hits = [], pd.DataFrame()
        for inv in inv_ids:
            res = find_best_invoice_match(df, inv)
            if res.empty:
                answers.append(f"❓ Invoice **{inv}** not found.")
            else:
                r = res.iloc[0]
                ans = f"Invoice **{inv}** from **{r.get('vendor_name','-')}** — Status: **{r.get('status','-')}**, Amount: **{fmt_money(r.get('amount'), r.get('currency'))}**, Due: **{r.get('due_date','-')}**."
                answers.append(ans)
                hits = pd.concat([hits, res])
        return "\n\n".join(answers), hits

    # 2️⃣ EMAIL QUERIES
    if "email" in ql or "emails" in ql:
        emails = working["vendor_email"].dropna().astype(str).str.strip()
        emails = [e for e in emails if e]
        emails = sorted(set(emails), key=str.lower)
        if not emails:
            return "No vendor emails found for this query.", None
        return f"📧 Found **{len(emails)}** vendor emails:\n\n" + "; ".join(emails), working.reset_index(drop=True)

    # 3️⃣ TOTALS
    if "sum" in ql or "total" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount for this selection: **{total:,.2f} EUR**", working

    # 4️⃣ OLDEST
    if "oldest" in ql:
        old = working.sort_values("due_date_parsed", ascending=True).head(1)
        if old.empty:
            return "No invoices found with valid due dates.", None
        r = old.iloc[0]
        return f"📄 The oldest invoice is **{r.get('invoice_no','-')}** from **{r.get('vendor_name','-')}**, due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, status: **{r.get('status','-')}**.", old

    # Default
    if working.empty:
        return "No invoices match your query.", None
    return f"Found **{len(working)}** invoices matching your filters.", working.reset_index(drop=True)

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

prompt = st.chat_input("Ask: 'emails for unpaid invoices', 'vendor emails for open', 'sum over 1000', 'oldest invoice'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)