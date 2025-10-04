import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'emails for unpaid', 'open over 1000', 'oldest unpaid invoice', 'vendor names for open'")

# --------------------------------------------------------------------
# COLUMN NORMALIZATION
# --------------------------------------------------------------------
SYNONYMS = {
    "invoice_no": ["invoice no", "invoice number", "invoice", "inv", "inv no", "inv#", "document", "doc no"],
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

# --------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------
def detect_invoice_ids(text: str):
    text = text.lower()
    candidates = re.findall(r"\b[a-z]{2,}[0-9]+[-/0-9a-z]*\b", text)
    ignore_words = {
        "paid","open","pending","invoice","invoices","inv","unpaid",
        "email","emails","mail","for","vendor","amount","currency",
        "due","payment","date","show","over","under","older","oldest",
        "newest","what","is","give","find","sum","total","before","after","since","on"
    }
    return [t for t in candidates if t not in ignore_words and not t.isalpha()]

def find_best_invoice_match(df, inv):
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    def normalize(x): return re.sub(r"[-_\s]", "", str(x).strip().lower())
    inv_norm = normalize(inv)
    df["__inv_norm__"] = df["invoice_no"].astype(str).apply(normalize)
    exact = df[df["__inv_norm__"] == inv_norm]
    if not exact.empty:
        return exact
    return df[df["__inv_norm__"].str.contains(inv_norm, na=False)]

def parse_user_date(s: str):
    for fmt in ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"]:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return pd.NaT

def extract_date_from_query(ql: str):
    m = re.search(r"(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", ql)
    if not m:
        return None
    dt = parse_user_date(m.group(1))
    return dt if pd.notna(dt) else None

def fmt_money(amount, currency):
    try:
        val = float(str(amount).replace(",", ""))
        return f"{val:,.2f} {currency or 'EUR'}"
    except Exception:
        return str(amount)

# --------------------------------------------------------------------
# MAIN QUERY LOGIC
# --------------------------------------------------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel first.", None

    ql = q.lower()
    inv_ids = detect_invoice_ids(ql)
    working = df.copy()
    # ✅ Normalize statuses and vendor emails for consistent matching
    working["status"] = working["status"].astype(str).str.strip().str.lower()
    working["vendor_name"] = working["vendor_name"].astype(str).str.strip()
    working["vendor_email"] = working["vendor_email"].astype(str).str.strip()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")

    wants_email = "email" in ql or "emails" in ql
    wants_open = any(k in ql for k in ["open", "unpaid", "pending"])
    wants_paid = "paid" in ql and not wants_open

    # Status filters
    # Status filters
    if wants_open:
        working = working[working["status"].str.contains("open|unpaid|pending", case=False, na=False)]
    elif wants_paid:
        working = working[working["status"].str.contains("paid", case=False, na=False)]

    # Amount filters
    m = re.search(r"(over|above|greater than|>=|more than)\s*([0-9][0-9,\.]*)", ql)
    if m:
        val = float(m.group(2).replace(",", ""))
        working = working[working["amount"] >= val]

    m2 = re.search(r"(under|below|less than|<=)\s*([0-9][0-9,\.]*)", ql)
    if m2:
        val2 = float(m2.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Date filters
    user_date = extract_date_from_query(ql)
    if user_date is not None:
        if "before" in ql:
            working = working[working["due_date_parsed"] < user_date]
        elif "after" in ql or "since" in ql:
            working = working[working["due_date_parsed"] > user_date]

    # Vendor filters
    vm = re.search(r"for\s+([a-z0-9 ._-]+)", ql)
    if vm:
        vendor = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(vendor, case=False, na=False)]

    # Specific invoice
    if inv_ids:
        answers, hits = [], pd.DataFrame()
        for inv in inv_ids:
            res = find_best_invoice_match(df, inv)
            if res.empty:
                answers.append(f"❓ Invoice **{inv}** not found.")
            else:
                r = res.iloc[0]
                answers.append(
                    f"Invoice **{inv}** from **{r.get('vendor_name','-')}** — "
                    f"Status: **{r.get('status','-')}**, Amount: **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                    f"Due: **{r.get('due_date','-')}**."
                )
                hits = pd.concat([hits, res])
        return "\n\n".join(answers), hits

    # EMAIL QUERIES
    if wants_email:
        emails = working["vendor_email"].dropna().astype(str).str.strip()
        emails = [e for e in emails if e]
        emails = sorted(set(emails), key=str.lower)
        if not emails:
            return "No vendor emails found for this query.", None
        if wants_open:
            return f"📧 Vendor emails for open/unpaid invoices:\n\n" + "; ".join(emails), working.reset_index(drop=True)
        elif wants_paid:
            return f"📧 Vendor emails for paid invoices:\n\n" + "; ".join(emails), working.reset_index(drop=True)
        else:
            return f"📧 All vendor emails:\n\n" + "; ".join(emails), working.reset_index(drop=True)

    # TOTALS
    if "sum" in ql or "total" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount: **{total:,.2f} EUR**", working

    # OLDEST INVOICE
    if "oldest" in ql:
        old = working.sort_values("due_date_parsed", ascending=True).head(1)
        if old.empty:
            return "No invoices found with valid due dates.", None
        r = old.iloc[0]
        return (
            f"📄 Oldest invoice: **{r.get('invoice_no','-')}** from **{r.get('vendor_name','-')}**, "
            f"due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
            f"status **{r.get('status','-')}**.",
            old
        )

    # DEFAULT
    if working.empty:
        return "No invoices match your query.", None
    return f"Found **{len(working)}** invoices matching your filters.", working.reset_index(drop=True)

# --------------------------------------------------------------------
# STREAMLIT INTERFACE
# --------------------------------------------------------------------
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

prompt = st.chat_input("Ask about invoices, e.g. 'emails for unpaid', 'open over 1000', 'oldest unpaid invoice'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)