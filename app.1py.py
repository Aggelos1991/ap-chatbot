import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'show open over 560', 'emails for unpaid invoices', 'due before 2024-10-01', 'oldest 10 unpaid'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no": ["invoice no", "invoice number", "invoice", "inv", "inv no", "inv#", "inv num", "document no", "doc no", "document"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name", "proveedor", "provider", "provider name"],
    "vendor_email": ["email", "vendor email", "supplier email", "correo", "mail", "contact email"],
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
# Strict invoice ID detection (avoid normal words)
# ---------------------------
def detect_invoice_ids(text: str):
    """
    Detect only true invoice IDs like INV1001, INV-1001, ESF1-2025.
    """
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
# Helpers: parse dates & amounts & N (top)
# ---------------------------
def parse_user_date(s: str):
    """
    Parse user date string robustly (tries dayfirst, then monthfirst).
    Supports 2024-10-01, 01/10/2024, 10/01/2024, etc.
    """
    try:
        # try dayfirst first (EU style)
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

def extract_top_n(ql: str):
    # "oldest 10", "top 5 oldest", "newest 3"
    m = re.search(r"(?:top\s*)?(\d+)\s*(oldest|newest)", ql)
    if not m:
        # also accept "oldest" alone -> default 10
        if "oldest" in ql or "newest" in ql:
            return 10
        return None
    return int(m.group(1))

def wants_oldest(ql: str):
    return "oldest" in ql or "earliest" in ql or "older" in ql

def wants_newest(ql: str):
    return "newest" in ql or "latest" in ql or "newer" in ql

# ---------------------------
# Build answers for single invoice
# ---------------------------
def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel first.", None

    ql = q.lower()
    inv_ids = detect_invoice_ids(ql)

    # 1) Direct invoice questions
    if inv_ids:
        answers, hits = [], pd.DataFrame()
        for inv in inv_ids:
            res = find_best_invoice_match(df, inv)
            if res.empty:
                answers.append(f"❓ Could not find invoice **{inv}**.")
            else:
                row = res.iloc[0]
                status = str(row.get("status", "-"))
                vend = row.get("vendor_name", "-")
                email = row.get("vendor_email", "-")
                amount = fmt_money(row.get("amount"), row.get("currency"))
                if any(k in ql for k in ["email","emails","mail","correo"]):
                    answers.append(f"The vendor email for **{inv}** ({vend}) is **{email or '-'}**.")
                elif any(k in ql for k in ["amount","total","value","importe"]):
                    answers.append(f"Invoice **{inv}** ({vend}) has an amount of **{amount}**.")
                elif any(k in ql for k in ["currency","moneda"]):
                    answers.append(f"Currency for **{inv}** ({vend}): **{row.get('currency','-') or '-'}**.")
                elif any(k in ql for k in ["due","venc"]):
                    answers.append(f"Due date for **{inv}** ({vend}): **{row.get('due_date','-') or '-'}**.")
                elif any(k in ql for k in ["payment date","paid date","fecha pago"]):
                    answers.append(f"Payment date for **{inv}** ({vend}): **{row.get('payment_date','-') or '-'}**.")
                else:
                    # status default
                    stxt = status.lower()
                    if "paid" in stxt:
                        answers.append(f"✅ Yes, invoice **{inv}** from **{vend}** is **PAID**.")
                    elif any(k in stxt for k in ["open","unpaid","pending"]):
                        answers.append(f"🕓 Invoice **{inv}** from **{vend}** is **{status.upper()}**.")
                    else:
                        answers.append(f"Invoice **{inv}** from **{vend}** has status: **{status or '-'}**.")
                hits = pd.concat([hits, res], axis=0)
        return "\n\n".join(answers), hits.reset_index(drop=True)

    # 2) Broad queries (no explicit invoice)
    working = df.copy()

    # numeric columns
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")

    # parse due dates
    # oldest/newest sorting + top N
    if "due_date_parsed" in working.columns:
        n = extract_top_n(ql)
        if wants_oldest(ql):
            working = working.sort_values("due_date_parsed", ascending=True)
            if n: working = working.head(n)
        elif wants_newest(ql):
            working = working.sort_values("due_date_parsed", ascending=False)
            if n: working = working.head(n)

    # amount filters
    m = re.search(r"(over|above|greater than|>=|more than)\s*([0-9][0-9,\.]*)", ql)
    if m:
        val = float(m.group(2).replace(",", ""))
        working = working[working["amount"] >= val]

    m2 = re.search(r"(under|below|less than|<=)\s*([0-9][0-9,\.]*)", ql)
    if m2:
        val2 = float(m2.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # vendor contains filter e.g. "for Iberia"
    vm = re.search(r"for\s+([a-z0-9 ._-]+)", ql)
    if vm:
        needle = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(re.escape(needle), case=False, na=False)]

    # ---- DUE DATE FILTERS ----
    # date present?
    user_date = extract_date_from_query(ql)
    if user_date is not None and "due_date_parsed" in working.columns:
        if any(k in ql for k in ["before","earlier than","<"]):
            working = working[working["due_date_parsed"] < user_date]
        elif any(k in ql for k in ["after","since",">","over"]):
            working = working[working["due_date_parsed"] > user_date]
        elif "on" in ql:
            working = working[working["due_date_parsed"].dt.date == user_date.date()]

    # oldest/newest sorting + top N
    if "due_date_parsed" in working.columns:
        n = extract_top_n(ql)
        if wants_oldest(ql):
            working = working.sort_values("due_date_parsed", ascending=True)
            if n: working = working.head(n)
        elif wants_newest(ql):
            working = working.sort_values("due_date_parsed", ascending=False)
            if n: working = working.head(n)

    # EMAIL collection for current filter
    if "email" in ql or "emails" in ql:
        emails = working["vendor_email"].dropna().astype(str).str.strip()
        emails = [e for e in emails if e]  # remove blanks
        emails = sorted(set(emails), key=str.lower)
        if not emails:
            return "No emails found for the requested criteria.", None
        return f"📧 **{len(emails)} emails**:\n\n" + "; ".join(emails), working.reset_index(drop=True)

    # totals
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
# Optional in-app restart button
if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask: 'show open over 560', 'emails for unpaid', 'due before 2024-10-01', 'oldest 10 unpaid'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if result_df is not None and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)