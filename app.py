import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'vendor name for INV1003', 'open amount for Technogym Iberia', 'paid invoices for Sani Resort', 'due between 2024-10-01 and 2025-02-01'")

# ----------------------------------------------------------
# Column normalization
# ----------------------------------------------------------
SYNONYMS = {
    "invoice_no": ["invoice no", "invoice number", "invoice", "inv", "inv no", "inv#", "inv num", "document no", "doc no", "document"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name"],
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

# ----------------------------------------------------------
# Helpers
# ----------------------------------------------------------
def detect_invoices(text):
    """
    Detect invoice numbers, including variants with -, _, or /.
    e.g. INV1001, INV-1001, INV_1001, INV/1001
    """
    text = re.sub(r"[,;]", " ", text.lower())
    found = re.findall(r"\b[a-z]{2,}[0-9]+(?:[-_/]?[0-9a-z]+)*\b", text)
    cleaned = [re.sub(r"[-_/]", "", f).strip() for f in found]
    return list(dict.fromkeys(cleaned))  # remove duplicates while preserving order


def parse_date_token(q):
    q = q.lower()
    m_between = re.search(r"between\s+(\d{4}-\d{2}-\d{2})\s+(?:and|to)\s+(\d{4}-\d{2}-\d{2})", q)
    if m_between:
        try:
            d1 = pd.to_datetime(m_between.group(1), errors="coerce")
            d2 = pd.to_datetime(m_between.group(2), errors="coerce")
            if pd.notna(d1) and pd.notna(d2):
                return ("between", d1, d2)
        except Exception:
            pass
    m_before = re.search(r"(before|smaller than|less than)\s+(\d{4}-\d{2}-\d{2})", q)
    if m_before:
        return ("before", pd.to_datetime(m_before.group(2), errors="coerce"), None)
    m_after = re.search(r"(after|greater than|bigger than)\s+(\d{4}-\d{2}-\d{2})", q)
    if m_after:
        return ("after", pd.to_datetime(m_after.group(2), errors="coerce"), None)
    return None

def fmt_money(a, cur="EUR"):
    try:
        val = float(str(a).replace(",", ""))
        return f"{val:,.2f} {cur}"
    except:
        return str(a)

# ----------------------------------------------------------
# Main Query Logic
# ----------------------------------------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel first.", None

    ql = q.lower()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
    df["status"] = df["status"].astype(str).str.lower()

    # -------------------------------------
    # Detect vendor name from the question
    # -------------------------------------
    vendor_match = None
    for v in df["vendor_name"].dropna().unique():
        if v.lower() in ql:
            vendor_match = v
            break
    if vendor_match:
        df = df[df["vendor_name"].astype(str).str.lower() == vendor_match.lower()]

    # Handle specific invoices
    invoices = detect_invoices(ql)
    if invoices:
        df["__inv_norm__"] = df["invoice_no"].astype(str).str.lower().str.replace(r"[-_/]", "", regex=True)
        rows = []
        for inv in invoices:
            match = df[df["__inv_norm__"].str.contains(inv, na=False)]
            if not match.empty:
                rows.append(match)
        if not rows:
            return f"❌ No results for invoices: {', '.join(invoices)}", None
        result = pd.concat(rows).drop_duplicates()

        if "vendor" in ql:
            vendors = "; ".join(result["vendor_name"].dropna().unique())
            return f"🏢 Vendors: {vendors}", result
        elif "email" in ql:
            emails = "; ".join(result["vendor_email"].dropna().unique())
            return f"📧 Emails: {emails}", result
        elif "amount" in ql:
            details = [f"{r['invoice_no']}: {fmt_money(r['amount'], r['currency'])}" for _, r in result.iterrows()]
            return f"💰 Amounts:\n" + "\n".join(details), result
        elif "status" in ql:
            statuses = "; ".join(result["status"].dropna().unique())
            return f"📊 Status: {statuses}", result
        elif "due" in ql:
            dues = "; ".join(result["due_date"].dropna().astype(str).unique())
            return f"📅 Due dates: {dues}", result
        else:
            return f"📄 Invoices found: {', '.join(result['invoice_no'].astype(str).unique())}", result

    # Status filters
    if "open" in ql:
        df = df[df["status"].str.contains("open|pending", na=False)]
    elif "paid" in ql:
        df = df[df["status"].str.contains("paid", na=False)]

    # Date filters
    date_cond = parse_date_token(ql)
    if date_cond:
        mode, d1, d2 = date_cond
        if mode == "before" and pd.notna(d1):
            df = df[df["due_date_parsed"] < d1]
        elif mode == "after" and pd.notna(d1):
            df = df[df["due_date_parsed"] > d1]
        elif mode == "between" and pd.notna(d1) and pd.notna(d2):
            df = df[(df["due_date_parsed"] >= d1) & (df["due_date_parsed"] <= d2)]

    if df.empty:
        return "❌ No invoices match your query.", None

    # Column-specific answers
    if "vendor" in ql:
        vendors = "; ".join(df["vendor_name"].dropna().unique())
        return f"🏢 Vendors: {vendors}", df
    elif "email" in ql:
        emails = "; ".join(df["vendor_email"].dropna().unique())
        return f"📧 Emails: {emails}", df
    elif "amount" in ql:
        details = [f"{r['invoice_no']}: {fmt_money(r['amount'], r['currency'])}" for _, r in df.iterrows()]
        header = f"💰 {'Open' if 'open' in ql else 'Paid' if 'paid' in ql else ''} invoice amounts"
        if vendor_match:
            header += f" for {vendor_match}"
        return header + ":\n" + "\n".join(details), df
    elif "invoice" in ql and "amount" not in ql:
        invoices_list = "; ".join(df["invoice_no"].dropna().unique())
        header = "📄 Invoice numbers"
        if vendor_match:
            header += f" for {vendor_match}"
        return f"{header}: {invoices_list}", df
    elif "due" in ql:
        due_dates = "; ".join(df["due_date"].dropna().astype(str).unique())
        return f"📅 Due dates for matching invoices: {due_dates}", df

    return f"Found {len(df)} invoice(s) matching your query.", df

# ----------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Invoice No, Vendor Name, Vendor Email, Status, Amount, Currency, Due Date, Payment Date, PO Number")

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

prompt = st.chat_input("Ask: e.g. 'open amount for Technogym Iberia', 'paid invoices for Sani Resort', 'vendor name for INV1003'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, df_out = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if df_out is not None and not df_out.empty:
        st.dataframe(df_out, use_container_width=True)
