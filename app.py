import re
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook
import unicodedata

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Accounts Payable Chatbot", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'open invoices', 'emails for paid invoices', 'due date < today', 'group by vendor', 'comment for Technogym Iberia: waiting payment'")

# ------------------------------------------------------------
# HEADER CLEANING
# ------------------------------------------------------------
def clean_headers(df):
    """Normalize headers (preserve unicode like Greek letters)."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^\w]+", "_", regex=True)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )
    return df

def strip_accents(s):
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def normalize_columns(df):
    """Map your Excel columns to standard names."""
    rename_map = {
        "trade_account": "trade_account",
        "issue_date": "issue_date",
        "due_date": "due_date",
        "document": "invoice_no",
        "invoice_number": "invoice_no",
        "invoice": "invoice_no",
        "alternative_document": "alternative_document",
        "open_amount": "amount",
        "open_amount_in_base_cur": "amount",
        "open_amount_in_base_currency": "amount",
        "amount": "amount",
        "currency": "currency",
        "due_month": "due_month",
        "payment_method_doc": "payment_method",
        "payment_method_supplier": "payment_method",
        "workflow_step": "workflow_step",
        "agreed": "agreed",
        "aggreed": "agreed",
        "approved": "agreed",
        "supp_name": "vendor_name",
        "supplier": "vendor_name",
        "vendor": "vendor_name",
        "vendor_name": "vendor_name",
        "ηλεκτρονική_διευθυνση": "vendor_email",
        "διευθυνση": "vendor_email",
        "email": "vendor_email",
        "correo": "vendor_email",
        "mail": "vendor_email",
        "payment_date": "payment_date",
    }
    # Normalize Greek-like or accented columns
    cols = {}
    for c in df.columns:
        c_plain = strip_accents(c)
        match = next((k for k in rename_map.keys() if k in c_plain), c)
        cols[c] = rename_map.get(match, c)
    df = df.rename(columns=cols)
    return df

# ------------------------------------------------------------
# DATE PARSER
# ------------------------------------------------------------
def parse_date_filter(q):
    q = q.lower()
    today = datetime.today()
    between = re.search(r"between\s+(\d{4}-\d{2}-\d{2})\s+(?:and|to)\s+(\d{4}-\d{2}-\d{2})", q)
    less = re.search(r"<\s*(today|\d{4}-\d{2}-\d{2})", q)
    greater = re.search(r">\s*(today|\d{4}-\d{2}-\d{2})", q)
    if between:
        return "between", pd.to_datetime(between.group(1)), pd.to_datetime(between.group(2))
    elif less:
        d = today if "today" in less.group(1) else pd.to_datetime(less.group(1))
        return "before", d, None
    elif greater:
        d = today if "today" in greater.group(1) else pd.to_datetime(greater.group(1))
        return "after", d, None
    return None

# ------------------------------------------------------------
# QUERY ENGINE
# ------------------------------------------------------------
def run_query(q, df):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel first.", None

    ql = q.lower()

    # Prepare data
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    if "due_date" in df.columns:
        df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
    if "agreed" in df.columns:
        df["agreed"] = pd.to_numeric(df["agreed"], errors="coerce").fillna(0).astype(int)
    else:
        df["agreed"] = 0

    # --- FILTERS ---
    open_intent = any(k in ql for k in ["open", "unpaid", "pending"])
    paid_intent = any(k in ql for k in ["paid", "settled", "agreed", "aggreed", "approved", "posted"])

    if open_intent:
        df = df[df["agreed"] == 0]
    elif paid_intent:
        df = df[df["agreed"] == 1]

    # Vendor filter
    vendor_match = None
    if "vendor_name" in df.columns:
        for v in df["vendor_name"].dropna().astype(str).unique():
            if v.lower() in ql:
                vendor_match = v
                df = df[df["vendor_name"].astype(str).str.lower() == v.lower()]
                break

    # Date filters
    cond = parse_date_filter(ql)
    if cond and "due_date_parsed" in df.columns:
        mode, d1, d2 = cond
        if mode == "before":
            df = df[df["due_date_parsed"] < d1]
        elif mode == "after":
            df = df[df["due_date_parsed"] > d1]
        elif mode == "between":
            df = df[(df["due_date_parsed"] >= d1) & (df["due_date_parsed"] <= d2)]

    if df.empty:
        return "❌ No invoices found.", None

    # --- EMAILS ---
    if "email" in ql:
        if "vendor_email" not in df.columns:
            return "⚠️ No 'vendor_email' column found.", None
        emails = "; ".join(sorted({str(x).strip() for x in df["vendor_email"].dropna() if str(x).strip()}))
        return f"📧 Vendor emails: {emails if emails else 'none found'}", None

    # --- GROUP BY VENDOR ---
    if "group" in ql or "amount per vendor" in ql or "totals by vendor" in ql:
        if "vendor_name" not in df.columns or "amount" not in df.columns:
            return "⚠️ Missing 'vendor_name' or 'amount' columns.", None
        grouped = (
            df.groupby("vendor_name", dropna=True)
              .agg(total_amount=("amount", "sum"), invoices=("invoice_no", "count"))
              .reset_index()
        )
        grouped["total_amount"] = grouped["total_amount"].map(lambda x: f"{x:,.2f}")
        return "📊 Totals by vendor:", grouped

    # --- VENDOR SUMMARY ---
    if vendor_match and "summary" in ql:
        open_df = df[df["agreed"] == 0]
        paid_df = df[df["agreed"] == 1]
        msg = f"📊 Vendor **{vendor_match}** summary:\n"
        msg += f"- Open invoices: {len(open_df)} (Total {open_df['amount'].sum():,.2f})\n"
        msg += f"- Paid invoices: {len(paid_df)} (Total {paid_df['amount'].sum():,.2f})"
        return msg, df

    # --- OPEN AMOUNT TOTAL ---
    if "open amount" in ql or "open amounts" in ql:
        total = df["amount"].sum(skipna=True)
        return f"💶 Total open amount: {total:,.2f}", df

    return f"Found **{len(df)}** matching invoices.", df

# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------
st.sidebar.header("📦 Upload Excel")
uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        file_bytes = uploaded.getvalue()
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb.active
        data = list(ws.values)
        headers = [str(h).strip() if h else f"Unnamed_{i}" for i, h in enumerate(data[0])]
        seen = {}
        new_headers = []
        for h in headers:
            if h in seen:
                seen[h] += 1
                new_headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 1
                new_headers.append(h)
        df = pd.DataFrame(data[1:], columns=new_headers)
        df = clean_headers(df)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success(f"✅ Excel loaded: {len(df)} rows, {len(df.columns)} columns.")
        st.write("**Detected columns:**", list(df.columns))
        st.dataframe(df.head(30), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

# ------------------------------------------------------------
# CHAT
# ------------------------------------------------------------
st.subheader("Chat")

if "history" not in st.session_state:
    st.session_state.history = []
if "comments" not in st.session_state:
    st.session_state.comments = {}

if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.session_state.comments = {}
    st.rerun()

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask or add comment...")

if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    pl = prompt.lower()

    # --- COMMENTS ---
    if pl.startswith("comment"):
        m = re.search(r"comment\s+for\s+(.+?):\s*(.+)", prompt, re.IGNORECASE)
        if m:
            vendor, note = m.group(1).strip(), m.group(2).strip()
            st.session_state.comments[vendor.lower()] = note
            resp = f"💬 Saved comment for {vendor}."
        else:
            m = re.search(r"comment\s+for\s+(.+)", prompt, re.IGNORECASE)
            if m:
                vendor = m.group(1).strip()
                note = st.session_state.comments.get(vendor.lower())
                resp = f"💬 Comment for {vendor}: {note}" if note else "⚠️ No comment found."
            else:
                resp = "ℹ️ Use format: `comment for VendorName: your note`"
        st.session_state.history.append(("assistant", resp))
        st.chat_message("assistant").write(resp)
    else:
        ans, df_out = run_query(prompt, st.session_state.df)
        st.session_state.history.append(("assistant", ans))
        st.chat_message("assistant").write(ans)
        if df_out is not None and not df_out.empty:
            st.dataframe(df_out, use_container_width=True)
