import re
import io
import unicodedata
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="Accounts Payable Chatbot", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Examples: 'open invoices', 'emails for paid invoices', 'group by vendor', "
           "'due date between 2025-01-01 and 2025-03-31', "
           "'comment for Technogym Iberia: waiting bank confirmation', "
           "'blocked for payment'")

# =========================
# HELPERS
# =========================
def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^\w]+", "_", regex=True)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )
    return df

def dedupe_headers(cols):
    seen = {}
    out = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out

def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    base_map = {
        "supp_name": "vendor_name",
        "supplier": "vendor_name",
        "vendor": "vendor_name",
        "trade_account": "trade_account",
        "document": "invoice_no",
        "invoice_number": "invoice_no",
        "open_amount_in_base_cur": "amount",
        "open_amount": "amount",
        "amount_in_eur": "amount",
        "currency": "currency",
        "due_date": "due_date",
        "issue_date": "issue_date",
        "payment_method_doc": "payment_method",
        "payment_method_supplier": "payment_method",
        "workflow_step": "workflow_step",
        "agreeded": "agreed",
        "agreed": "agreed",
        "email": "vendor_email",
        "ηλεκτρονική_διεύθυνση": "vendor_email",
        "διευθυνση": "vendor_email",
        "correo": "vendor_email",
        "alternative_document": "alternative_document",
    }
    df = df.rename(columns=lambda c: base_map.get(c, c))
    return df

def parse_date_filter(q: str):
    today = datetime.today()
    q = q.lower()
    between = re.search(r"between\s+(\d{4}-\d{2}-\d{2})\s+(?:and|to)\s+(\d{4}-\d{2}-\d{2})", q)
    less = re.search(r"<\s*(today|\d{4}-\d{2}-\d{2})", q)
    greater = re.search(r">\s*(today|\d{4}-\d{2}-\d{2})", q)
    if between:
        return "between", pd.to_datetime(between.group(1)), pd.to_datetime(between.group(2))
    if less:
        d = today if "today" in less.group(1) else pd.to_datetime(less.group(1))
        return "before", d, None
    if greater:
        d = today if "today" in greater.group(1) else pd.to_datetime(greater.group(1))
        return "after", d, None
    return None

# =========================
# QUERY ENGINE
# =========================
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel first.", None

    ql = q.lower()
    df = df.copy()

    # Prepare main columns
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    if "due_date" in df.columns:
        df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")

    # Define invoice column fallback
    invoice_col = "invoice_no"
    if invoice_col not in df.columns:
        if "alternative_document" in df.columns:
            invoice_col = "alternative_document"

    # Ensure 'agreed' column
    if "agreed" in df.columns:
        df["agreed"] = pd.to_numeric(df["agreed"], errors="coerce").fillna(0).astype(int)
    else:
        df["agreed"] = 0

    # ---------------- FILTERS ----------------
    is_open = any(k in ql for k in ["open", "unpaid", "pending"])
    is_paid = any(k in ql for k in ["paid", "settled", "agreed", "aggreed", "approved", "posted", "reconciled"])

    if is_open:
        df = df[df["agreed"] == 0]
    elif is_paid:
        df = df[df["agreed"] == 1]

    # Blocked for payment
    if "blocked" in ql or "block" in ql:
        if "workflow_step" in df.columns:
            df = df[df["workflow_step"].astype(str).str.contains("block", case=False, na=False)]
            if df.empty:
                return "✅ No invoices currently blocked for payment.", None
            else:
                return f"🚫 Found {len(df)} blocked invoices.", df
        else:
            return "⚠️ No 'workflow_step' column found to check blocked invoices.", None

    # Vendor filter
    vendor_match = None
    if "vendor_name" in df.columns:
        for v in df["vendor_name"].dropna().astype(str).unique():
            if v.lower() in ql:
                vendor_match = v
                df = df[df["vendor_name"].astype(str).str.lower() == v.lower()]
                break

    # Date filter
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

    # ---------------- EMAILS ----------------
    if "email" in ql:
        if "vendor_email" not in df.columns:
            return "⚠️ No 'vendor_email' column found.", None

        # Split by open / paid context
        if "open" in ql:
            emails_df = df[df["agreed"] == 0]
        elif "paid" in ql:
            emails_df = df[df["agreed"] == 1]
        else:
            emails_df = df

        emails = "; ".join(sorted({str(x).strip() for x in emails_df["vendor_email"].dropna() if str(x).strip()}))
        return f"📧 Vendor emails: {emails if emails else 'none found'}", None

    # ---------------- GROUP BY VENDOR ----------------
    if "group" in ql or "vendor by amount" in ql or "amount per vendor" in ql:
        if "vendor_name" not in df.columns or "amount" not in df.columns:
            return "⚠️ Missing 'vendor_name' or 'amount' columns.", None
        grouped = (
            df.groupby("vendor_name", dropna=True)
              .agg(total_amount=("amount", "sum"), invoices=(invoice_col, "count"))
              .reset_index()
        )
        grouped["total_amount"] = grouped["total_amount"].map(lambda x: f"{x:,.2f}")
        return "📊 Totals by vendor:", grouped

    # ---------------- VENDOR SUMMARY ----------------
    if vendor_match and "summary" in ql:
        vendor_all = df.copy()
        open_df = vendor_all[vendor_all["agreed"] == 0]
        paid_df = vendor_all[vendor_all["agreed"] == 1]
        msg = (f"📊 Vendor **{vendor_match}** summary:\n"
               f"- Open invoices: {len(open_df)} (Total {open_df['amount'].sum():,.2f})\n"
               f"- Paid invoices: {len(paid_df)} (Total {paid_df['amount'].sum():,.2f})")
        return msg, df

    # ---------------- TOTALS ----------------
    if "open amount" in ql or "open amounts" in ql:
        total = df["amount"].sum(skipna=True)
        return f"💶 Total open amount: {total:,.2f}", df

    if "paid amount" in ql or "paid amounts" in ql:
        total = df["amount"].sum(skipna=True)
        return f"✅ Total paid amount: {total:,.2f}", df

    # ---------------- DEFAULT ----------------
    return f"Found **{len(df)}** matching invoices.", df

# =========================
# FILE UPLOAD
# =========================
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
        headers_raw = [str(h).strip() if h else f"Unnamed_{i}" for i, h in enumerate(data[0])]
        headers_dedup = dedupe_headers(headers_raw)
        df = pd.DataFrame(data[1:], columns=headers_dedup)
        df = clean_headers(df)
        df.columns = dedupe_headers(df.columns)
        df = normalize_columns(df)
        st.session_state.df = df
        st.success(f"✅ Excel loaded: {len(df)} rows | {len(df.columns)} cols.")
        st.dataframe(df.head(30), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

# =========================
# CHAT + COMMENTS
# =========================
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

    # ----- Comments -----
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
                resp = "ℹ️ Use: `comment for VendorName: your note`"
        st.session_state.history.append(("assistant", resp))
        st.chat_message("assistant").write(resp)
    else:
        # ----- Normal query -----
        answer, df_out = run_query(prompt, st.session_state.df)
        st.session_state.history.append(("assistant", answer))
        st.chat_message("assistant").write(answer)
        if df_out is not None and not df_out.empty:
            st.dataframe(df_out, use_container_width=True)
