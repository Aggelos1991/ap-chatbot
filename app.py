import re
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Accounts Payable Chatbot", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Try: 'open amounts', 'emails for paid invoices', 'due date < today', 'group by vendor', 'comment for Technogym Iberia: awaiting payment'")

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_headers(df):
    """Normalize column names: lowercase, underscores, no spaces."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )
    return df

def normalize_columns(df):
    """Map alternative header names to standardized ones."""
    mapping = {
        "supp_name": "vendor_name",
        "supplier": "vendor_name",
        "vendor": "vendor_name",
        "document": "invoice_no",
        "invoice_number": "invoice_no",
        "open_amount_in_base_cur": "amount",
        "open_amount": "amount",
        "amount_in_eur": "amount",
        "ηλεκτρονική_διευθυνση": "vendor_email",
        "email": "vendor_email",
        "payment_method_doc": "payment_method",
        "payment_method_supplier": "payment_method",
    }
    df = df.rename(columns=lambda c: mapping.get(c, c))
    return df

def parse_date_filter(q):
    """Extract date filters from user text."""
    today = datetime.today()
    q = q.lower()
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

    # Prepare columns safely
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    if "due_date" in df.columns:
        df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")

    # Ensure 'agreed' exists
    if "agreed" in df.columns:
        df["agreed"] = pd.to_numeric(df["agreed"], errors="coerce").fillna(0).astype(int)
    else:
        df["agreed"] = 0

    # -------------------- FILTERS --------------------
    if "open" in ql:
        df = df[df["agreed"] == 0]
    elif "paid" in ql:
        df = df[df["agreed"] == 1]

    # Vendor filter
    vendor_match = None
    if "vendor_name" in df.columns:
        for v in df["vendor_name"].dropna().unique():
            if v.lower() in ql:
                vendor_match = v
                df = df[df["vendor_name"].str.lower() == v.lower()]
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

    # -------------------- EMAILS --------------------
    if "email" in ql:
        if "vendor_email" not in df.columns:
            return "⚠️ No 'vendor_email' column found.", None
        emails = "; ".join(sorted(df["vendor_email"].dropna().unique()))
        return f"📧 Vendor emails: {emails if emails else 'none found'}", None

    # -------------------- GROUP BY VENDOR --------------------
    if "group" in ql or "amount per vendor" in ql:
        if "vendor_name" not in df.columns or "amount" not in df.columns:
            return "⚠️ Missing vendor_name or amount columns.", None
        grouped = (
            df.groupby("vendor_name")
            .agg(total_amount=("amount", "sum"), invoices=("invoice_no", "count"))
            .reset_index()
        )
        grouped["total_amount"] = grouped["total_amount"].map(lambda x: f"{x:,.2f}")
        return "📊 Totals by vendor:", grouped

    # -------------------- VENDOR SUMMARY --------------------
    if vendor_match and "summary" in ql:
        open_df = df[df["agreed"] == 0]
        paid_df = df[df["agreed"] == 1]
        msg = f"📊 Vendor **{vendor_match}** summary:\n"
        msg += f"- Open invoices: {len(open_df)} (Total {open_df['amount'].sum():,.2f})\n"
        msg += f"- Paid invoices: {len(paid_df)} (Total {paid_df['amount'].sum():,.2f})"
        return msg, df

    # -------------------- OPEN AMOUNTS TOTAL --------------------
    if "open amount" in ql or "open amounts" in ql:
        total = df["amount"].sum()
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

        # Deduplicate headers before DataFrame creation
        seen = {}
        final_headers = []
        for h in headers:
            if h in seen:
                seen[h] += 1
                final_headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 1
                final_headers.append(h)

        df = pd.DataFrame(data[1:], columns=final_headers)
        df = clean_headers(df)
        df = normalize_columns(df)
        st.session_state.df = df

        st.success(f"✅ Excel loaded: {len(df)} rows, {len(df.columns)} columns.")
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
    lower = prompt.lower()

    # -------------------- COMMENTS --------------------
    if lower.startswith("comment"):
        m = re.search(r"comment\s+for\s+(.+?):\s*(.+)", lower, re.IGNORECASE)
        if m:
            vendor, comment = m.group(1).strip(), m.group(2).strip()
            st.session_state.comments[vendor.lower()] = comment
            response = f"💬 Saved comment for {vendor}."
        else:
            m = re.search(r"comment\s+for\s+(.+)", lower, re.IGNORECASE)
            if m:
                vendor = m.group(1).strip().lower()
                note = st.session_state.comments.get(vendor)
                response = f"💬 Comment for {vendor}: {note}" if note else "⚠️ No comment found."
            else:
                response = "ℹ️ Use: `comment for VendorName: your note`"
        st.session_state.history.append(("assistant", response))
        st.chat_message("assistant").write(response)

    # -------------------- NORMAL QUERY --------------------
    else:
        ans, df_out = run_query(prompt, st.session_state.df)
        st.session_state.history.append(("assistant", ans))
        st.chat_message("assistant").write(ans)
        if df_out is not None and not df_out.empty:
            st.dataframe(df_out, use_container_width=True)
