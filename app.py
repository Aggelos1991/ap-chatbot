import re
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook
import warnings

warnings.filterwarnings("ignore", message="Duplicate column names found", category=UserWarning)

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Accounts Payable Chatbot — Excel-driven", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Try: 'open amount for vendor test', 'emails for paid invoices', 'due date invoices < today', 'vendor Technogym Iberia summary'")

# ------------------------------------------------------------
# CLEANING HELPERS
# ------------------------------------------------------------
def clean_excel_headers(df):
    """Normalize headers: lowercase, underscores, trim spaces."""
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^a-z0-9]+", "_", regex=True)
        .str.replace(r"_+", "_", regex=True)
        .str.strip("_")
    )
    return df

def fmt_money(x, cur="EUR"):
    try:
        val = float(str(x).replace(",", "").strip())
        return f"{val:,.2f} {cur}"
    except:
        return str(x)

def parse_date_token(q):
    """Detects <, >, or between date conditions."""
    q = q.lower()
    today = datetime.today()
    between = re.search(r"between\s+(\d{4}-\d{2}-\d{2})\s+(?:and|to)\s+(\d{4}-\d{2}-\d{2})", q)
    less = re.search(r"<\s*(today|\d{4}-\d{2}-\d{2})", q)
    greater = re.search(r">\s*(today|\d{4}-\d{2}-\d{2})", q)

    if between:
        return ("between",
                pd.to_datetime(between.group(1), errors="coerce"),
                pd.to_datetime(between.group(2), errors="coerce"))
    elif less:
        d = today if "today" in less.group(1) else pd.to_datetime(less.group(1), errors="coerce")
        return ("before", d, None)
    elif greater:
        d = today if "today" in greater.group(1) else pd.to_datetime(greater.group(1), errors="coerce")
        return ("after", d, None)
    return None

def normalize_columns(df):
    """Map synonyms to consistent names."""
    colmap = {
        "trade_account": "trade_account",
        "issue_date": "issue_date",
        "due_date": "due_date",
        "document": "invoice_no",
        "invoice_no": "invoice_no",
        "alternative_document": "alternative_document",
        "open_amount": "amount",
        "open_amount_in_base_cur": "amount",
        "amount": "amount",
        "currency": "currency",
        "due_month": "due_month",
        "payment_method_doc": "payment_method",
        "payment_method_supplier": "payment_method",
        "workflow_step": "workflow_step",
        "agreed": "agreed",
        "agreeded": "agreed",
        "supp_name": "vendor_name",
        "vendor_name": "vendor_name",
        "vendor_email": "vendor_email",
        "ηλεκτρονική_διευθυνση": "vendor_email",
        "payment_date": "payment_date",
    }
    df = df.rename(columns=lambda c: colmap.get(c, c))
    return df

# ------------------------------------------------------------
# QUERY ENGINE
# ------------------------------------------------------------
def run_query(q, df):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower()
    df["amount"] = pd.to_numeric(df.get("amount"), errors="coerce")
    df["due_date_parsed"] = pd.to_datetime(df.get("due_date"), errors="coerce")
    df["agreed"] = pd.to_numeric(df.get("agreed"), errors="coerce").fillna(0).astype(int)

    # Filters
    if "open" in ql or "unpaid" in ql:
        df = df[df["agreed"] == 0]
    elif "paid" in ql:
        df = df[df["agreed"] == 1]

    # Vendor filter
    vendor_match = None
    for v in df["vendor_name"].dropna().unique():
        if v.lower() in ql:
            vendor_match = v
            df = df[df["vendor_name"].str.lower() == v.lower()]
            break

    # Date filters
    date_cond = parse_date_token(ql)
    if date_cond:
        mode, d1, d2 = date_cond
        if mode == "before":
            df = df[df["due_date_parsed"] < d1]
        elif mode == "after":
            df = df[df["due_date_parsed"] > d1]
        elif mode == "between":
            df = df[(df["due_date_parsed"] >= d1) & (df["due_date_parsed"] <= d2)]

    if df.empty:
        return "❌ No invoices match your query.", None

    # Vendor summary
    if vendor_match and "summary" in ql:
        open_df = df[df["agreed"] == 0]
        paid_df = df[df["agreed"] == 1]
        msg = f"📊 Vendor **{vendor_match}** summary:\n"
        msg += f"- Open invoices: {len(open_df)}, total {fmt_money(open_df['amount'].sum(), open_df.get('currency', 'EUR').iloc[0] if not open_df.empty else 'EUR')}\n"
        msg += f"- Paid invoices: {len(paid_df)}, total {fmt_money(paid_df['amount'].sum(), paid_df.get('currency', 'EUR').iloc[0] if not paid_df.empty else 'EUR')}"
        return msg, df

    # Emails
    if "email" in ql:
        emails = "; ".join(df["vendor_email"].dropna().unique())
        return f"📧 Vendor emails: {emails if emails else 'none found'}", df

    # Workflow
    if "workflow" in ql or "block" in ql:
        steps = df["workflow_step"].dropna().unique()
        return f"🔧 Workflow steps: {', '.join(steps)}", df

    # Totals by vendor
    if "amount" in ql and "vendor" in ql:
        grouped = df.groupby("vendor_name", dropna=True)["amount"].sum().reset_index()
        grouped["amount"] = grouped["amount"].map(lambda x: f"{x:,.2f}")
        return "📊 Totals by vendor:", grouped

    return f"Found **{len(df)}** invoice(s) matching your query.", df

# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Trade account, Issue date, Due date, Document, Alternative Document, Payment method, Workflow step, Agreed, Supp name, Email, Amount, Currency, etc.")
uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        # --- Read Excel safely with openpyxl ---
        file_bytes = uploaded.getvalue()
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

        # If multiple sheets exist, pick the first with content
        sheet_name = None
        for name in wb.sheetnames:
            ws = wb[name]
            if ws.max_row > 1 and ws.max_column > 1:
                sheet_name = name
                break
        ws = wb[sheet_name or wb.active.title]

        data = list(ws.values)
        if not data:
            st.error("❌ Excel file is empty.")
        else:
            # --- Prepare headers safely ---
            headers = [str(h).strip() if h else f"Unnamed_{i}" for i, h in enumerate(data[0])]
            
            # ✅ Deduplicate headers BEFORE DataFrame creation
            seen = {}
            new_headers = []
            for h in headers:
                if h in seen:
                    seen[h] += 1
                    new_headers.append(f"{h}_{seen[h]}")  # e.g. amount_2
                else:
                    seen[h] = 1
                    new_headers.append(h)

            rows = data[1:]
            raw_df = pd.DataFrame(rows, columns=new_headers)

            # --- Clean & normalize ---
            raw_df = clean_excel_headers(raw_df)
            raw_df = normalize_columns(raw_df)
            
            st.session_state.df = raw_df
            st.success("✅ Excel loaded successfully (duplicates handled).")
            st.caption(f"Loaded {len(raw_df)} rows and {len(raw_df.columns)} columns.")
            st.dataframe(raw_df.head(50), use_container_width=True)
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")

# ------------------------------------------------------------
# CHAT
# ------------------------------------------------------------
st.subheader("Chat")
if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask about invoices...")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, df_out = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if df_out is not None and not df_out.empty:
        st.dataframe(df_out, use_container_width=True)
