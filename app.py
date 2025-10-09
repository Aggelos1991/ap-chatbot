import io
import re
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
st.caption("Examples: 'open amounts', 'emails for open invoices', "
           "'group by vendor', 'due date between 2025-01-01 and 2025-03-31', "
           "'give me the open amounts emails separate with ; per language'")

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

        "document": "document",
        "invoice_number": "invoice_no",
        "invoice": "invoice_no",

        "open_amount_in_base_cur": "amount",
        "open_amount": "amount",
        "amount_in_eur": "amount",
        "amount_2": "amount",

        "currency": "currency",
        "due_date": "due_date",
        "issue_date": "issue_date",
        "due_month": "due_month",

        "payment_method_doc": "payment_method_descri",
        "payment_method_supplier": "payment_method_descri",

        "workflow_step": "workflow_step",
        "agreeded": "agreed",
        "agreed": "agreed",

        "email": "vendor_email",
        "e_mail": "vendor_email",
        "correo": "vendor_email",
        "country": "country"
    }
    df = df.rename(columns=lambda c: base_map.get(c, c))

    if "vendor_email" not in df.columns:
        for c in df.columns:
            c_plain = strip_accents(c)
            if any(x in c_plain.lower() for x in ["email", "correo", "mail", "ηλεκτρον", "διευθυν"]):
                df = df.rename(columns={c: "vendor_email"})
                break

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

    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    if "due_date" in df.columns:
        df["due_date_parsed"] = pd.to_datetime(df["due_date"], errors="coerce")
    if "agreed" in df.columns:
        df["agreed"] = pd.to_numeric(df["agreed"], errors="coerce").fillna(0).astype(int)
    else:
        df["agreed"] = 0

    # ---- NEW PROMPT ----
    if "open amounts emails" in ql:
        if "vendor_name" not in df.columns or "vendor_email" not in df.columns:
            return "⚠️ Missing 'vendor_name' or 'vendor_email' columns.", None

        if "country" in df.columns:
            df["country_norm"] = df["country"].astype(str).str.lower()
        else:
            df["country_norm"] = "other"

        df["lang"] = df["country_norm"].apply(
            lambda x: "ES" if "spain" in x or x.strip() in ["es", "esp", "españa"] else "EN"
        )

        df["vendor_email"] = df["vendor_email"].astype(str)
        grouped = (
            df.groupby(["lang", "vendor_name"], dropna=True)["vendor_email"]
            .apply(lambda x: "; ".join(sorted({e.strip() for e in x if e.strip()})))
            .reset_index()
        )

        es_df = grouped[grouped["lang"] == "ES"].drop(columns=["lang"])
        en_df = grouped[grouped["lang"] == "EN"].drop(columns=["lang"])

        st.write("🇪🇸 **Spanish Vendors (Spain)**")
        st.dataframe(es_df, use_container_width=True)

        st.write("🇬🇧 **English Vendors (Other Countries)**")
        st.dataframe(en_df, use_container_width=True)

        return "📧 Open amounts emails separated per language.", None

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

        ws = None
        for name in wb.sheetnames:
            w = wb[name]
            if w.max_row > 1 and w.max_column > 1:
                ws = w
                break
        if ws is None:
            ws = wb.active

        data = list(ws.values)
        if not data:
            st.error("❌ Excel is empty.")
        else:
            headers_raw = [str(h).strip() if h else f"Unnamed_{i}" for i, h in enumerate(data[0])]
            headers_dedup = dedupe_headers(headers_raw)
            df = pd.DataFrame(data[1:], columns=headers_dedup)

            df = clean_headers(df)
            df.columns = dedupe_headers(df.columns)
            df = normalize_columns(df)

            # =======================
            # APPLY FILTERS
            # =======================
            if "document" in df.columns:
                df = df[~df["document"].astype(str).str.contains("F&B", case=False, na=False)]

            if "type" in df.columns:
                df = df[df["type"].astype(str).str.upper() == "XPI"]

            if "payment_method_descri" in df.columns:
                df = df[~df["payment_method_descri"].astype(str).str.lower().isin(
                    ["downpayment", "direct debit", "cash", "credit card"]
                )]

            agreed_col = "agreed" if "agreed" in df.columns else "agreeded" if "agreeded" in df.columns else None
            if agreed_col:
                df = df[pd.to_numeric(df[agreed_col], errors="coerce").fillna(0) == 0]

            st.session_state.df = df
            st.success(f"✅ Excel loaded and filtered: {len(df)} rows | {len(df.columns)} cols.")
            st.dataframe(df.head(30), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

# =========================
# CHAT SECTION
# =========================
st.subheader("Chat")

if "history" not in st.session_state:
    st.session_state.history = []

if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask or add comment...")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, df_out = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if df_out is not None and not df_out.empty:
        st.dataframe(df_out, use_container_width=True)
