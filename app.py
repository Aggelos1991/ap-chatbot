import re
import pandas as pd
import streamlit as st
from datetime import date

# ----------------------------------------------------------
# Streamlit UI setup
# ----------------------------------------------------------
st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Try: 'open amount for vendor test', 'workflow step for vendor test', 'payment method for vendor test'")

# ----------------------------------------------------------
# Column normalization
# ----------------------------------------------------------
SYNONYMS = {
    "alternative_document": ["alternative document", "alt doc", "alt document", "alternative", "alt"],
    "vendor_name": ["vendor", "vendor name", "supplier", "supplier name"],
    "vendor_email": ["email", "vendor email", "supplier email", "correo", "mail"],
    "amount": ["amount", "total", "invoice amount", "importe", "value"],
    "currency": ["currency", "curr", "moneda"],
    "due_date": ["due date", "vencimiento", "fecha vencimiento"],
    "payment_date": ["payment date", "fecha pago", "paid date"],
    "agreed": ["agreed", "is agreed", "approved", "paid flag"],
    "workflow_step": ["workflow step", "workflow", "step", "process"],
    "payment_method": ["payment method", "doc", "method", "payment type"],
    "status": ["status", "state", "open?", "paid?"],
}
STANDARD_COLS = list(SYNONYMS.keys())

def _clean(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize headers to internal standard names."""
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
def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

def detect_invoice_ids(text: str):
    text = text.lower()
    found = re.findall(r"\b[a-z]{2,}[0-9]+(?:[-_/]?[0-9a-z]+)*\b", text)
    cleaned = [re.sub(r"[-_/]", "", f).strip() for f in found]
    return list(dict.fromkeys(cleaned))

def match_invoice(df: pd.DataFrame, inv: str):
    """Match invoice only using AlternativeDocument."""
    def norm(x): return re.sub(r"[-_/]", "", str(x).lower())
    target = norm(inv)
    if "alternative_document" not in df.columns:
        return pd.DataFrame()
    df["__alt_norm__"] = df["alternative_document"].astype(str).apply(norm)
    match = df[df["__alt_norm__"].str.contains(target, na=False)]
    return match.drop(columns="__alt_norm__", errors="ignore")

def unique_nonempty(series: pd.Series):
    vals = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(vals, key=str.lower)

# ----------------------------------------------------------
# Core Query Logic
# ----------------------------------------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower()
    working = df.copy()

    # Normalize
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["agreed"] = pd.to_numeric(working["agreed"], errors="coerce")

    # Vendor filter
    vendor_match = None
    for v in working["vendor_name"].dropna().unique():
        if v.lower() in ql:
            vendor_match = v
            break
    if vendor_match:
        working = working[working["vendor_name"].astype(str).str.lower() == vendor_match.lower()]

    # Invoice filter (AlternativeDocument only)
    invs = detect_invoice_ids(ql)
    if invs:
        rows = []
        for inv in invs:
            m = match_invoice(working, inv)
            if not m.empty:
                rows.append(m)
        if rows:
            working = pd.concat(rows).drop_duplicates()
        else:
            return f"❌ No invoices found for: {', '.join(invs)}", None

    # Workflow step
    if "workflow" in ql or "block" in ql:
        steps = unique_nonempty(working["workflow_step"])
        if not steps:
            return "❌ No workflow step found for this query.", None
        return f"🔄 Workflow step(s): {', '.join(steps)}", working

    # Payment method (DOC)
    if "payment method" in ql or "doc" in ql:
        pm = unique_nonempty(working["payment_method"])
        if not pm:
            return "❌ No payment method found for this query.", None
        return f"💳 Payment method(s): {', '.join(pm)}", working

    # Agreed logic: 1 = paid / 0 = open
    if "open" in ql or "unpaid" in ql:
        working = working[working["agreed"] == 0]
    elif "paid" in ql or "approved" in ql:
        working = working[working["agreed"] == 1]

    # Amount queries
    if "amount" in ql or "total" in ql:
        if working.empty:
            return "❌ No matching invoices found.", None
        total = working["amount"].sum()
        header = "💰 "
        if "open" in ql:
            header += "Total open amount"
        elif "paid" in ql:
            header += "Total paid amount"
        else:
            header += "Total amount"
        if vendor_match:
            header += f" for {vendor_match}"
        return f"{header}: **{fmt_money(total, working['currency'].iloc[0] if 'currency' in working else 'EUR')}**", working

    # Grouped totals by vendor
    if "amount by vendor" in ql or "vendor totals" in ql or "total by vendor" in ql:
        g = (
            working.groupby("vendor_name", dropna=True)["amount"].sum()
            .reset_index()
            .rename(columns={"amount": "total_amount"})
            .sort_values("total_amount", ascending=False)
        )
        if g.empty:
            return "No vendor totals found.", None
        g["total_amount"] = g["total_amount"].map(lambda x: f"{x:,.2f}")
        return "📊 Totals by vendor:", g

    if working.empty:
        return "No invoices match your query.", None

    return f"Found **{len(working)}** matching invoice(s).", working.reset_index(drop=True)

# ----------------------------------------------------------
# UI / File Upload
# ----------------------------------------------------------
st.sidebar.header("📦 Upload Excel")
st.sidebar.write("Columns: Alternative Document, Vendor Name, Amount, Agreed, Workflow Step, Payment Method (DOC).")

uploaded = st.file_uploader("Upload your Excel (.xlsx)", type=["xlsx"])

if "df" not in st.session_state:
    st.session_state.df = None

if uploaded:
    try:
        # --- SAFE EXCEL LOAD (handles duplicate column names) ---
        excel_data = pd.read_excel(uploaded, dtype=str, header=0)
        seen = {}
        new_columns = []
        renamed = {}
        for c in excel_data.columns:
            if c in seen:
                seen[c] += 1
                new_name = f"{c}_{seen[c]}"
                new_columns.append(new_name)
                renamed[c] = renamed.get(c, []) + [new_name]
            else:
                seen[c] = 0
                new_columns.append(c)
        excel_data.columns = new_columns

        # Normalize
        df = normalize_columns(excel_data)
        st.session_state.df = df

        st.success("✅ Excel loaded successfully.")
        st.dataframe(df.head(30), use_container_width=True)

        # Sidebar info for duplicates
        if renamed:
            st.sidebar.warning("⚠️ Duplicate headers were found and renamed automatically:")
            for base, dups in renamed.items():
                st.sidebar.write(f"- **{base}** → {', '.join(dups)}")

    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")

# ----------------------------------------------------------
# Chat section
# ----------------------------------------------------------
st.subheader("Chat")

if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask: e.g. 'open amount for vendor test', 'workflow step for vendor test', 'payment method for vendor test'")
if prompt:
    st.session_state.history.append(("user", prompt))
    st.chat_message("user").write(prompt)
    answer, result_df = run_query(prompt, st.session_state.df)
    st.session_state.history.append(("assistant", answer))
    st.chat_message("assistant").write(answer)
    if isinstance(result_df, pd.DataFrame) and not result_df.empty:
        st.dataframe(result_df, use_container_width=True)
        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download results as CSV", csv, file_name="results.csv", mime="text/csv")
