import re
import pandas as pd
import streamlit as st
from datetime import date

st.set_page_config(page_title="AP Chatbot (Excel)", page_icon="💼", layout="wide")
st.title("💬 Accounts Payable Chatbot — Excel-driven")
st.caption("Try: 'open over 560', 'emails for unpaid', "
           "'due before 2025-05-10', 'due smaller than 2025-05-10', "
           "'between 2024-10-01 and 2025-02-01', 'oldest unpaid invoice'")

# ---------------------------
# Column normalization
# ---------------------------
SYNONYMS = {
    "invoice_no":   ["invoice no","invoice number","invoice","inv","inv no","inv#","document no","doc no","document"],
    "vendor_name":  ["vendor","vendor name","supplier","supplier name","proveedor"],
    "vendor_email": ["email","vendor email","supplier email","correo","mail"],
    "status":       ["status","state","paid?","open?","payment status"],
    "amount":       ["amount","total","invoice amount","importe","value"],
    "currency":     ["currency","curr","moneda"],
    "due_date":     ["due date","vencimiento","fecha vencimiento"],
    "payment_date": ["payment date","fecha pago","paid date"],
    "po_number":    ["po","po number","purchase order"],
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
# Invoice IDs
# ---------------------------
def detect_invoice_ids(text: str):
    text = text.lower()
    candidates = re.findall(r"\b[a-z]{2,}[0-9]+[-/0-9a-z]*\b", text)
    ignore = {
        "paid","open","pending","invoice","invoices","inv","unpaid","status",
        "email","emails","mail","for","the","can","you","bring","vendor","vendors","supplier","suppliers",
        "amount","currency","due","payment","date","show","list","over","under","below","older","oldest","newest",
        "what","is","tell","give","please","find","greater","than","less","more","sum","total","before","after",
        "since","on","names","totals","by","top","smaller","bigger"
    }
    out, seen = [], set()
    for t in candidates:
        if t in ignore or t.isalpha():
            continue
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out

def find_best_invoice_match(df: pd.DataFrame, inv: str):
    if "invoice_no" not in df.columns:
        return pd.DataFrame()
    def norm(x): return re.sub(r"[-_\s]", "", str(x).strip().lower())
    tmp = df.copy()
    tmp["__inv_norm__"] = tmp["invoice_no"].astype(str).apply(norm)
    target = norm(inv)
    exact = tmp[tmp["__inv_norm__"] == target]
    if not exact.empty:
        return exact.drop(columns="__inv_norm__", errors="ignore")
    like = tmp[tmp["__inv_norm__"].str.contains(target, na=False)]
    return like.drop(columns="__inv_norm__", errors="ignore")

# ---------------------------
# Dates
# ---------------------------
def parse_user_date(s: str):
    for dayfirst in (True, False):
        try:
            return pd.to_datetime(s, dayfirst=dayfirst, errors="raise")
        except Exception:
            pass
    return pd.NaT

DATE_RX = r"(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"

def extract_date_intent(ql: str):
    """
    Returns dict like:
      {"mode": "before"/"after"/"on"/"between", "d1": Timestamp, "d2": Timestamp|None}
    or None if no clear intent.
    """
    ql = ql.lower().strip()

    # between X and Y
    m_between = re.search(rf"between\s+{DATE_RX}\s+and\s+{DATE_RX}", ql)
    if m_between:
        d1 = parse_user_date(m_between.group(1))
        d2 = parse_user_date(m_between.group(2))
        if pd.notna(d1) and pd.notna(d2):
            if d2 < d1: d1, d2 = d2, d1
            return {"mode":"between","d1":d1,"d2":d2}

    # single date capture
    m = re.search(DATE_RX, ql)
    if not m:
        return None
    d = parse_user_date(m.group(1))
    if pd.isna(d):
        return None

    before_words = ("before","earlier than","smaller than","less than","until","by","on or before","<","<=")
    after_words  = ("after","later than","greater than","bigger than","from","since","on or after",">",">=")
    on_words     = ("on",)

    if any(w in ql for w in before_words):
        return {"mode":"before","d1":d,"d2":None}
    if any(w in ql for w in after_words):
        return {"mode":"after","d1":d,"d2":None}
    if any(re.search(rf"\b{w}\b", ql) for w in on_words):
        return {"mode":"on","d1":d,"d2":None}

    # If they said "due smaller than ..." we may not catch via word boundary; handled by before_words above.
    # Fallback: if they used "due" with a date, default to "on"
    if "due" in ql:
        return {"mode":"on","d1":d,"d2":None}
    return None

# ---------------------------
# Helpers
# ---------------------------
def extract_top_n_old_new(ql: str):
    m = re.search(r"(?:top\s*)?(\d+)\s*(oldest|newest)", ql)
    if not m:
        if "oldest" in ql or "newest" in ql: return 10
        return None
    return int(m.group(1))

def extract_top_n_generic(ql: str):
    m = re.search(r"(?:top|first)\s*(\d+)", ql)
    return int(m.group(1)) if m else None

def wants_oldest(ql: str): return any(w in ql for w in ("oldest","earliest","older"))
def wants_newest(ql: str): return any(w in ql for w in ("newest","latest","newer"))

def fmt_money(amount, currency):
    try:
        a = float(str(amount).replace(",", "").strip())
        cur = currency if isinstance(currency, str) and currency else "EUR"
        return f"{a:,.2f} {cur}"
    except Exception:
        return str(amount)

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

# ---------------------------
# Core
# ---------------------------
def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "Please upload an Excel first.", None

    ql = q.lower()

    # Working view + normalized helpers
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")
    working["status_norm"] = working["status"].astype(str).str.lower().str.strip()

    # Status filters
    if any(k in ql for k in ("open","unpaid","pending")):
        working = working[working["status_norm"].str.contains("open|unpaid|pending", na=False)]
    if "paid" in ql and not any(k in ql for k in ("unpaid","not paid","open","pending")):
        working = working[working["status_norm"].str.contains("paid", na=False)]

    # Vendor name filter: "for sani" / "for technogym"
    vm = re.search(r"\bfor\s+([a-z0-9 ._-]+)", ql)
    if vm:
        needle = vm.group(1).strip()
        working = working[working["vendor_name"].astype(str).str.contains(re.escape(needle), case=False, na=False)]

    # Amount filters (numeric)
    m_over = re.search(r"(over|above|greater than|>=|more than|bigger than)\s*([0-9][0-9,\.]*)", ql)
    if m_over:
        val = float(m_over.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    m_under = re.search(r"(under|below|less than|<=|smaller than)\s*([0-9][0-9,\.]*)", ql)
    if m_under:
        val2 = float(m_under.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Date intent on due date
    di = extract_date_intent(ql)
    if di is not None:
        d1, d2, mode = di["d1"], di["d2"], di["mode"]
        if mode == "before":
            working = working[working["due_date_parsed"] < d1]
        elif mode == "after":
            working = working[working["due_date_parsed"] > d1]
        elif mode == "on":
            working = working[working["due_date_parsed"].dt.date == d1.date()]
        elif mode == "between":
            working = working[(working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)]

    # Oldest / newest by due date
    if "due_date_parsed" in working.columns:
        n_old_new = extract_top_n_old_new(ql)
        if wants_oldest(ql):
            ws = working.sort_values("due_date_parsed", ascending=True)
            if "oldest invoice" in ql or "the oldest" in ql:
                if ws.empty: return "No invoices found with a valid due date.", None
                r = ws.iloc[0]
                return (f"📄 Oldest invoice: **{r.get('invoice_no','-')}** — **{r.get('vendor_name','-')}**, "
                        f"due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                        f"status **{r.get('status','-')}**." , ws.head(1).reset_index(drop=True))
            if n_old_new: working = ws.head(n_old_new)
        elif wants_newest(ql):
            ws = working.sort_values("due_date_parsed", ascending=False)
            if "newest invoice" in ql or "latest invoice" in ql:
                if ws.empty: return "No invoices found with a valid due date.", None
                r = ws.iloc[0]
                return (f"📄 Newest invoice: **{r.get('invoice_no','-')}** — **{r.get('vendor_name','-')}**, "
                        f"due **{r.get('due_date','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                        f"status **{r.get('status','-')}**." , ws.head(1).reset_index(drop=True))
            if n_old_new: working = ws.head(n_old_new)

    # Specific invoice questions (keep after filters if they also restrict)
    inv_ids = detect_invoice_ids(ql)
    if inv_ids:
        rows, msg_parts = [], []
        for inv in inv_ids:
            found = find_best_invoice_match(working, inv)
            if found.empty:
                msg_parts.append(f"❓ Could not find invoice **{inv}**.")
            else:
                rows.append(found.iloc[0])
                r = found.iloc[0]
                msg_parts.append(
                    f"Invoice **{r.get('invoice_no','-')}** — vendor **{r.get('vendor_name','-')}**, "
                    f"status **{r.get('status','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                    f"due **{r.get('due_date','-')}**."
                )
        if rows:
            return "\n\n".join(msg_parts), pd.DataFrame(rows)
        return "\n\n".join(msg_parts), None

    # Field-specific outputs
    wants_emails = "email" in ql or "emails" in ql
    wants_names  = ("vendor" in ql or "supplier" in ql) and ("name" in ql or "names" in ql) and not wants_emails
    wants_amounts = "amounts" in ql or "amount list" in ql

    if wants_emails:
        emails = unique_nonempty(working["vendor_email"])
        if not emails: return "No vendor emails found for this query.", None
        return f"📧 **{len(emails)} vendor email(s):**\n\n" + "; ".join(emails), pd.DataFrame({"vendor_email":emails})

    if wants_names:
        names = unique_nonempty(working["vendor_name"])
        if not names: return "No vendor names match your filters.", None
        return f"👤 **{len(names)} vendor(s):**\n\n- " + "\n- ".join(names), pd.DataFrame({"vendor_name":names})

    if wants_amounts:
        if working.empty: return "No invoices match your filters.", None
        out = working[["invoice_no","vendor_name","amount","currency"]].copy()
        out["amount_fmt"] = out.apply(lambda r: fmt_money(r["amount"], r["currency"]), axis=1)
        return "💵 Amounts for matching invoices:", out[["invoice_no","vendor_name","amount_fmt"]]

    # Grouped totals by vendor
    if ("vendor" in ql or "supplier" in ql) and any(k in ql for k in ("amounts","amount by","amount per","totals","total by","open amounts")):
        g = (working.groupby("vendor_name", dropna=True)["amount"].sum()
             .reset_index().rename(columns={"amount":"total_amount"})
             .sort_values("total_amount", ascending=False))
        n = extract_top_n_generic(ql)
        if n: g = g.head(n)
        if g.empty: return "No vendor totals for this query.", None
        g["total_amount"] = g["total_amount"].map(lambda x: f"{x:,.2f}")
        return f"📊 Totals by vendor ({'top '+str(n) if n else 'all'}):", g

    # Sum total
    if "sum" in ql or "total" in ql:
        total = pd.to_numeric(working["amount"], errors="coerce").sum()
        return f"💰 Total amount: **{total:,.2f}**", working.reset_index(drop=True)

    if working.empty:
        return "No invoices match your query.", None

    # Default list of matches
    return f"Found **{len(working)}** invoice(s) matching your query.", working.reset_index(drop=True)

# ---------------------------
# UI
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
if st.button("🔄 Restart Chat"):
    st.session_state.history = []
    st.rerun()

if "history" not in st.session_state:
    st.session_state.history = []

for role, msg in st.session_state.history:
    st.chat_message(role).write(msg)

prompt = st.chat_input("Ask about invoices: e.g., 'emails for unpaid', 'open over 1000', "
                       "'vendor names for open', 'due smaller than 2025-05-10', 'between 2024-10-01 and 2025-02-01'")
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